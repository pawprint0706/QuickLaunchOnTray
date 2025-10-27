using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Threading.Tasks;

namespace QuickLaunchOnTray
{
    static class Program
    {
        /// <summary>
        /// 해당 응용 프로그램의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new TrayApplicationContext());
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("프로그램 실행 중 오류가 발생했습니다: {0}", ex.Message),
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                Environment.Exit(1);
            }
        }

        // 프로그램 정보를 담는 클래스
        public class ProgramItem
        {
            public string Name { get; set; }  // ini 파일의 key 값 또는 단순 경로인 경우 파일명(확장자 제외)
            public string Path { get; set; }  // 프로그램 경로
        }

        // 시스템 언어가 한국어인지 확인하는 메서드
        private static bool IsKoreanSystem()
        {
            return CultureInfo.CurrentUICulture.TwoLetterISOLanguageName.Equals("ko", StringComparison.OrdinalIgnoreCase);
        }

        // 언어에 따른 텍스트 반환 메서드
        private static string GetLocalizedText(string koreanText, string englishText)
        {
            return IsKoreanSystem() ? koreanText : englishText;
        }

        /// <summary>
        /// 폼 없이 트레이 아이콘들을 관리하는 ApplicationContext 클래스
        /// </summary>
        public class TrayApplicationContext : ApplicationContext
        {
            [DllImport("user32.dll")]
            private static extern bool GetCursorPos(out POINT lpPoint);

            [DllImport("shell32.dll", CharSet = CharSet.Auto)]
            private static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            private static extern bool DestroyIcon(IntPtr handle);

            [DllImport("user32.dll")]
            private static extern bool SetForegroundWindow(IntPtr hWnd);

            [StructLayout(LayoutKind.Sequential)]
            private struct POINT
            {
                public int X;
                public int Y;
            }

            // 생성한 트레이 아이콘들을 보관할 리스트
            private readonly List<NotifyIcon> trayIcons = new List<NotifyIcon>();
            private Icon folderIcon;
            private Bitmap folderBitmap;
            private Bitmap defaultFileBitmap;
            private Form hiddenForm;
            private readonly ConcurrentDictionary<string, Bitmap> iconCache = new ConcurrentDictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);
            private readonly ConcurrentDictionary<string, FolderCacheEntry> folderCache = new ConcurrentDictionary<string, FolderCacheEntry>(StringComparer.OrdinalIgnoreCase);
            private readonly TimeSpan cacheDuration = TimeSpan.FromSeconds(30);
            private readonly object loadingTag = new object();

            public TrayApplicationContext()
            {
                // 폴더 아이콘 초기화
                try
                {
                    string systemFolder = Environment.GetFolderPath(Environment.SpecialFolder.System);
                    IntPtr hIcon = ExtractIcon(IntPtr.Zero, Path.Combine(systemFolder, "shell32.dll"), 3); // 3은 폴더 아이콘의 인덱스
                    if (hIcon != IntPtr.Zero)
                    {
                        folderIcon = Icon.FromHandle(hIcon);
                    }
                    else
                    {
                        folderIcon = SystemIcons.Application;
                    }
                }
                catch
                {
                    folderIcon = SystemIcons.Application;
                }

                folderBitmap = folderIcon != null ? folderIcon.ToBitmap() : null;
                defaultFileBitmap = SystemIcons.Application.ToBitmap();

                // 실행 파일과 같은 폴더의 config.ini 파일 경로 지정
                string iniPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini");
                if (!File.Exists(iniPath))
                {
                    MessageBox.Show(GetLocalizedText(
                        "'config.ini' 파일을 찾을 수 없습니다:\n" + iniPath,
                        "Could not find 'config.ini' file:\n" + iniPath));
                    Environment.Exit(0);
                    return;
                }

                // ini 파일에서 프로그램 정보(이름 및 경로) 읽어오기
                List<ProgramItem> programItems = LoadProgramItems(iniPath);
                if (programItems.Count == 0)
                {
                    MessageBox.Show(GetLocalizedText(
                        "'config.ini' 파일에 실행할 프로그램 정보가 없습니다.",
                        "There is no information about the program to run in the 'config.ini' file."));
                    Environment.Exit(0);
                    return;
                }

                // 각 프로그램 정보에 대해 트레이 아이콘 생성
                foreach (ProgramItem item in programItems)
                {
                    CreateTrayIconForProgram(item);
                }

                hiddenForm = new Form();
                hiddenForm.FormBorderStyle = FormBorderStyle.None;
                hiddenForm.ShowInTaskbar = false;
                hiddenForm.StartPosition = FormStartPosition.Manual;
                hiddenForm.Size = new Size(1, 1);
                hiddenForm.Location = new Point(-2000, -2000); // 화면 밖
                hiddenForm.Opacity = 0;
                hiddenForm.Show();
                hiddenForm.Hide();
            }

            /// <summary>
            /// ini 파일에서 [Programs] 섹션 내의 프로그램 정보를 읽어 리스트로 반환합니다.
            /// key=value 형식인 경우 key를 이름으로 사용하며, 단순 경로인 경우 파일명(확장자 제외)을 이름으로 사용합니다.
            /// </summary>
            private List<ProgramItem> LoadProgramItems(string iniPath)
            {
                List<ProgramItem> items = new List<ProgramItem>();
                bool inProgramsSection = false;

                foreach (string line in File.ReadAllLines(iniPath))
                {
                    string trimmed = line.Trim();

                    // 빈 줄이나 주석은 무시
                    if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith(";") || trimmed.StartsWith("#"))
                        continue;

                    // 섹션 구분 [Programs] 등
                    if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                    {
                        string section = trimmed.Substring(1, trimmed.Length - 2);
                        inProgramsSection = section.Equals("Programs", StringComparison.OrdinalIgnoreCase);
                        continue;
                    }

                    if (inProgramsSection)
                    {
                        // key=value 형식인 경우 key는 이름, value는 경로로 사용
                        if (trimmed.Contains("="))
                        {
                            int idx = trimmed.IndexOf('=');
                            string key = trimmed.Substring(0, idx).Trim();
                            string value = trimmed.Substring(idx + 1).Trim();

                            if (!string.IsNullOrEmpty(value))
                                items.Add(new ProgramItem { Name = key, Path = value });
                        }
                        else
                        {
                            // 단순 경로인 경우 파일명(확장자 제외)을 이름으로 사용
                            string fileName = Path.GetFileNameWithoutExtension(trimmed);
                            items.Add(new ProgramItem { Name = fileName, Path = trimmed });
                        }
                    }
                }
                return items;
            }

            /// <summary>
            /// 지정한 프로그램 정보를 바탕으로 트레이 아이콘을 생성 및 등록합니다.
            /// </summary>
            private void CreateTrayIconForProgram(ProgramItem item)
            {
                // 아이콘 추출: 프로그램 파일이 존재하면 해당 아이콘을, 폴더인 경우 폴더 아이콘을, 그렇지 않으면 기본 아이콘 사용
                Icon icon;
                try
                {
                    if (Directory.Exists(item.Path))
                        icon = folderIcon;
                    else if (File.Exists(item.Path))
                        icon = Icon.ExtractAssociatedIcon(item.Path);
                    else
                        icon = SystemIcons.Application;
                }
                catch
                {
                    icon = SystemIcons.Application;
                }

                NotifyIcon notifyIcon = new NotifyIcon
                {
                    Icon = icon,
                    Text = item.Name,
                    Visible = true
                };

                // 컨텍스트 메뉴 생성
                ContextMenuStrip menu = new ContextMenuStrip();

                // "Run this program" 또는 "Open folder" 메뉴 항목
                string menuText = Directory.Exists(item.Path) 
                    ? GetLocalizedText("폴더 열기", "Open folder")
                    : GetLocalizedText("프로그램 실행", "Run this program");
                ToolStripMenuItem runItem = new ToolStripMenuItem(menuText);
                runItem.Click += (s, e) =>
                {
                    if (Directory.Exists(item.Path))
                    {
                        Process.Start("explorer.exe", item.Path);
                    }
                    else
                    {
                        RunProgram(item.Path);
                    }
                };
                menu.Items.Add(runItem);

                // "Terminate 'QuickLaunchOnTray'" 메뉴 항목
                ToolStripMenuItem exitItem = new ToolStripMenuItem(GetLocalizedText(
                    "'QuickLaunchOnTray' 종료",
                    "Terminate 'QuickLaunchOnTray'"));
                exitItem.Click += (s, e) =>
                {
                    DialogResult result = MessageBox.Show(
                        GetLocalizedText(
                            "'QuickLaunchOnTray'를 종료하시겠습니까? 모든 빠른 실행 아이콘이 사라집니다.",
                            "Are you sure you want to terminate 'QuickLaunchOnTray'? All quick launch icons will disappear."),
                        GetLocalizedText("확인", "Confirm"),
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        ExitThread();
                    }
                };
                menu.Items.Add(exitItem);

                notifyIcon.ContextMenuStrip = menu;

                // 클릭 이벤트 처리
                notifyIcon.MouseClick += (s, e) =>
                {
                    if (e.Button == MouseButtons.Left)
                    {
                        if (Directory.Exists(item.Path))
                        {
                            ShowFolderContextMenuWithFiles(item.Path);
                        }
                        else
                        {
                            RunProgram(item.Path);
                        }
                    }
                };

                trayIcons.Add(notifyIcon);
            }

            private void ShowFolderContextMenuWithFiles(string folderPath)
            {
                try
                {
                    if (!Directory.Exists(folderPath))
                    {
                        throw new DirectoryNotFoundException(folderPath);
                    }

                    ContextMenuStrip folderMenu = new ContextMenuStrip();
                    AddLoadingPlaceholder(folderMenu.Items);
                    LoadFolderMenuItemsAsync(folderPath, folderMenu.Items, includeRootExtras: true, ownerMenuItem: null);

                    POINT cursorPos;
                    GetCursorPos(out cursorPos);
                    folderMenu.Show(hiddenForm, hiddenForm.PointToClient(new Point(cursorPos.X, cursorPos.Y)));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(GetLocalizedText(
                        "폴더 메뉴 표시 중 오류 발생: " + ex.Message,
                        "Error showing folder menu: " + ex.Message));
                }
            }

            private void LoadFolderMenuItemsAsync(string folderPath, ToolStripItemCollection items, bool includeRootExtras, ToolStripMenuItem ownerMenuItem)
            {
                Task.Run(() =>
                {
                    try
                    {
                        var entries = GetFolderEntries(folderPath);
                        if (hiddenForm == null || hiddenForm.IsDisposed)
                        {
                            return;
                        }

                        hiddenForm.BeginInvoke(new Action(() =>
                        {
                            if ((ownerMenuItem != null && ownerMenuItem.IsDisposed) || items == null)
                            {
                                return;
                            }

                            ToolStripDropDown dropDown = items.Owner as ToolStripDropDown;
                            if (dropDown != null && dropDown.IsDisposed)
                            {
                                return;
                            }

                            ApplyFolderEntries(folderPath, items, entries, includeRootExtras, ownerMenuItem);
                        }));
                    }
                    catch (Exception ex)
                    {
                        if (hiddenForm == null || hiddenForm.IsDisposed)
                        {
                            return;
                        }

                        hiddenForm.BeginInvoke(new Action(() =>
                        {
                            if ((ownerMenuItem != null && ownerMenuItem.IsDisposed) || items == null)
                            {
                                return;
                            }

                            ToolStripDropDown dropDown = items.Owner as ToolStripDropDown;
                            if (dropDown != null && dropDown.IsDisposed)
                            {
                                return;
                            }

                            items.Clear();
                            items.Add(new ToolStripMenuItem(GetLocalizedText("오류", "Error"))
                            {
                                Enabled = false
                            });

                            if (includeRootExtras)
                            {
                                AddRootFooterItems(folderPath, items);
                            }
                            else
                            {
                                AddFolderFooterItems(folderPath, items);
                            }

                            FolderMenuMetadata metadata = null;
                            if (ownerMenuItem != null)
                            {
                                metadata = ownerMenuItem.Tag as FolderMenuMetadata;
                            }
                            if (metadata != null)
                            {
                                metadata.IsLoading = false;
                            }

                            MessageBox.Show(GetLocalizedText(
                                "폴더 메뉴 구성 중 오류 발생: " + ex.Message,
                                "Error building folder menu: " + ex.Message));
                        }));
                    }
                });
            }

            private void ApplyFolderEntries(string folderPath, ToolStripItemCollection items, IReadOnlyList<FolderItemInfo> entries, bool includeRootExtras, ToolStripMenuItem ownerMenuItem)
            {
                items.Clear();

                int addedCount = 0;
                foreach (var folder in entries.Where(e => e.IsFolder))
                {
                    items.Add(CreateFolderMenuItem(folder));
                    addedCount++;
                }

                foreach (var file in entries.Where(e => !e.IsFolder))
                {
                    items.Add(CreateFileMenuItem(file));
                    addedCount++;
                }

                if (addedCount == 0)
                {
                    AddEmptyPlaceholder(items);
                }

                if (includeRootExtras)
                {
                    AddRootFooterItems(folderPath, items);
                }
                else
                {
                    AddFolderFooterItems(folderPath, items);
                }

                if (ownerMenuItem != null)
                {
                    FolderMenuMetadata metadata = ownerMenuItem.Tag as FolderMenuMetadata;
                    if (metadata != null)
                    {
                        metadata.IsLoading = false;
                        metadata.HasEverLoaded = true;
                        metadata.LastLoaded = DateTime.UtcNow;
                    }
                }
            }

            private void AddRootFooterItems(string folderPath, ToolStripItemCollection items)
            {
                if (items.Count > 0)
                {
                    items.Add(new ToolStripSeparator());
                }

                var openInExplorer = new ToolStripMenuItem(GetLocalizedText("탐색기에서 열기", "Open in Explorer"));
                openInExplorer.Click += (s, e) => { Process.Start("explorer.exe", folderPath); };
                items.Add(openInExplorer);

                var closeMenu = new ToolStripMenuItem(GetLocalizedText("닫기", "Close"));
                closeMenu.Click += (s, e) =>
                {
                    ToolStripMenuItem menuItem = s as ToolStripMenuItem;
                    if (menuItem != null)
                    {
                        ToolStrip currentParent = menuItem.GetCurrentParent();
                        if (currentParent != null)
                        {
                            currentParent.Close();
                        }
                    }
                };
                items.Add(closeMenu);
            }

            private void AddFolderFooterItems(string folderPath, ToolStripItemCollection items)
            {
                if (items.Count > 0)
                {
                    items.Add(new ToolStripSeparator());
                }

                var openInExplorer = new ToolStripMenuItem(GetLocalizedText("탐색기에서 열기", "Open in Explorer"));
                openInExplorer.Click += (s, e) => { Process.Start("explorer.exe", folderPath); };
                items.Add(openInExplorer);
            }

            private void AddLoadingPlaceholder(ToolStripItemCollection items)
            {
                items.Clear();
                items.Add(new ToolStripMenuItem(GetLocalizedText("불러오는 중...", "Loading..."))
                {
                    Enabled = false,
                    Tag = loadingTag
                });
            }

            private void AddEmptyPlaceholder(ToolStripItemCollection items)
            {
                items.Add(new ToolStripMenuItem(GetLocalizedText("표시할 항목이 없습니다", "No items available"))
                {
                    Enabled = false
                });
            }

            private ToolStripMenuItem CreateFolderMenuItem(FolderItemInfo folder)
            {
                var folderItem = new ToolStripMenuItem(folder.Name, folder.Icon ?? folderBitmap);
                var metadata = new FolderMenuMetadata
                {
                    FolderPath = folder.Path,
                    SupportsLoading = folder.HasChildren
                };
                folderItem.Tag = metadata;
                folderItem.DropDownOpening += FolderItem_DropDownOpening;

                if (folder.HasChildren)
                {
                    AddLoadingPlaceholder(folderItem.DropDownItems);
                }
                else
                {
                    metadata.HasEverLoaded = true;
                    metadata.LastLoaded = DateTime.UtcNow;
                    AddEmptyPlaceholder(folderItem.DropDownItems);
                    AddFolderFooterItems(folder.Path, folderItem.DropDownItems);
                }

                return folderItem;
            }

            private ToolStripMenuItem CreateFileMenuItem(FolderItemInfo file)
            {
                var fileItem = new ToolStripMenuItem(file.Name, file.Icon);
                fileItem.Click += (s, e) => { RunProgram(file.Path); };
                return fileItem;
            }

            private void FolderItem_DropDownOpening(object sender, EventArgs e)
            {
                ToolStripMenuItem folderItem = sender as ToolStripMenuItem;
                if (folderItem == null)
                {
                    return;
                }

                FolderMenuMetadata metadata = folderItem.Tag as FolderMenuMetadata;
                if (metadata == null)
                {
                    return;
                }

                if (!metadata.SupportsLoading)
                {
                    return;
                }

                if (metadata.IsLoading)
                {
                    return;
                }

                bool hasPlaceholder = folderItem.DropDownItems.Count == 1 && ReferenceEquals(folderItem.DropDownItems[0].Tag, loadingTag);
                bool shouldReload = !metadata.HasEverLoaded || DateTime.UtcNow - metadata.LastLoaded > cacheDuration || hasPlaceholder;

                if (!shouldReload)
                {
                    return;
                }

                metadata.IsLoading = true;
                AddLoadingPlaceholder(folderItem.DropDownItems);
                LoadFolderMenuItemsAsync(metadata.FolderPath, folderItem.DropDownItems, false, folderItem);
            }

            private IReadOnlyList<FolderItemInfo> GetFolderEntries(string folderPath)
            {
                var now = DateTime.UtcNow;
                if (folderCache.TryGetValue(folderPath, out var cacheEntry))
                {
                    if (now - cacheEntry.CachedAt <= cacheDuration)
                    {
                        return cacheEntry.Items;
                    }
                }

                var items = GetFolderItemsFromDisk(folderPath);
                var entry = new FolderCacheEntry
                {
                    CachedAt = now,
                    Items = items
                };
                folderCache[folderPath] = entry;
                return entry.Items;
            }

            private IReadOnlyList<FolderItemInfo> GetFolderItemsFromDisk(string folderPath)
            {
                var items = new List<FolderItemInfo>();

                foreach (var subFolder in Directory.EnumerateDirectories(folderPath))
                {
                    bool hasChildren = false;
                    try
                    {
                        using (var enumerator = Directory.EnumerateFileSystemEntries(subFolder).GetEnumerator())
                        {
                            hasChildren = enumerator.MoveNext();
                        }
                    }
                    catch
                    {
                        hasChildren = false;
                    }

                    items.Add(new FolderItemInfo
                    {
                        Name = Path.GetFileName(subFolder),
                        Path = subFolder,
                        IsFolder = true,
                        HasChildren = hasChildren,
                        Icon = folderBitmap
                    });
                }

                foreach (var file in Directory.EnumerateFiles(folderPath))
                {
                    items.Add(new FolderItemInfo
                    {
                        Name = Path.GetFileName(file),
                        Path = file,
                        IsFolder = false,
                        HasChildren = false,
                        Icon = GetIconBitmap(file)
                    });
                }

                items.Sort((x, y) =>
                {
                    int typeCompare = y.IsFolder.CompareTo(x.IsFolder);
                    if (typeCompare != 0)
                    {
                        return typeCompare;
                    }

                    return string.Compare(x.Name, y.Name, StringComparison.CurrentCultureIgnoreCase);
                });
                return items;
            }

            private class FolderItemInfo
            {
                public string Name { get; set; }
                public string Path { get; set; }
                public bool IsFolder { get; set; }
                public bool HasChildren { get; set; }
                public Bitmap Icon { get; set; }
            }

            private class FolderCacheEntry
            {
                public DateTime CachedAt { get; set; }
                public IReadOnlyList<FolderItemInfo> Items { get; set; }
            }

            private class FolderMenuMetadata
            {
                public string FolderPath { get; set; }
                public bool SupportsLoading { get; set; }
                public bool HasEverLoaded { get; set; }
                public DateTime LastLoaded { get; set; }
                public bool IsLoading { get; set; }
            }

            private Bitmap GetIconBitmap(string filePath)
            {
                if (string.IsNullOrEmpty(filePath))
                {
                    return defaultFileBitmap;
                }

                string extension = Path.GetExtension(filePath) ?? string.Empty;
                bool useFullPath = extension.Equals(".exe", StringComparison.OrdinalIgnoreCase) || extension.Equals(".lnk", StringComparison.OrdinalIgnoreCase);
                string cacheKey = useFullPath ? filePath : extension;

                return iconCache.GetOrAdd(cacheKey, key =>
                {
                    try
                    {
                        if (File.Exists(filePath))
                        {
                            using (var icon = Icon.ExtractAssociatedIcon(filePath))
                            {
                                if (icon != null)
                                {
                                    var bitmap = icon.ToBitmap();
                                    if (bitmap != null)
                                    {
                                        return bitmap;
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        // ignore and fallback
                    }

                    return defaultFileBitmap;
                });
            }

            /// <summary>
            /// 지정한 경로의 프로그램을 실행합니다.
            /// </summary>
            private void RunProgram(string programPath)
            {
                try
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = programPath,
                        UseShellExecute = true
                    };
                    Process.Start(startInfo);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(GetLocalizedText(
                        "프로그램 실행 실패:\n" + ex.Message,
                        "Run this program failed:\n" + ex.Message));
                }
            }

            /// <summary>
            /// ApplicationContext 종료 시 모든 트레이 아이콘을 정리합니다.
            /// </summary>
            protected override void ExitThreadCore()
            {
                // 모든 NotifyIcon을 숨기고 Dispose
                foreach (var icon in trayIcons)
                {
                    icon.Visible = false;
                    icon.Dispose();
                }

                var disposedBitmaps = new HashSet<Bitmap>();
                foreach (var bitmap in iconCache.Values)
                {
                    if (bitmap == null)
                        continue;
                    if (ReferenceEquals(bitmap, defaultFileBitmap) || ReferenceEquals(bitmap, folderBitmap))
                        continue;
                    if (disposedBitmaps.Add(bitmap))
                    {
                        bitmap.Dispose();
                    }
                }
                iconCache.Clear();
                folderCache.Clear();

                if (folderBitmap != null)
                {
                    folderBitmap.Dispose();
                }

                if (defaultFileBitmap != null)
                {
                    defaultFileBitmap.Dispose();
                }

                // 폴더 아이콘 해제
                if (folderIcon != null)
                {
                    folderIcon.Dispose();
                }

                if (hiddenForm != null) hiddenForm.Dispose();

                base.ExitThreadCore();
            }
        }
    }
}

