using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using QuickLaunchOnTray.Models;
using QuickLaunchOnTray.Services;

namespace QuickLaunchOnTray
{
    public class TrayApplicationContext : ApplicationContext, IDisposable
    {
        /// <summary>NotifyIcon.Text 상한. 이 값을 넘기면 ArgumentOutOfRangeException이 발생한다.</summary>
        private const int MaxTooltipLength = 63;

        /// <summary>설정된 폴더의 목록/아이콘을 미리 읽기 시작할 때까지의 지연(로그온 직후 디스크 경합 완화).</summary>
        private const int WarmupDelayMilliseconds = 2000;

        /// <summary>메뉴 표시 후 전경 상태를 넘겨주기 위해 보내는 빈 메시지.</summary>
        private const uint WmNull = 0x0000;

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        private readonly List<NotifyIcon> _trayIcons = new List<NotifyIcon>();
        private readonly ConfigurationService _configService;
        private readonly LocalizationService _localizationService;
        private readonly IconImageCache _iconImages;
        private readonly FolderIndexCache _folderIndexCache;
        private readonly TimeSpan _cacheDuration = TimeSpan.FromSeconds(30);
        private readonly object _loadingTag = new object();

        private readonly Icon _folderIcon;
        private readonly Bitmap _folderBitmap;
        private readonly Icon _defaultFileIcon;
        private readonly Bitmap _defaultFileBitmap;
        private readonly Form _hiddenForm;
        private System.Threading.Timer _warmupTimer;
        private bool _disposed;

        /// <summary>
        /// config.ini의 각 항목에 대한 트레이 아이콘을 등록한다.
        /// 항목 단위 실패는 경고로 모아 알리고, 등록된 아이콘이 하나도 없으면 예외를 던진다.
        /// </summary>
        public TrayApplicationContext()
        {
            _configService = new ConfigurationService();
            _localizationService = LocalizationService.Instance;

            _folderIcon = IconLoader.LoadShellIcon(IconLoader.FolderIconIndex) ?? IconLoader.CreateDefaultIcon();
            _folderBitmap = _folderIcon.ToBitmap();
            _defaultFileIcon = IconLoader.CreateDefaultIcon();
            _defaultFileBitmap = _defaultFileIcon.ToBitmap();

            _iconImages = new IconImageCache();
            _folderIndexCache = new FolderIndexCache(_iconImages);

            var loadResult = _configService.LoadProgramItems();
            var warnings = new List<string>(loadResult.Warnings);

            foreach (var item in loadResult.Items)
            {
                try
                {
                    CreateTrayIconForProgram(item);
                }
                catch (Exception ex)
                {
                    warnings.Add(_localizationService.GetString("EntryCreateSkipped", item.Name, ex.Message));
                }
            }

            if (_trayIcons.Count == 0)
            {
                throw new InvalidOperationException(
                    warnings.Count > 0
                        ? string.Join(Environment.NewLine, warnings)
                        : _localizationService.GetString("NoProgramInfo"));
            }

            _hiddenForm = CreateHiddenForm();

            if (warnings.Count > 0)
            {
                MessageBox.Show(
                    string.Join(Environment.NewLine, warnings),
                    _localizationService.GetString("Warning"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }

            StartWarmup(loadResult.Items.Where(item => Directory.Exists(item.Path)).Select(item => item.Path).ToList());
        }

        /// <summary>설정된 폴더들의 항목 목록과 아이콘을 백그라운드에서 미리 읽어 첫 클릭을 빠르게 한다.</summary>
        private void StartWarmup(IReadOnlyList<string> folderPaths)
        {
            if (folderPaths == null || folderPaths.Count == 0)
            {
                return;
            }

            _warmupTimer = new System.Threading.Timer(
                state =>
                {
                    try
                    {
                        foreach (string folderPath in folderPaths)
                        {
                            _folderIndexCache.Warm(folderPath);
                        }

                        CacheStorage.PruneAll();
                    }
                    catch
                    {
                        // 예열과 정리는 실패해도 무시한다.
                    }
                },
                null,
                WarmupDelayMilliseconds,
                System.Threading.Timeout.Infinite);
        }

        /// <summary>
        /// 폴더 메뉴의 owner 겸 BeginInvoke 대상으로 쓰는 보이지 않는 1x1 폼.
        /// </summary>
        private static Form CreateHiddenForm()
        {
            var form = new Form
            {
                FormBorderStyle = FormBorderStyle.None,
                ShowInTaskbar = false,
                StartPosition = FormStartPosition.Manual,
                Size = new Size(1, 1),
                Location = new Point(-2000, -2000),
                Opacity = 0
            };
            form.Show();
            form.Hide();
            return form;
        }

        private void CreateTrayIconForProgram(ProgramItem item)
        {
            bool isFolder = Directory.Exists(item.Path);

            Icon icon;
            try
            {
                icon = isFolder
                    ? _folderIcon
                    : IconLoader.LoadFileIcon(item.Path) ?? _defaultFileIcon;
            }
            catch
            {
                icon = _defaultFileIcon;
            }

            var notifyIcon = new NotifyIcon
            {
                Icon = icon,
                Text = BuildTooltip(item.Name, item.Path),
                Visible = true
            };

            var menu = new ContextMenuStrip();
            string menuText = isFolder
                ? _localizationService.GetString("OpenFolder")
                : _localizationService.GetString("RunProgram");

            var runItem = new ToolStripMenuItem(menuText);
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

            var exitItem = new ToolStripMenuItem(_localizationService.GetString("TerminateApp"));
            exitItem.Click += (s, e) =>
            {
                var result = MessageBox.Show(
                    _localizationService.GetString("ConfirmTermination"),
                    _localizationService.GetString("Confirm"),
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    ExitThread();
                }
            };
            menu.Items.Add(exitItem);

            notifyIcon.ContextMenuStrip = menu;
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

            _trayIcons.Add(notifyIcon);
        }

        /// <summary>툴팁 문자열을 만든다. 빈 이름은 경로에서 만들고, 상한을 넘는 이름은 잘라낸다.</summary>
        private static string BuildTooltip(string name, string path)
        {
            string text = name;

            if (string.IsNullOrWhiteSpace(text))
            {
                text = ProgramItem.DeriveDisplayName(path);
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                text = path;
            }

            if (text == null)
            {
                return string.Empty;
            }

            return text.Length > MaxTooltipLength ? text.Substring(0, MaxTooltipLength) : text;
        }

        private void ShowFolderContextMenuWithFiles(string folderPath)
        {
            try
            {
                if (!Directory.Exists(folderPath))
                {
                    throw new DirectoryNotFoundException(folderPath);
                }

                var folderMenu = new ContextMenuStrip();

                // Closed에서 곧바로 Dispose하면 ToolStripDropDown.Close() 내부가 파괴된 핸들을 만져
                // ObjectDisposedException이 발생한다(실측). Close()가 끝난 뒤 메시지 큐에서 정리한다.
                folderMenu.Closed += (s, e) => DisposeMenuDeferred(folderMenu);

                AddLoadingPlaceholder(folderMenu.Items);
                LoadFolderMenuItemsAsync(folderPath, folderMenu, folderMenu.Items, true, null);

                if (_hiddenForm != null && !_hiddenForm.IsDisposed)
                {
                    // 메뉴를 띄우기 전에 이 프로세스를 전경으로 만든다.
                    // 그러지 않으면 바깥 클릭으로는 닫히지만, 창 전환(Alt+Tab 등)만으로는 닫히지 않는다.
                    // 셸 트레이 메뉴가 쓰는 표준 방식(SetForegroundWindow → 표시 → WM_NULL)이다.
                    SetForegroundWindow(_hiddenForm.Handle);

                    POINT cursor;
                    if (GetCursorPos(out cursor))
                    {
                        folderMenu.Show(_hiddenForm, _hiddenForm.PointToClient(new Point(cursor.X, cursor.Y)));
                    }
                    else
                    {
                        folderMenu.Show(Cursor.Position);
                    }

                    PostMessage(_hiddenForm.Handle, WmNull, IntPtr.Zero, IntPtr.Zero);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    _localizationService.GetString("ErrorShowingFolderMenu", ex.Message),
                    _localizationService.GetString("Error"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 닫힌 폴더 메뉴를 메시지 큐에서 정리한다. Closed 핸들러에서 바로 Dispose하면
        /// ToolStripDropDown.Close() 내부가 파괴된 핸들을 만져 ObjectDisposedException이 발생한다(실측).
        /// </summary>
        private void DisposeMenuDeferred(ContextMenuStrip menu)
        {
            var hiddenForm = _hiddenForm;
            if (hiddenForm == null || hiddenForm.IsDisposed)
            {
                // 종료 중: 남은 창 핸들은 프로세스 종료로 회수된다.
                return;
            }

            try
            {
                hiddenForm.BeginInvoke(new Action(menu.Dispose));
            }
            catch (InvalidOperationException)
            {
                // 이미 창 핸들이 파괴된 경우(종료 중). ObjectDisposedException도 이 형식에 포함된다.
            }
        }

        private void LoadFolderMenuItemsAsync(string folderPath, ToolStrip owner, ToolStripItemCollection items, bool includeRootExtras, ToolStripMenuItem ownerMenuItem)
        {
            Task.Run(() =>
            {
                try
                {
                    var entries = _folderIndexCache.GetEntries(folderPath);
                    if (_hiddenForm == null || _hiddenForm.IsDisposed)
                    {
                        return;
                    }

                    _hiddenForm.BeginInvoke(new Action(() =>
                    {
                        if ((ownerMenuItem != null && ownerMenuItem.IsDisposed) || items == null)
                        {
                            return;
                        }

                        if (owner != null && owner.IsDisposed)
                        {
                            return;
                        }

                        ApplyFolderEntries(folderPath, owner, items, entries, includeRootExtras, ownerMenuItem);
                    }));
                }
                catch (Exception ex)
                {
                    if (_hiddenForm == null || _hiddenForm.IsDisposed)
                    {
                        return;
                    }

                    _hiddenForm.BeginInvoke(new Action(() =>
                    {
                        if ((ownerMenuItem != null && ownerMenuItem.IsDisposed) || items == null)
                        {
                            return;
                        }

                        if (owner != null && owner.IsDisposed)
                        {
                            return;
                        }

                        items.Clear();
                        items.Add(new ToolStripMenuItem(_localizationService.GetString("Error"))
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

                        var metadata = ownerMenuItem != null ? ownerMenuItem.Tag as FolderMenuMetadata : null;
                        if (metadata != null)
                        {
                            metadata.IsLoading = false;
                        }

                        MessageBox.Show(
                            _localizationService.GetString("ErrorBuildingFolderMenu", ex.Message),
                            _localizationService.GetString("Error"),
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }));
                }
            });
        }

        private void ApplyFolderEntries(string folderPath, ToolStrip owner, ToolStripItemCollection items, IReadOnlyList<FolderEntry> entries, bool includeRootExtras, ToolStripMenuItem ownerMenuItem)
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
                var metadata = ownerMenuItem.Tag as FolderMenuMetadata;
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

            items.Add(CreateRefreshMenuItem(folderPath));

            var openInExplorer = new ToolStripMenuItem(_localizationService.GetString("OpenInExplorer"));
            openInExplorer.Click += (s, e) => Process.Start("explorer.exe", folderPath);
            items.Add(openInExplorer);
        }

        private void AddFolderFooterItems(string folderPath, ToolStripItemCollection items)
        {
            if (items.Count > 0)
            {
                items.Add(new ToolStripSeparator());
            }

            items.Add(CreateRefreshMenuItem(folderPath));

            var openInExplorer = new ToolStripMenuItem(_localizationService.GetString("OpenInExplorer"));
            openInExplorer.Click += (s, e) => Process.Start("explorer.exe", folderPath);
            items.Add(openInExplorer);
        }

        private ToolStripMenuItem CreateRefreshMenuItem(string folderPath)
        {
            var refreshItem = new ToolStripMenuItem(_localizationService.GetString("Refresh"));
            refreshItem.Click += (s, e) => RefreshMenu(folderPath, s as ToolStripMenuItem);
            return refreshItem;
        }

        /// <summary>캐시를 버리고 해당 폴더 메뉴를 다시 읽는다. 실행 중인 메뉴는 클릭으로 닫히며, 다시 열면 갱신된 목록이 보인다.</summary>
        private void RefreshMenu(string folderPath, ToolStripMenuItem clickedItem)
        {
            _folderIndexCache.Invalidate(folderPath);

            var dropdown = clickedItem != null ? clickedItem.GetCurrentParent() as ToolStripDropDown : null;
            var ownerMenuItem = dropdown != null ? dropdown.OwnerItem as ToolStripMenuItem : null;

            if (ownerMenuItem != null)
            {
                var metadata = ownerMenuItem.Tag as FolderMenuMetadata;
                if (metadata != null)
                {
                    metadata.IsLoading = false;
                    metadata.HasEverLoaded = false;
                    metadata.SupportsLoading = true;
                }
            }

            if (dropdown == null)
            {
                return;
            }

            AddLoadingPlaceholder(dropdown.Items);
            LoadFolderMenuItemsAsync(folderPath, dropdown, dropdown.Items, ownerMenuItem == null, ownerMenuItem);
        }

        private void AddLoadingPlaceholder(ToolStripItemCollection items)
        {
            items.Clear();
            items.Add(new ToolStripMenuItem(_localizationService.GetString("Loading"))
            {
                Enabled = false,
                Tag = _loadingTag
            });
        }

        private void AddEmptyPlaceholder(ToolStripItemCollection items)
        {
            items.Add(new ToolStripMenuItem(_localizationService.GetString("NoItemsAvailable"))
            {
                Enabled = false
            });
        }

        private ToolStripMenuItem CreateFolderMenuItem(FolderEntry folder)
        {
            var folderItem = new ToolStripMenuItem(folder.Name, folder.Image ?? _folderBitmap);
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

        private ToolStripMenuItem CreateFileMenuItem(FolderEntry file)
        {
            var fileItem = new ToolStripMenuItem(file.Name, file.Image ?? _defaultFileBitmap);
            fileItem.Click += (s, e) => RunProgram(file.Path);
            return fileItem;
        }

        private void FolderItem_DropDownOpening(object sender, EventArgs e)
        {
            var folderItem = sender as ToolStripMenuItem;
            if (folderItem == null)
            {
                return;
            }

            var metadata = folderItem.Tag as FolderMenuMetadata;
            if (metadata == null)
            {
                return;
            }

            if (!metadata.SupportsLoading)
            {
                // 캐시된 "하위 항목 없음" 상태는 낡을 수 있다(자식 폴더의 변경은 부모 폴더의 수정 시각에 영향을 주지 않는다).
                // 이때만 디스크를 한 번 확인하고, 항목이 생겼으면 다시 읽는다.
                if (!FolderIndexCache.HasChildrenNow(metadata.FolderPath))
                {
                    return;
                }

                metadata.SupportsLoading = true;
                metadata.HasEverLoaded = false;
            }

            if (metadata.IsLoading)
            {
                return;
            }

            bool hasPlaceholder = folderItem.DropDownItems.Count == 1 && ReferenceEquals(folderItem.DropDownItems[0].Tag, _loadingTag);
            bool shouldReload = !metadata.HasEverLoaded || DateTime.UtcNow - metadata.LastLoaded > _cacheDuration || hasPlaceholder;

            if (!shouldReload)
            {
                return;
            }

            metadata.IsLoading = true;
            AddLoadingPlaceholder(folderItem.DropDownItems);
            LoadFolderMenuItemsAsync(metadata.FolderPath, folderItem.DropDown, folderItem.DropDownItems, false, folderItem);
        }

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
                MessageBox.Show(
                    _localizationService.GetString("ProgramRunError", ex.Message),
                    _localizationService.GetString("Error"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        protected override void ExitThreadCore()
        {
            var warmupTimer = _warmupTimer;
            _warmupTimer = null;
            if (warmupTimer != null)
            {
                warmupTimer.Dispose();
            }

            foreach (var icon in _trayIcons)
            {
                var menu = icon.ContextMenuStrip;
                icon.Visible = false;
                icon.ContextMenuStrip = null;
                icon.Dispose();
                if (menu != null)
                {
                    menu.Dispose();
                }
            }
            _trayIcons.Clear();

            _folderIndexCache.Dispose();
            _iconImages.Dispose();

            _folderBitmap?.Dispose();
            _defaultFileBitmap?.Dispose();
            _folderIcon?.Dispose();
            _defaultFileIcon?.Dispose();
            _hiddenForm?.Dispose();

            base.ExitThreadCore();
        }

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    ExitThreadCore();
                }

                _disposed = true;
            }

            base.Dispose(disposing);
        }

        private class FolderMenuMetadata
        {
            public string FolderPath { get; set; }
            public bool SupportsLoading { get; set; }
            public bool HasEverLoaded { get; set; }
            public DateTime LastLoaded { get; set; }
            public bool IsLoading { get; set; }
        }
    }
}
