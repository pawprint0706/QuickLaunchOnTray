using System;
using System.Collections.Concurrent;
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
        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool DestroyIcon(IntPtr handle);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        private readonly List<NotifyIcon> _trayIcons = new List<NotifyIcon>();
        private readonly ConfigurationService _configService;
        private readonly LocalizationService _localizationService;
        private readonly ConcurrentDictionary<string, Bitmap> _iconCache =
            new ConcurrentDictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);
        private readonly ConcurrentDictionary<string, FolderCacheEntry> _folderCache =
            new ConcurrentDictionary<string, FolderCacheEntry>(StringComparer.OrdinalIgnoreCase);
        private readonly TimeSpan _cacheDuration = TimeSpan.FromSeconds(30);
        private readonly object _loadingTag = new object();

        private readonly Icon _folderIcon;
        private readonly Bitmap _folderBitmap;
        private readonly Bitmap _defaultFileBitmap;
        private readonly Form _hiddenForm;
        private bool _disposed;

        public TrayApplicationContext()
        {
            _configService = new ConfigurationService();
            _localizationService = LocalizationService.Instance;

            _folderIcon = InitializeFolderIcon();
            _folderBitmap = _folderIcon?.ToBitmap();
            _defaultFileBitmap = SystemIcons.Application.ToBitmap();

            try
            {
                var programItems = _configService.LoadProgramItems();
                foreach (var item in programItems)
                {
                    CreateTrayIconForProgram(item);
                }

                _hiddenForm = new Form
                {
                    FormBorderStyle = FormBorderStyle.None,
                    ShowInTaskbar = false,
                    StartPosition = FormStartPosition.Manual,
                    Size = new Size(1, 1),
                    Location = new Point(-2000, -2000),
                    Opacity = 0
                };
                _hiddenForm.Show();
                _hiddenForm.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    _localizationService.GetString("Error"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                Environment.Exit(1);
            }
        }

        private Icon InitializeFolderIcon()
        {
            try
            {
                string systemFolder = Environment.GetFolderPath(Environment.SpecialFolder.System);
                IntPtr handle = ExtractIcon(IntPtr.Zero, Path.Combine(systemFolder, "shell32.dll"), 3);
                if (handle != IntPtr.Zero)
                {
                    using (var tempIcon = Icon.FromHandle(handle))
                    {
                        Icon clone = (Icon)tempIcon.Clone();
                        DestroyIcon(handle);
                        return clone;
                    }
                }
            }
            catch
            {
                // Ignore and fall back to default.
            }

            return SystemIcons.Application;
        }

        private void CreateTrayIconForProgram(ProgramItem item)
        {
            Icon icon;
            try
            {
                if (Directory.Exists(item.Path))
                {
                    icon = _folderIcon;
                }
                else if (File.Exists(item.Path))
                {
                    icon = Icon.ExtractAssociatedIcon(item.Path) ?? SystemIcons.Application;
                }
                else
                {
                    icon = SystemIcons.Application;
                }
            }
            catch
            {
                icon = SystemIcons.Application;
            }

            var notifyIcon = new NotifyIcon
            {
                Icon = icon,
                Text = item.Name,
                Visible = true
            };

            var menu = new ContextMenuStrip();
            string menuText = Directory.Exists(item.Path)
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

        private void ShowFolderContextMenuWithFiles(string folderPath)
        {
            try
            {
                if (!Directory.Exists(folderPath))
                {
                    throw new DirectoryNotFoundException(folderPath);
                }

                var folderMenu = new ContextMenuStrip();
                AddLoadingPlaceholder(folderMenu.Items);
                LoadFolderMenuItemsAsync(folderPath, folderMenu, folderMenu.Items, true, null);

                if (_hiddenForm != null && !_hiddenForm.IsDisposed)
                {
                    if (GetCursorPos(out POINT cursor))
                    {
                        folderMenu.Show(_hiddenForm, _hiddenForm.PointToClient(new Point(cursor.X, cursor.Y)));
                    }
                    else
                    {
                        folderMenu.Show(Cursor.Position);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    GetLocalizedText("폴더 메뉴 표시 중 오류 발생: " + ex.Message, "Error showing folder menu: " + ex.Message),
                    _localizationService.GetString("Error"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void LoadFolderMenuItemsAsync(string folderPath, ToolStrip owner, ToolStripItemCollection items, bool includeRootExtras, ToolStripMenuItem ownerMenuItem)
        {
            Task.Run(() =>
            {
                try
                {
                    var entries = GetFolderEntries(folderPath);
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

                        var metadata = ownerMenuItem != null ? ownerMenuItem.Tag as FolderMenuMetadata : null;
                        if (metadata != null)
                        {
                            metadata.IsLoading = false;
                        }

                        MessageBox.Show(
                            GetLocalizedText("폴더 메뉴 구성 중 오류 발생: " + ex.Message, "Error building folder menu: " + ex.Message),
                            _localizationService.GetString("Error"),
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }));
                }
            });
        }

        private void ApplyFolderEntries(string folderPath, ToolStrip owner, ToolStripItemCollection items, IReadOnlyList<FolderItemInfo> entries, bool includeRootExtras, ToolStripMenuItem ownerMenuItem)
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

            var openInExplorer = new ToolStripMenuItem(GetLocalizedText("탐색기에서 열기", "Open in Explorer"));
            openInExplorer.Click += (s, e) => Process.Start("explorer.exe", folderPath);
            items.Add(openInExplorer);

            var closeMenuItem = new ToolStripMenuItem(GetLocalizedText("닫기", "Close"));
            closeMenuItem.Click += (s, e) =>
            {
                var menuItem = s as ToolStripMenuItem;
                var parent = menuItem?.GetCurrentParent() as ToolStripDropDown;
                parent?.Close();
            };
            items.Add(closeMenuItem);
        }

        private void AddFolderFooterItems(string folderPath, ToolStripItemCollection items)
        {
            if (items.Count > 0)
            {
                items.Add(new ToolStripSeparator());
            }

            var openInExplorer = new ToolStripMenuItem(GetLocalizedText("탐색기에서 열기", "Open in Explorer"));
            openInExplorer.Click += (s, e) => Process.Start("explorer.exe", folderPath);
            items.Add(openInExplorer);
        }

        private void AddLoadingPlaceholder(ToolStripItemCollection items)
        {
            items.Clear();
            items.Add(new ToolStripMenuItem(GetLocalizedText("불러오는 중...", "Loading..."))
            {
                Enabled = false,
                Tag = _loadingTag
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
            var folderItem = new ToolStripMenuItem(folder.Name, folder.Icon ?? _folderBitmap);
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
            var fileItem = new ToolStripMenuItem(file.Name, file.Icon ?? _defaultFileBitmap);
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
            if (metadata == null || !metadata.SupportsLoading)
            {
                return;
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

        private IReadOnlyList<FolderItemInfo> GetFolderEntries(string folderPath)
        {
            var now = DateTime.UtcNow;
            FolderCacheEntry cacheEntry;
            if (_folderCache.TryGetValue(folderPath, out cacheEntry))
            {
                if (now - cacheEntry.CachedAt <= _cacheDuration)
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
            _folderCache[folderPath] = entry;
            return entry.Items;
        }

        private IReadOnlyList<FolderItemInfo> GetFolderItemsFromDisk(string folderPath)
        {
            var items = new List<FolderItemInfo>();

            try
            {
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
                        Icon = _folderBitmap
                    });
                }
            }
            catch
            {
                // Ignore directory enumeration issues.
            }

            try
            {
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
            }
            catch
            {
                // Ignore file enumeration issues.
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

        private Bitmap GetIconBitmap(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return _defaultFileBitmap;
            }

            string extension = Path.GetExtension(filePath) ?? string.Empty;
            bool useFullPath = extension.Equals(".exe", StringComparison.OrdinalIgnoreCase) || extension.Equals(".lnk", StringComparison.OrdinalIgnoreCase);
            string cacheKey = useFullPath ? filePath : extension;

            return _iconCache.GetOrAdd(cacheKey, key =>
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
                    // Ignore icon extraction errors.
                }

                return _defaultFileBitmap;
            });
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
            foreach (var icon in _trayIcons)
            {
                icon.Visible = false;
                icon.Dispose();
            }
            _trayIcons.Clear();

            var disposedBitmaps = new HashSet<Bitmap>();
            foreach (var bitmap in _iconCache.Values)
            {
                if (bitmap == null)
                {
                    continue;
                }

                if (ReferenceEquals(bitmap, _defaultFileBitmap) || ReferenceEquals(bitmap, _folderBitmap))
                {
                    continue;
                }

                if (disposedBitmaps.Add(bitmap))
                {
                    bitmap.Dispose();
                }
            }
            _iconCache.Clear();
            _folderCache.Clear();

            _folderBitmap?.Dispose();
            _defaultFileBitmap?.Dispose();
            _folderIcon?.Dispose();
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

        private string GetLocalizedText(string koreanText, string englishText)
        {
            return _localizationService.IsKoreanSystem() ? koreanText : englishText;
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
    }
}
