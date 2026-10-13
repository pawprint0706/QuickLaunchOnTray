using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;

namespace QuickLaunchOnTray.Services
{
    /// <summary>
    /// 트레이 알림 영역에 맞는 작은 크기(DPI 배율 반영) 아이콘을 만든다.
    /// 모든 핸들은 복제 후 즉시 DestroyIcon으로 해제하므로 호출자는 반환된 Icon만 Dispose하면 된다.
    /// </summary>
    internal static class IconLoader
    {
        /// <summary>shell32.dll에서 폴더 아이콘으로 쓰는 인덱스.</summary>
        public const int FolderIconIndex = 3;

        private const int SM_CXSMICON = 49;
        private const uint SHGFI_ICON = 0x00000100;
        private const uint SHGFI_SMALLICON = 0x00000001;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        }

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern uint ExtractIconEx(string lpszFile, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

        [DllImport("user32.dll")]
        private static extern bool DestroyIcon(IntPtr handle);

        /// <summary>알림 영역 아이콘 크기(픽셀). DPI 인식 프로세스에서는 배율이 반영된다.</summary>
        public static int SmallIconSize
        {
            get
            {
                int size = GetSystemMetrics(SM_CXSMICON);
                return size > 0 ? size : 16;
            }
        }

        /// <summary>실행 중인 시스템의 기본 아이콘을 작은 크기로 복제한다. 실패하지 않는 폴백용.</summary>
        public static Icon CreateDefaultIcon()
        {
            return Resize(SystemIcons.Application, SmallIconSize);
        }

        /// <summary>shell32.dll의 아이콘 인덱스를 작은 크기로 추출한다. 실패하면 null.</summary>
        public static Icon LoadShellIcon(int iconIndex)
        {
            string shell32 = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                "shell32.dll");
            return ExtractSmallIcon(shell32, iconIndex);
        }

        /// <summary>
        /// 파일의 작은 아이콘을 돌려준다. 바로가기는 대상 프로그램의 아이콘을 사용하므로 바로가기 화살표가 붙지 않는다.
        /// </summary>
        public static Icon LoadFileIcon(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return null;
            }

            if (IsShortcut(filePath))
            {
                Icon targetIcon = LoadShortcutTargetIcon(filePath);
                if (targetIcon != null)
                {
                    return targetIcon;
                }

                // 대상을 해석할 수 없는 바로가기(대상 없음, 셸 항목)는 셸이 표시하는 아이콘을 그대로 쓴다.
            }

            return LoadFileIconCore(filePath);
        }

        private static bool IsShortcut(string filePath)
        {
            return filePath.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 바로가기가 가리키는 대상(또는 바로가기에 명시된 아이콘 파일)의 아이콘을 돌려준다.
        /// </summary>
        private static Icon LoadShortcutTargetIcon(string filePath)
        {
            string targetPath;
            string iconFile;
            int iconIndex;
            if (!ShortcutResolver.TryResolve(filePath, out targetPath, out iconFile, out iconIndex))
            {
                return null;
            }

            // 바로가기에 아이콘이 지정되어 있으면 탐색기와 같이 그것을 우선한다.
            if (!string.IsNullOrEmpty(iconFile))
            {
                Icon explicitIcon = ExtractSmallIcon(iconFile, iconIndex);
                if (explicitIcon != null)
                {
                    return ResizeToSmall(explicitIcon);
                }
            }

            if (string.IsNullOrEmpty(targetPath))
            {
                return null;
            }

            if (Directory.Exists(targetPath))
            {
                return LoadShellIcon(FolderIconIndex);
            }

            return File.Exists(targetPath) ? LoadFileIconCore(targetPath) : null;
        }

        /// <summary>확보한 아이콘을 원하는 크기로 맞추고 원본을 해제한다.</summary>
        private static Icon ResizeToSmall(Icon icon)
        {
            Icon resized = Resize(icon, SmallIconSize);
            if (resized != icon)
            {
                DisposeIcon(icon);
            }

            return resized;
        }

        /// <summary>
        /// 파일 자체의 아이콘을 크기 우선 규칙으로 얻는다.
        /// 1) ExtractIconEx(파일 고유 아이콘), 2) 셸이 표시하는 작은 아이콘, 3) 확보한 아이콘을 원하는 크기로 맞춤.
        /// </summary>
        private static Icon LoadFileIconCore(string filePath)
        {
            int desiredSize = SmallIconSize;

            Icon fileIcon = ExtractSmallIcon(filePath, 0);
            if (HasSize(fileIcon, desiredSize))
            {
                return fileIcon;
            }

            Icon shellIcon = LoadSmallShellIcon(filePath);
            if (HasSize(shellIcon, desiredSize))
            {
                DisposeIcon(fileIcon);
                return shellIcon;
            }

            Icon source = fileIcon ?? shellIcon ?? LoadAssociatedIcon(filePath);
            Icon resized = Resize(source, desiredSize);

            DisposeIcon(fileIcon);
            DisposeIcon(shellIcon);
            if (source != fileIcon && source != shellIcon)
            {
                DisposeIcon(source);
            }

            return resized;
        }

        /// <summary>셸이 해당 파일에 대해 표시하는 작은 아이콘. 실패하면 null.</summary>
        private static Icon LoadSmallShellIcon(string filePath)
        {
            try
            {
                var info = new SHFILEINFO();
                IntPtr result = SHGetFileInfo(
                    filePath,
                    0,
                    ref info,
                    (uint)Marshal.SizeOf(typeof(SHFILEINFO)),
                    SHGFI_ICON | SHGFI_SMALLICON);

                if (result == IntPtr.Zero || info.hIcon == IntPtr.Zero)
                {
                    return null;
                }

                try
                {
                    using (var icon = Icon.FromHandle(info.hIcon))
                    {
                        return (Icon)icon.Clone();
                    }
                }
                finally
                {
                    DestroyIcon(info.hIcon);
                }
            }
            catch
            {
                return null;
            }
        }

        private static Icon LoadAssociatedIcon(string filePath)
        {
            try
            {
                return Icon.ExtractAssociatedIcon(filePath);
            }
            catch
            {
                return null;
            }
        }

        private static bool HasSize(Icon icon, int size)
        {
            return icon != null && icon.Width == size && icon.Height == size;
        }

        private static void DisposeIcon(Icon icon)
        {
            if (icon != null)
            {
                icon.Dispose();
            }
        }

        private static Icon ExtractSmallIcon(string filePath, int iconIndex)
        {
            IntPtr[] smallIcons = new IntPtr[1];
            try
            {
                if (ExtractIconEx(filePath, iconIndex, null, smallIcons, 1) == 0 || smallIcons[0] == IntPtr.Zero)
                {
                    return null;
                }

                using (var icon = Icon.FromHandle(smallIcons[0]))
                {
                    return (Icon)icon.Clone();
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                if (smallIcons[0] != IntPtr.Zero)
                {
                    DestroyIcon(smallIcons[0]);
                }
            }
        }

        private static Icon Resize(Icon source, int size)
        {
            if (source == null)
            {
                return null;
            }

            return source.Width == size && source.Height == size
                ? (Icon)source.Clone()
                : new Icon(source, new Size(size, size));
        }
    }
}
