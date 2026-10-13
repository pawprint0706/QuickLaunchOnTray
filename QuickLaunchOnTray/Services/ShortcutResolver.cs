using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace QuickLaunchOnTray.Services
{
    /// <summary>
    /// 바로가기(.lnk)가 가리키는 대상 경로와 아이콘 위치를 읽는다.
    /// 셸 아이콘 대신 대상 자체의 아이콘을 쓰기 위한 것으로, 해석에 실패하면 false를 돌려준다(호출자가 폴백).
    /// </summary>
    internal static class ShortcutResolver
    {
        private const int PathBufferLength = 1024;

        /// <summary>STGM_READ: 읽기 전용으로 바로가기를 연다.</summary>
        private const uint StgmRead = 0x00000000;

        /// <summary>SLGP_RAWPATH: 환경 변수를 확장하지 않은 원본 경로를 요청한다(직접 확장한다).</summary>
        private const uint SlgpRawPath = 0x00000004;

        /// <summary>바로가기의 대상 경로와 아이콘 위치를 읽는다.</summary>
        /// <param name="targetPath">대상 경로(환경 변수 확장됨). 없으면 null.</param>
        /// <param name="iconFile">바로가기에 명시된 아이콘 파일. 없으면 null.</param>
        /// <param name="iconIndex">명시된 아이콘 인덱스(음수면 리소스 ID).</param>
        public static bool TryResolve(string shortcutPath, out string targetPath, out string iconFile, out int iconIndex)
        {
            targetPath = null;
            iconFile = null;
            iconIndex = 0;

            object linkObject = null;
            try
            {
                linkObject = new ShellLink();
                var link = (IShellLinkW)linkObject;
                ((IPersistFile)linkObject).Load(shortcutPath, StgmRead);

                var pathBuffer = new StringBuilder(PathBufferLength);
                link.GetPath(pathBuffer, pathBuffer.Capacity, IntPtr.Zero, SlgpRawPath);
                targetPath = Expand(pathBuffer.ToString());

                var iconBuffer = new StringBuilder(PathBufferLength);
                int index;
                link.GetIconLocation(iconBuffer, iconBuffer.Capacity, out index);
                string resolvedIcon = Expand(iconBuffer.ToString());
                if (!string.IsNullOrEmpty(resolvedIcon))
                {
                    iconFile = MakeAbsolute(shortcutPath, resolvedIcon);
                    iconIndex = index;
                }

                return !string.IsNullOrEmpty(targetPath) || !string.IsNullOrEmpty(iconFile);
            }
            catch
            {
                // .lnk가 아니거나 대상이 셸 항목인 경우 등: 호출자가 셸 아이콘으로 폴백한다.
                targetPath = null;
                iconFile = null;
                iconIndex = 0;
                return false;
            }
            finally
            {
                if (linkObject != null && Marshal.IsComObject(linkObject))
                {
                    Marshal.ReleaseComObject(linkObject);
                }
            }
        }

        private static string Expand(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }

            try
            {
                return Environment.ExpandEnvironmentVariables(value);
            }
            catch
            {
                return value;
            }
        }

        /// <summary>상대 경로로 저장된 아이콘 위치는 바로가기가 있는 폴더를 기준으로 해석한다.</summary>
        private static string MakeAbsolute(string shortcutPath, string iconPath)
        {
            try
            {
                if (Path.IsPathRooted(iconPath))
                {
                    return iconPath;
                }

                string directory = Path.GetDirectoryName(shortcutPath);
                return string.IsNullOrEmpty(directory) ? iconPath : Path.GetFullPath(Path.Combine(directory, iconPath));
            }
            catch
            {
                return iconPath;
            }
        }

        [ComImport]
        [Guid("00021401-0000-0000-C000-000000000046")]
        private class ShellLink
        {
        }

        [ComImport]
        [Guid("000214F9-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellLinkW
        {
            void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cch, IntPtr pfd, uint fFlags);

            void GetIDList(out IntPtr ppidl);

            void SetIDList(IntPtr pidl);

            void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);

            void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);

            void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cch);

            void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);

            void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cch);

            void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);

            void GetHotkey(out short pwHotkey);

            void SetHotkey(short wHotkey);

            void GetShowCmd(out int piShowCmd);

            void SetShowCmd(int iShowCmd);

            void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cch, out int piIcon);

            void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);

            void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);

            void Resolve(IntPtr hwnd, uint fFlags);

            void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
        }

        [ComImport]
        [Guid("0000010b-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPersistFile
        {
            void GetClassID(out Guid pClassID);

            // S_FALSE(=1)을 실패로 취급하지 않도록 HRESULT를 그대로 받는다.
            [PreserveSig]
            int IsDirty();

            void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, uint dwMode);

            void Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, bool fRemember);

            void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);

            void GetCurFile([Out, MarshalAs(UnmanagedType.LPWStr)] out string ppszFileName);
        }
    }
}
