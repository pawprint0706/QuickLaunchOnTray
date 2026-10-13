using System;
using System.IO;

namespace QuickLaunchOnTray.Models
{
    /// <summary>트레이에 등록할 프로그램 또는 폴더 한 항목.</summary>
    public class ProgramItem
    {
        /// <summary>툴팁에 표시할 이름. config.ini의 key 값, key가 없으면 경로에서 만든 이름.</summary>
        public string Name { get; set; }

        /// <summary>실행 파일 또는 폴더의 경로.</summary>
        public string Path { get; set; }

        /// <summary>
        /// 경로에서 표시 이름을 만든다. 파일은 확장자를 제외한 파일명, 폴더는 폴더명을 사용한다.
        /// 경로에서 이름을 얻을 수 없으면(예: 드라이브 루트) 경로를 그대로 반환한다.
        /// </summary>
        public static string DeriveDisplayName(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return path;
            }

            string trimmed = path.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar);
            string name = System.IO.Path.GetFileName(trimmed);

            if (string.IsNullOrEmpty(name))
            {
                return path;
            }

            if (Directory.Exists(path))
            {
                return name;
            }

            int extensionIndex = name.LastIndexOf('.');
            return extensionIndex > 0 ? name.Substring(0, extensionIndex) : name;
        }
    }
}
