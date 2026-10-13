using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace QuickLaunchOnTray.Services
{
    /// <summary>
    /// 디스크 캐시의 위치와 파일 조작을 담당한다.
    /// 캐시는 언제든 없을 수 있으므로 모든 메서드는 예외를 던지지 않고 실패를 알리기만 한다(기능 저하 없음).
    /// </summary>
    internal static class CacheStorage
    {
        /// <summary>캐시 포맷 버전. 저장 형식을 바꾸면 올린다(다른 버전의 파일은 읽지 않는다).</summary>
        public const string FormatVersion = "1";

        private const int MaxFolderIndexFiles = 5000;
        private const int MaxIconFiles = 20000;

        private static readonly TimeSpan FolderIndexMaxAge = TimeSpan.FromDays(90);
        private static readonly TimeSpan IconMaxAge = TimeSpan.FromDays(30);

        private static readonly string _root = BuildRoot();

        public static string Root
        {
            get { return _root; }
        }

        public static string FolderIndexDirectory
        {
            get { return Path.Combine(_root, "folders"); }
        }

        public static string IconDirectory
        {
            get { return Path.Combine(_root, "icons"); }
        }

        /// <summary>캐시 폴더를 만들고 경로를 반환한다. 만들 수 없으면 null(캐시 없이 동작).</summary>
        public static string EnsureDirectory(string directory)
        {
            if (string.IsNullOrEmpty(directory))
            {
                return null;
            }

            try
            {
                Directory.CreateDirectory(directory);
                return Directory.Exists(directory) ? directory : null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>경로를 파일명에 쓸 수 있는 고정 길이 키로 바꾼다(대소문자 무시).</summary>
        public static string ComputeKey(string value)
        {
            try
            {
                using (var sha1 = SHA1.Create())
                {
                    byte[] hash = sha1.ComputeHash(Encoding.UTF8.GetBytes((value ?? string.Empty).ToUpperInvariant()));
                    var builder = new StringBuilder(hash.Length * 2);
                    foreach (byte b in hash)
                    {
                        builder.Append(b.ToString("x2"));
                    }
                    return builder.ToString();
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>임시 파일에 쓴 뒤 교체해 읽는 쪽이 반쪽 파일을 보지 않게 한다.</summary>
        public static bool TryWrite(string path, IEnumerable<string> lines)
        {
            return TryWrite(path, new UTF8Encoding(false).GetBytes(string.Join(Environment.NewLine, lines) + Environment.NewLine));
        }

        public static bool TryWrite(string path, byte[] data)
        {
            if (string.IsNullOrEmpty(path) || data == null)
            {
                return false;
            }

            string temporary = path + ".tmp";
            try
            {
                File.WriteAllBytes(temporary, data);
                if (File.Exists(path))
                {
                    File.Delete(path);
                }

                File.Move(temporary, path);
                return true;
            }
            catch
            {
                TryDelete(temporary);
                return false;
            }
        }

        public static void TryDelete(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }

            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
                // 캐시 정리 실패는 무시한다.
            }
        }

        /// <summary>지정한 폴더에서 오래된 파일을 지우고 남은 파일 수가 상한을 넘지 않게 정리한다.</summary>
        public static void Prune(string directory, TimeSpan maxAge, int maxFiles)
        {
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
            {
                return;
            }

            try
            {
                var files = new List<FileInfo>();
                DateTime limit = DateTime.UtcNow - maxAge;

                foreach (string file in Directory.EnumerateFiles(directory))
                {
                    try
                    {
                        var info = new FileInfo(file);
                        if (info.LastWriteTimeUtc < limit)
                        {
                            TryDelete(file);
                        }
                        else
                        {
                            files.Add(info);
                        }
                    }
                    catch
                    {
                        // 개별 파일 오류는 무시한다.
                    }
                }

                if (files.Count <= maxFiles)
                {
                    return;
                }

                files.Sort((x, y) => x.LastWriteTimeUtc.CompareTo(y.LastWriteTimeUtc));
                for (int i = 0; i < files.Count - maxFiles; i++)
                {
                    TryDelete(files[i].FullName);
                }
            }
            catch
            {
                // 캐시 정리 실패는 무시한다.
            }
        }

        /// <summary>폴더 목록 캐시와 아이콘 캐시를 각각의 상한으로 정리한다.</summary>
        public static void PruneAll()
        {
            Prune(FolderIndexDirectory, FolderIndexMaxAge, MaxFolderIndexFiles);
            Prune(IconDirectory, IconMaxAge, MaxIconFiles);
        }

        private static string BuildRoot()
        {
            string baseDirectory = null;
            try
            {
                baseDirectory = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            }
            catch
            {
                // 아래 폴백을 사용한다.
            }

            if (string.IsNullOrEmpty(baseDirectory))
            {
                try
                {
                    baseDirectory = Path.GetTempPath();
                }
                catch
                {
                    return null;
                }
            }

            return Path.Combine(baseDirectory, "QuickLaunchOnTray", "cache", "v" + FormatVersion);
        }
    }
}
