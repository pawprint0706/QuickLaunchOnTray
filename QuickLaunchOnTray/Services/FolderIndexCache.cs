using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using QuickLaunchOnTray.Models;

namespace QuickLaunchOnTray.Services
{
    /// <summary>
    /// 폴더 항목 목록을 메모리와 디스크에 캐시한다.
    /// 디스크 캐시는 폴더의 마지막 수정 시각으로 검증하므로, 폴더 내용이 그대로면 지난 실행에서 만들어 둔 목록을 그대로 쓴다.
    /// (파일 내용만 바뀐 경우에는 폴더 수정 시각이 바뀌지 않는데, 이는 목록과 무관하므로 문제가 되지 않는다.)
    /// 캐시 실패는 기능에 영향을 주지 않는다(그때마다 디스크를 직접 읽는다).
    /// </summary>
    internal class FolderIndexCache : IDisposable
    {
        private const string HeaderPrefix = "#QLT-FOLDER-";
        private const string PathField = "P";
        private const string TimeField = "T";
        private const string DirectoryWithChildrenField = "D";
        private const string DirectoryWithoutChildrenField = "E";
        private const string FileField = "F";
        private const string CacheFileExtension = ".idx";

        private readonly TimeSpan _memoryLifetime = TimeSpan.FromSeconds(30);
        private readonly ConcurrentDictionary<string, MemoryEntry> _memory =
            new ConcurrentDictionary<string, MemoryEntry>(StringComparer.OrdinalIgnoreCase);
        private readonly IconImageCache _iconImages;
        private readonly string _directory;
        private bool _disposed;

        public FolderIndexCache(IconImageCache iconImages)
        {
            _iconImages = iconImages;
            _directory = CacheStorage.EnsureDirectory(CacheStorage.FolderIndexDirectory);
        }

        /// <summary>폴더 항목 목록을 반환한다. 폴더가 없거나 읽을 수 없으면 빈 목록.</summary>
        public IReadOnlyList<FolderEntry> GetEntries(string folderPath)
        {
            if (string.IsNullOrEmpty(folderPath))
            {
                return new List<FolderEntry>();
            }

            DateTime lastWriteUtc = GetLastWriteTimeUtc(folderPath);

            MemoryEntry cached;
            if (_memory.TryGetValue(folderPath, out cached)
                && DateTime.UtcNow - cached.LoadedAt <= _memoryLifetime
                && cached.LastWriteUtc == lastWriteUtc)
            {
                return cached.Entries;
            }

            string cacheFile = BuildCacheFilePath(folderPath);
            var fromDisk = TryLoadFromDisk(cacheFile, folderPath, lastWriteUtc);

            bool exists = Directory.Exists(folderPath);
            var entries = fromDisk ?? Enumerate(folderPath);

            if (fromDisk == null && exists)
            {
                // 폴더가 없는 상태(빈 목록)는 캐시하지 않는다.
                TrySaveToDisk(cacheFile, folderPath, lastWriteUtc, entries);
            }

            _memory[folderPath] = new MemoryEntry
            {
                LoadedAt = DateTime.UtcNow,
                LastWriteUtc = lastWriteUtc,
                Entries = entries
            };

            return entries;
        }

        /// <summary>메모리와 디스크 캐시를 모두 버린다(다음 조회 때 디스크를 다시 읽는다).</summary>
        public void Invalidate(string folderPath)
        {
            if (string.IsNullOrEmpty(folderPath))
            {
                return;
            }

            MemoryEntry ignored;
            _memory.TryRemove(folderPath, out ignored);
            CacheStorage.TryDelete(BuildCacheFilePath(folderPath));
        }

        /// <summary>백그라운드 예열용. 예외를 던지지 않는다.</summary>
        public void Warm(string folderPath)
        {
            try
            {
                GetEntries(folderPath);
            }
            catch
            {
                // 예열 실패는 무시한다.
            }
        }

        /// <summary>하위 폴더에 지금 자식이 있는지 확인한다. 캐시된 "비어 있음" 상태가 낡았을 때만 쓴다.</summary>
        public static bool HasChildrenNow(string folderPath)
        {
            try
            {
                using (var enumerator = Directory.EnumerateFileSystemEntries(folderPath).GetEnumerator())
                {
                    return enumerator.MoveNext();
                }
            }
            catch
            {
                return false;
            }
        }

        private static DateTime GetLastWriteTimeUtc(string folderPath)
        {
            try
            {
                return Directory.GetLastWriteTimeUtc(folderPath);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>폴더를 실제로 읽는다. 접근할 수 없는 하위 폴더는 캐시 단계에서 걸러진다.</summary>
        private IReadOnlyList<FolderEntry> Enumerate(string folderPath)
        {
            var entries = new List<FolderEntry>();

            try
            {
                foreach (string subFolder in Directory.EnumerateDirectories(folderPath))
                {
                    entries.Add(new FolderEntry
                    {
                        Name = Path.GetFileName(subFolder),
                        Path = subFolder,
                        IsFolder = true,
                        HasChildren = HasChildrenNow(subFolder)
                    });
                }
            }
            catch
            {
                // 폴더 열거 실패는 무시한다.
            }

            try
            {
                foreach (string file in Directory.EnumerateFiles(folderPath))
                {
                    entries.Add(new FolderEntry
                    {
                        Name = Path.GetFileName(file),
                        Path = file,
                        Image = _iconImages.GetImage(file)
                    });
                }
            }
            catch
            {
                // 파일 열거 실패는 무시한다.
            }

            entries.Sort((x, y) =>
            {
                int typeCompare = y.IsFolder.CompareTo(x.IsFolder);
                return typeCompare != 0
                    ? typeCompare
                    : string.Compare(x.Name, y.Name, StringComparison.CurrentCultureIgnoreCase);
            });

            return entries;
        }

        private string BuildCacheFilePath(string folderPath)
        {
            if (_directory == null || string.IsNullOrEmpty(folderPath))
            {
                return null;
            }

            string key = CacheStorage.ComputeKey(folderPath);
            return key == null ? null : Path.Combine(_directory, key + CacheFileExtension);
        }

        private IReadOnlyList<FolderEntry> TryLoadFromDisk(string cacheFile, string folderPath, DateTime lastWriteUtc)
        {
            if (cacheFile == null)
            {
                return null;
            }

            try
            {
                if (!File.Exists(cacheFile))
                {
                    return null;
                }

                string[] lines = File.ReadAllLines(cacheFile);
                if (lines.Length < 3)
                {
                    return null;
                }

                if (!string.Equals(lines[0].Trim(), HeaderPrefix + CacheStorage.FormatVersion, StringComparison.Ordinal))
                {
                    return null;
                }

                string recordedPath;
                if (!TryReadField(lines[1], PathField, out recordedPath)
                    || !string.Equals(recordedPath, folderPath, StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                string recordedTicks;
                long ticks;
                if (!TryReadField(lines[2], TimeField, out recordedTicks)
                    || !long.TryParse(recordedTicks, NumberStyles.Integer, CultureInfo.InvariantCulture, out ticks)
                    || ticks != lastWriteUtc.Ticks)
                {
                    return null;
                }

                var entries = new List<FolderEntry>(Math.Max(lines.Length - 3, 0));
                for (int index = 3; index < lines.Length; index++)
                {
                    string line = lines[index];
                    if (string.IsNullOrEmpty(line))
                    {
                        continue;
                    }

                    string path;
                    char field = line[0];
                    if (!TryReadField(line, field.ToString(), out path) || string.IsNullOrEmpty(path))
                    {
                        return null;
                    }

                    var entry = new FolderEntry { Path = path, Name = Path.GetFileName(path) };
                    switch (field)
                    {
                        case 'D':
                            entry.IsFolder = true;
                            entry.HasChildren = true;
                            break;
                        case 'E':
                            entry.IsFolder = true;
                            break;
                        case 'F':
                            entry.Image = _iconImages.GetImage(path);
                            break;
                        default:
                            return null;
                    }

                    entries.Add(entry);
                }

                return entries;
            }
            catch
            {
                return null;
            }
        }

        private static bool TryReadField(string line, string field, out string value)
        {
            value = null;
            if (line == null || line.Length < field.Length + 2 || !line.StartsWith(field, StringComparison.Ordinal) || line[field.Length] != '\t')
            {
                return false;
            }

            value = line.Substring(field.Length + 1);
            return true;
        }

        private static void TrySaveToDisk(string cacheFile, string folderPath, DateTime lastWriteUtc, IReadOnlyList<FolderEntry> entries)
        {
            if (cacheFile == null)
            {
                return;
            }

            var lines = new List<string>(entries.Count + 3)
            {
                HeaderPrefix + CacheStorage.FormatVersion,
                PathField + "\t" + folderPath,
                TimeField + "\t" + lastWriteUtc.Ticks.ToString(CultureInfo.InvariantCulture)
            };

            foreach (var entry in entries)
            {
                string field = !entry.IsFolder
                    ? FileField
                    : (entry.HasChildren ? DirectoryWithChildrenField : DirectoryWithoutChildrenField);
                lines.Add(field + "\t" + entry.Path);
            }

            CacheStorage.TryWrite(cacheFile, lines);
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            // 항목의 이미지는 IconImageCache가 소유하므로 여기서는 참조만 버린다.
            _memory.Clear();
            _disposed = true;
        }

        private class MemoryEntry
        {
            public DateTime LoadedAt { get; set; }
            public DateTime LastWriteUtc { get; set; }
            public IReadOnlyList<FolderEntry> Entries { get; set; }
        }
    }
}
