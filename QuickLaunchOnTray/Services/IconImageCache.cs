using System;
using System.Collections.Concurrent;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;

namespace QuickLaunchOnTray.Services
{
    /// <summary>
    /// 메뉴에 쓸 아이콘 이미지를 메모리와 디스크에 캐시한다.
    /// 아이콘 추출은 개당 0.7~7ms인데 반해 캐시된 PNG 로드는 개당 0.06ms 수준이라(실측) 다시 실행할 때 큰 이득이 있다.
    /// 캐시 실패는 기능에 영향을 주지 않는다(그때마다 추출).
    /// </summary>
    internal class IconImageCache : IDisposable
    {
        private const string FilePrefix = "p";
        private const string ExtensionPrefix = "x";
        private const string PngExtension = ".png";

        /// <summary>파일별로 아이콘을 따로 캐시하는 확장자(실행 파일과 바로가기).</summary>
        private static readonly string[] PerFilePathExtensions = { ".exe", ".lnk" };

        private readonly ConcurrentDictionary<string, Bitmap> _memory =
            new ConcurrentDictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);

        private readonly string _directory;
        private bool _disposed;

        public IconImageCache()
        {
            _directory = CacheStorage.EnsureDirectory(CacheStorage.IconDirectory);
        }

        /// <summary>파일 표시용 아이콘 이미지를 돌려준다. 아이콘을 만들 수 없으면 null.</summary>
        public Bitmap GetImage(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return null;
            }

            string extension;
            try
            {
                extension = Path.GetExtension(filePath);
            }
            catch
            {
                return null;
            }

            bool perFilePath = IsPerFilePathExtension(extension);
            string memoryKey = perFilePath ? filePath : extension;

            // null 값은 "아이콘 없음"을 뜻한다(같은 실패를 반복하지 않기 위해 저장한다).
            Bitmap cached;
            if (_memory.TryGetValue(memoryKey, out cached))
            {
                return cached;
            }

            Bitmap image = perFilePath ? LoadPerFilePath(filePath) : LoadPerExtension(extension, filePath);
            _memory[memoryKey] = image;
            return image;
        }

        private static bool IsPerFilePathExtension(string extension)
        {
            foreach (string candidate in PerFilePathExtensions)
            {
                if (extension.Equals(candidate, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsShortcut(string filePath)
        {
            return filePath.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 실행 파일/바로가기는 파일 상태(수정 시각·크기)를 파일명에 넣어 스스로 무효화되게 한다.
        /// 파일이 바뀌면 다른 파일명을 찾게 되어 자동으로 다시 추출한다.
        /// </summary>
        private Bitmap LoadPerFilePath(string filePath)
        {
            long lastWriteTicks;
            long length;
            try
            {
                var info = new FileInfo(filePath);
                if (!info.Exists)
                {
                    return null;
                }

                lastWriteTicks = info.LastWriteTimeUtc.Ticks;
                length = info.Length;
            }
            catch
            {
                return null;
            }

            string key = CacheStorage.ComputeKey(filePath);
            if (key == null)
            {
                return ExtractImage(filePath);
            }

            // 바로가기 아이콘은 대상 프로그램에서 오므로, 대상이 바뀌어도 바로가기 자체의 수정 시각은 그대로다.
            // 그래서 바로가기는 하루 단위 파일명을 써서 대상 아이콘 변경이 하루 안에 반영되게 한다.
            string bucket = IsShortcut(filePath)
                ? "_" + DateTime.UtcNow.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                : string.Empty;

            string fileName = string.Format(
                CultureInfo.InvariantCulture,
                "{0}_{1}_{2}_{3}{4}{5}",
                FilePrefix, key, lastWriteTicks, length, bucket, PngExtension);

            Bitmap image = TryLoadImage(fileName);
            if (image != null)
            {
                return image;
            }

            image = ExtractImage(filePath);
            if (image != null)
            {
                TrySaveImage(image, fileName);
                DeleteStaleEntries(key, fileName);
            }

            return image;
        }

        /// <summary>
        /// 일반 파일은 확장자 단위로 한 번만 추출한다. 기본 연결 프로그램이 바뀔 수 있으므로 하루 단위 파일명을 쓴다
        /// (오래된 파일은 정리 대상이며, 최대 하루 지난 아이콘만 보일 수 있다).
        /// </summary>
        private Bitmap LoadPerExtension(string extension, string filePath)
        {
            string fileName = string.Format(
                CultureInfo.InvariantCulture,
                "{0}_{1}_{2}{3}",
                ExtensionPrefix, SanitizeExtension(extension), DateTime.UtcNow.ToString("yyyyMMdd", CultureInfo.InvariantCulture), PngExtension);

            Bitmap image = TryLoadImage(fileName);
            if (image != null)
            {
                return image;
            }

            image = ExtractImage(filePath);
            if (image != null)
            {
                TrySaveImage(image, fileName);
            }

            return image;
        }

        private static Bitmap ExtractImage(string filePath)
        {
            try
            {
                using (var icon = IconLoader.LoadFileIcon(filePath))
                {
                    return icon == null ? null : icon.ToBitmap();
                }
            }
            catch
            {
                return null;
            }
        }

        private Bitmap TryLoadImage(string fileName)
        {
            if (_directory == null)
            {
                return null;
            }

            string path;
            try
            {
                path = Path.Combine(_directory, fileName);
            }
            catch
            {
                return null;
            }

            try
            {
                if (!File.Exists(path))
                {
                    return null;
                }

                byte[] data = File.ReadAllBytes(path);
                using (var stream = new MemoryStream(data))
                using (var image = Image.FromStream(stream))
                {
                    return new Bitmap(image);
                }
            }
            catch
            {
                // 손상되었거나 읽을 수 없는 캐시 파일은 지우고 다시 만든다.
                CacheStorage.TryDelete(path);
                return null;
            }
        }

        private void TrySaveImage(Bitmap image, string fileName)
        {
            if (_directory == null || image == null)
            {
                return;
            }

            try
            {
                using (var stream = new MemoryStream())
                {
                    image.Save(stream, ImageFormat.Png);
                    CacheStorage.TryWrite(Path.Combine(_directory, fileName), stream.ToArray());
                }
            }
            catch
            {
                // 캐시 저장 실패는 무시한다.
            }
        }

        /// <summary>같은 파일의 이전 상태로 만들어 둔 캐시 파일을 지운다.</summary>
        private void DeleteStaleEntries(string key, string currentFileName)
        {
            if (_directory == null)
            {
                return;
            }

            try
            {
                foreach (string file in Directory.EnumerateFiles(_directory, FilePrefix + "_" + key + "_*" + PngExtension))
                {
                    if (!string.Equals(Path.GetFileName(file), currentFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        CacheStorage.TryDelete(file);
                    }
                }
            }
            catch
            {
                // 정리 실패는 무시한다.
            }
        }

        private static string SanitizeExtension(string extension)
        {
            if (string.IsNullOrEmpty(extension))
            {
                return "none";
            }

            string trimmed = extension.TrimStart('.');
            if (trimmed.Length == 0)
            {
                return "none";
            }

            foreach (char invalid in Path.GetInvalidFileNameChars())
            {
                trimmed = trimmed.Replace(invalid, '_');
            }

            return trimmed.ToLowerInvariant();
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            foreach (var pair in _memory)
            {
                if (pair.Value != null)
                {
                    pair.Value.Dispose();
                }
            }

            _memory.Clear();
            _disposed = true;
        }
    }
}
