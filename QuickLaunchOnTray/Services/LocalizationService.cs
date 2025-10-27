using System;
using System.Globalization;
using QuickLaunchOnTray.Resources;

namespace QuickLaunchOnTray.Services
{
    public class LocalizationService
    {
        private static readonly Lazy<LocalizationService> _instance = 
            new Lazy<LocalizationService>(() => new LocalizationService());

        public static LocalizationService Instance => _instance.Value;

        private LocalizationService()
        {
            // 리소스의 Culture를 시스템 언어에 맞게 설정
            Strings.Culture = CultureInfo.CurrentUICulture;
        }

        public bool IsKoreanSystem()
        {
            return CultureInfo.CurrentUICulture.TwoLetterISOLanguageName.Equals("ko", StringComparison.OrdinalIgnoreCase);
        }

        public string GetString(string key, params object[] args)
        {
            string resourceKey = IsKoreanSystem() ? $"{key}_ko" : key;

            // Try exact key for current locale suffix first
            string format = Strings.ResourceManager.GetString(resourceKey, Strings.Culture);

            // Fallback to base key if locale-suffixed key missing
            if (string.IsNullOrEmpty(format))
            {
                format = Strings.ResourceManager.GetString(key, Strings.Culture);
            }

            // Final fallback to the key itself
            if (string.IsNullOrEmpty(format))
            {
                return key;
            }

            return (args != null && args.Length > 0) ? string.Format(format, args) : format;
        }
    }
} 
