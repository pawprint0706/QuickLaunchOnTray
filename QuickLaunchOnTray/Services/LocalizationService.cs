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
            var property = typeof(Strings).GetProperty(resourceKey);
            
            if (property == null)
            {
                return key;
            }

            string format = (string)property.GetValue(null);
            return args.Length > 0 ? string.Format(format, args) : format;
        }
    }
} 