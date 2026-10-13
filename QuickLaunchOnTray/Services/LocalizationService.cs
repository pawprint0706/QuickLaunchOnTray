using System;
using System.Globalization;
using QuickLaunchOnTray.Resources;

namespace QuickLaunchOnTray.Services
{
    /// <summary>
    /// 리소스 기반 문자열 조회. 조회 문화권은 <see cref="Strings.Culture"/> 하나만 사용하며
    /// 키는 언어 파일(중립 Strings.resx / 위성 Strings.&lt;culture&gt;.resx)에 동일하게 정의한다.
    /// </summary>
    public class LocalizationService
    {
        private static readonly Lazy<LocalizationService> _instance =
            new Lazy<LocalizationService>(() => new LocalizationService());

        public static LocalizationService Instance
        {
            get { return _instance.Value; }
        }

        private LocalizationService()
        {
            Strings.Culture = CultureInfo.CurrentUICulture;
        }

        /// <summary>현재 UI 문화권에 해당하는 문자열을 반환한다. 키가 없으면 키 자체를 반환한다.</summary>
        public string GetString(string key, params object[] args)
        {
            string format = Strings.ResourceManager.GetString(key, Strings.Culture);

            if (string.IsNullOrEmpty(format))
            {
                return key;
            }

            return (args != null && args.Length > 0) ? string.Format(Strings.Culture, format, args) : format;
        }
    }
}
