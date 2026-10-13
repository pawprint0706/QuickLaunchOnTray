using System.Globalization;
using System.Resources;

namespace QuickLaunchOnTray.Resources
{
    /// <summary>
    /// 문자열 리소스 조회 진입점.
    /// 중립 리소스는 Strings.resx(영어), 지역화 리소스는 Strings.&lt;culture&gt;.resx 이며
    /// 빌드 시 위성 어셈블리(예: ko\QuickLaunchOnTray.resources.dll)로 생성된다.
    /// 키는 언어별로 중복 정의하지 않고 동일 키를 각 언어 파일에서 각각 정의한다.
    /// </summary>
    internal static class Strings
    {
        private static readonly ResourceManager _resourceManager =
            new ResourceManager("QuickLaunchOnTray.Resources.Strings", typeof(Strings).Assembly);

        /// <summary>이 클래스가 사용하는 ResourceManager.</summary>
        public static ResourceManager ResourceManager
        {
            get { return _resourceManager; }
        }

        /// <summary>리소스 조회에 사용할 UI 문화권. 프로세스 시작 시점의 CurrentUICulture가 기본값이다.</summary>
        public static CultureInfo Culture { get; set; } = CultureInfo.CurrentUICulture;
    }
}
