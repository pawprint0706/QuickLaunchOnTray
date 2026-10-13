using System.Collections.Generic;

namespace QuickLaunchOnTray.Models
{
    /// <summary>
    /// config.ini 파싱 결과. 등록 가능한 항목과, 건너뛴 항목에 대한 사용자 안내 문구를 담는다.
    /// </summary>
    public class ConfigurationLoadResult
    {
        public ConfigurationLoadResult()
        {
            Items = new List<ProgramItem>();
            Warnings = new List<string>();
        }

        /// <summary>등록 가능한 항목.</summary>
        public List<ProgramItem> Items { get; private set; }

        /// <summary>건너뛴 항목 등 사용자에게 알릴 경고 문구.</summary>
        public List<string> Warnings { get; private set; }
    }
}
