using System;
using System.Collections.Generic;
using System.IO;
using QuickLaunchOnTray.Models;

namespace QuickLaunchOnTray.Services
{
    /// <summary>
    /// 실행 파일과 같은 폴더의 config.ini에서 [Programs] 섹션 항목을 읽는다.
    /// 유효하지 않은 줄은 건너뛰고 경고로 수집하며, 등록 가능한 항목이 하나도 없을 때만 예외를 던진다.
    /// </summary>
    public class ConfigurationService
    {
        private const string ProgramsSectionName = "Programs";

        private readonly string _iniPath;
        private readonly LocalizationService _localizationService;

        public ConfigurationService()
        {
            _iniPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini");
            _localizationService = LocalizationService.Instance;
        }

        /// <summary>config.ini를 읽어 항목 목록과 경고 목록을 반환한다.</summary>
        /// <exception cref="FileNotFoundException">config.ini가 없을 때.</exception>
        /// <exception cref="InvalidOperationException">파일을 읽을 수 없거나 등록할 항목이 하나도 없을 때.</exception>
        public ConfigurationLoadResult LoadProgramItems()
        {
            if (!File.Exists(_iniPath))
            {
                throw new FileNotFoundException(
                    _localizationService.GetString("ConfigFileNotFound", _iniPath),
                    _iniPath);
            }

            string[] lines;
            try
            {
                lines = File.ReadAllLines(_iniPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    _localizationService.GetString("ConfigReadError"),
                    ex);
            }

            var result = new ConfigurationLoadResult();
            bool inProgramsSection = false;
            bool hasProgramsSection = false;

            for (int index = 0; index < lines.Length; index++)
            {
                string trimmed = lines[index].Trim();

                if (trimmed.Length == 0 || trimmed.StartsWith(";") || trimmed.StartsWith("#"))
                {
                    continue;
                }

                if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                {
                    inProgramsSection = trimmed.Substring(1, trimmed.Length - 2).Trim()
                        .Equals(ProgramsSectionName, StringComparison.OrdinalIgnoreCase);
                    hasProgramsSection |= inProgramsSection;
                    continue;
                }

                if (!inProgramsSection)
                {
                    continue;
                }

                AddEntry(result, trimmed, index + 1);
            }

            if (result.Items.Count == 0)
            {
                string message = _localizationService.GetString(hasProgramsSection ? "NoProgramInfo" : "NoProgramsSection");
                throw new InvalidOperationException(message + BuildWarningText(result.Warnings));
            }

            return result;
        }

        private void AddEntry(ConfigurationLoadResult result, string line, int lineNumber)
        {
            int separator = line.IndexOf('=');
            string name = separator >= 0 ? line.Substring(0, separator).Trim() : string.Empty;
            string path = (separator >= 0 ? line.Substring(separator + 1) : line).Trim();

            if (path.Length == 0)
            {
                AddSkippedWarning(result, lineNumber, _localizationService.GetString("EmptyPath"));
                return;
            }

            if (!Directory.Exists(path) && !File.Exists(path))
            {
                AddSkippedWarning(result, lineNumber, _localizationService.GetString("PathNotExist", path));
                return;
            }

            if (name.Length == 0)
            {
                name = ProgramItem.DeriveDisplayName(path);
            }

            result.Items.Add(new ProgramItem { Name = name, Path = path });
        }

        private void AddSkippedWarning(ConfigurationLoadResult result, int lineNumber, string reason)
        {
            result.Warnings.Add(_localizationService.GetString("ConfigEntrySkipped", lineNumber, reason));
        }

        private static string BuildWarningText(IReadOnlyList<string> warnings)
        {
            return warnings.Count == 0
                ? string.Empty
                : Environment.NewLine + Environment.NewLine + string.Join(Environment.NewLine, warnings);
        }
    }
}
