using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using QuickLaunchOnTray.Models;

namespace QuickLaunchOnTray.Services
{
    public class ConfigurationService
    {
        private readonly string _iniPath;
        private readonly LocalizationService _localizationService;

        public ConfigurationService()
        {
            _iniPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini");
            _localizationService = LocalizationService.Instance;
        }

        public List<ProgramItem> LoadProgramItems()
        {
            if (!File.Exists(_iniPath))
            {
                throw new FileNotFoundException(
                    _localizationService.GetString("ConfigFileNotFound", _iniPath),
                    _iniPath);
            }

            List<ProgramItem> items = new List<ProgramItem>();
            bool inProgramsSection = false;

            try
            {
                foreach (string line in File.ReadAllLines(_iniPath))
                {
                    string trimmed = line.Trim();

                    if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith(";") || trimmed.StartsWith("#"))
                        continue;

                    if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                    {
                        string section = trimmed.Substring(1, trimmed.Length - 2);
                        inProgramsSection = section.Equals("Programs", StringComparison.OrdinalIgnoreCase);
                        continue;
                    }

                    if (inProgramsSection)
                    {
                        if (trimmed.Contains("="))
                        {
                            int idx = trimmed.IndexOf('=');
                            string key = trimmed.Substring(0, idx).Trim();
                            string value = trimmed.Substring(idx + 1).Trim();

                            if (!string.IsNullOrEmpty(value))
                            {
                                ValidatePath(value);
                                items.Add(new ProgramItem { Name = key, Path = value });
                            }
                        }
                        else
                        {
                            ValidatePath(trimmed);
                            string fileName = Path.GetFileNameWithoutExtension(trimmed);
                            items.Add(new ProgramItem { Name = fileName, Path = trimmed });
                        }
                    }
                }

                if (items.Count == 0)
                {
                    throw new InvalidOperationException(
                        _localizationService.GetString("NoProgramInfo"));
                }

                return items;
            }
            catch (Exception ex) when (!(ex is FileNotFoundException) && !(ex is InvalidOperationException))
            {
                throw new InvalidOperationException(
                    _localizationService.GetString("ConfigReadError"),
                    ex);
            }
        }

        private void ValidatePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new InvalidOperationException(
                    _localizationService.GetString("EmptyPath"));
            }

            if (!Directory.Exists(path) && !File.Exists(path))
            {
                throw new InvalidOperationException(
                    _localizationService.GetString("PathNotExist", path));
            }
        }
    }
} 