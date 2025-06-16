using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using QuickLaunchOnTray.Models;
using QuickLaunchOnTray.Services;

namespace QuickLaunchOnTray
{
    public class TrayApplicationContext : ApplicationContext, IDisposable
    {
        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool DestroyIcon(IntPtr handle);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        private readonly List<NotifyIcon> _trayIcons = new List<NotifyIcon>();
        private readonly Icon _folderIcon;
        private readonly Form _hiddenForm;
        private readonly ConfigurationService _configService;
        private readonly LocalizationService _localizationService;
        private bool _disposed;

        public TrayApplicationContext()
        {
            _configService = new ConfigurationService();
            _localizationService = LocalizationService.Instance;

            // 폴더 아이콘 초기화
            try
            {
                string systemFolder = Environment.GetFolderPath(Environment.SpecialFolder.System);
                IntPtr hIcon = ExtractIcon(IntPtr.Zero, Path.Combine(systemFolder, "shell32.dll"), 3);
                if (hIcon != IntPtr.Zero)
                {
                    _folderIcon = Icon.FromHandle(hIcon);
                }
                else
                {
                    _folderIcon = SystemIcons.Application;
                }
            }
            catch
            {
                _folderIcon = SystemIcons.Application;
            }

            try
            {
                // 프로그램 정보 로드 및 트레이 아이콘 생성
                var programItems = _configService.LoadProgramItems();
                foreach (var item in programItems)
                {
                    CreateTrayIconForProgram(item);
                }

                // 숨겨진 폼 초기화
                _hiddenForm = new Form
                {
                    FormBorderStyle = FormBorderStyle.None,
                    ShowInTaskbar = false,
                    StartPosition = FormStartPosition.Manual,
                    Size = new Size(1, 1),
                    Location = new Point(-2000, -2000),
                    Opacity = 0
                };
                _hiddenForm.Show();
                _hiddenForm.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, _localizationService.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1);
            }
        }

        private void CreateTrayIconForProgram(ProgramItem item)
        {
            Icon icon;
            try
            {
                if (Directory.Exists(item.Path))
                    icon = _folderIcon;
                else if (File.Exists(item.Path))
                    icon = Icon.ExtractAssociatedIcon(item.Path);
                else
                    icon = SystemIcons.Application;
            }
            catch
            {
                icon = SystemIcons.Application;
            }

            var notifyIcon = new NotifyIcon
            {
                Icon = icon,
                Text = item.Name,
                Visible = true
            };

            var menu = new ContextMenuStrip();
            string menuText = Directory.Exists(item.Path)
                ? _localizationService.GetString("OpenFolder")
                : _localizationService.GetString("RunProgram");

            var runItem = new ToolStripMenuItem(menuText);
            runItem.Click += (s, e) =>
            {
                try
                {
                    if (Directory.Exists(item.Path))
                    {
                        Process.Start("explorer.exe", item.Path);
                    }
                    else
                    {
                        RunProgram(item.Path);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        _localizationService.GetString("ProgramRunError", ex.Message),
                        _localizationService.GetString("Error"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            };
            menu.Items.Add(runItem);

            var exitItem = new ToolStripMenuItem(_localizationService.GetString("TerminateApp"));
            exitItem.Click += (s, e) =>
            {
                var result = MessageBox.Show(
                    _localizationService.GetString("ConfirmTermination"),
                    _localizationService.GetString("Confirm"),
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    ExitThread();
                }
            };
            menu.Items.Add(exitItem);

            notifyIcon.ContextMenuStrip = menu;
            _trayIcons.Add(notifyIcon);
        }

        private void RunProgram(string programPath)
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = programPath,
                    UseShellExecute = true
                };
                Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    _localizationService.GetString("CannotRunProgram", ex.Message),
                    ex);
            }
        }

        protected override void ExitThreadCore()
        {
            foreach (var icon in _trayIcons)
            {
                icon.Visible = false;
                icon.Dispose();
            }
            _trayIcons.Clear();

            if (_folderIcon != null)
            {
                _folderIcon.Dispose();
            }

            base.ExitThreadCore();
        }

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    foreach (var icon in _trayIcons)
                    {
                        icon.Dispose();
                    }
                    _trayIcons.Clear();

                    if (_folderIcon != null)
                    {
                        _folderIcon.Dispose();
                    }

                    if (_hiddenForm != null)
                    {
                        _hiddenForm.Dispose();
                    }
                }
                _disposed = true;
            }
            base.Dispose(disposing);
        }
    }
} 