using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace QuickLaunchOnTray
{
    static class Program
    {
        /// <summary>
        /// 해당 응용 프로그램의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // WinForms 초기화
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // TrayApplicationContext를 생성하여 실행 (폼 없이 트레이 아이콘만 표시)
            Application.Run(new TrayApplicationContext());
        }

        // 프로그램 정보를 담는 클래스
        public class ProgramItem
        {
            public string Name { get; set; }  // ini 파일의 key 값 또는 단순 경로인 경우 파일명(확장자 제외)
            public string Path { get; set; }  // 프로그램 경로
        }

        /// <summary>
        /// 폼 없이 트레이 아이콘들을 관리하는 ApplicationContext 클래스
        /// </summary>
        public class TrayApplicationContext : ApplicationContext
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

            // 생성한 트레이 아이콘들을 보관할 리스트
            private List<NotifyIcon> trayIcons = new List<NotifyIcon>();
            private Icon folderIcon;
            private Form hiddenForm;

            public TrayApplicationContext()
            {
                // 폴더 아이콘 초기화
                try
                {
                    string systemFolder = Environment.GetFolderPath(Environment.SpecialFolder.System);
                    IntPtr hIcon = ExtractIcon(IntPtr.Zero, Path.Combine(systemFolder, "shell32.dll"), 3); // 3은 폴더 아이콘의 인덱스
                    if (hIcon != IntPtr.Zero)
                    {
                        folderIcon = Icon.FromHandle(hIcon);
                    }
                    else
                    {
                        folderIcon = SystemIcons.Application;
                    }
                }
                catch
                {
                    folderIcon = SystemIcons.Application;
                }

                // 실행 파일과 같은 폴더의 config.ini 파일 경로 지정
                string iniPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini");
                if (!File.Exists(iniPath))
                {
                    MessageBox.Show("Could not find 'config.ini' file:\n" + iniPath);
                    Environment.Exit(0);
                    return;
                }

                // ini 파일에서 프로그램 정보(이름 및 경로) 읽어오기
                List<ProgramItem> programItems = LoadProgramItems(iniPath);
                if (programItems.Count == 0)
                {
                    MessageBox.Show("There is no information about the program to run in the 'config.ini' file.");
                    Environment.Exit(0);
                    return;
                }

                // 각 프로그램 정보에 대해 트레이 아이콘 생성
                foreach (ProgramItem item in programItems)
                {
                    CreateTrayIconForProgram(item);
                }

                hiddenForm = new Form();
                hiddenForm.FormBorderStyle = FormBorderStyle.None;
                hiddenForm.ShowInTaskbar = false;
                hiddenForm.StartPosition = FormStartPosition.Manual;
                hiddenForm.Size = new Size(1, 1);
                hiddenForm.Location = new Point(-2000, -2000); // 화면 밖
                hiddenForm.Opacity = 0;
                hiddenForm.Show();
                hiddenForm.Hide();
            }

            /// <summary>
            /// ini 파일에서 [Programs] 섹션 내의 프로그램 정보를 읽어 리스트로 반환합니다.
            /// key=value 형식인 경우 key를 이름으로 사용하며, 단순 경로인 경우 파일명(확장자 제외)을 이름으로 사용합니다.
            /// </summary>
            private List<ProgramItem> LoadProgramItems(string iniPath)
            {
                List<ProgramItem> items = new List<ProgramItem>();
                bool inProgramsSection = false;

                foreach (string line in File.ReadAllLines(iniPath))
                {
                    string trimmed = line.Trim();

                    // 빈 줄이나 주석은 무시
                    if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith(";") || trimmed.StartsWith("#"))
                        continue;

                    // 섹션 구분 [Programs] 등
                    if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                    {
                        string section = trimmed.Substring(1, trimmed.Length - 2);
                        inProgramsSection = section.Equals("Programs", StringComparison.OrdinalIgnoreCase);
                        continue;
                    }

                    if (inProgramsSection)
                    {
                        // key=value 형식인 경우 key는 이름, value는 경로로 사용
                        if (trimmed.Contains("="))
                        {
                            int idx = trimmed.IndexOf('=');
                            string key = trimmed.Substring(0, idx).Trim();
                            string value = trimmed.Substring(idx + 1).Trim();

                            if (!string.IsNullOrEmpty(value))
                                items.Add(new ProgramItem { Name = key, Path = value });
                        }
                        else
                        {
                            // 단순 경로인 경우 파일명(확장자 제외)을 이름으로 사용
                            string fileName = Path.GetFileNameWithoutExtension(trimmed);
                            items.Add(new ProgramItem { Name = fileName, Path = trimmed });
                        }
                    }
                }
                return items;
            }

            /// <summary>
            /// 지정한 프로그램 정보를 바탕으로 트레이 아이콘을 생성 및 등록합니다.
            /// </summary>
            private void CreateTrayIconForProgram(ProgramItem item)
            {
                // 아이콘 추출: 프로그램 파일이 존재하면 해당 아이콘을, 폴더인 경우 폴더 아이콘을, 그렇지 않으면 기본 아이콘 사용
                Icon icon;
                try
                {
                    if (Directory.Exists(item.Path))
                        icon = folderIcon;
                    else if (File.Exists(item.Path))
                        icon = Icon.ExtractAssociatedIcon(item.Path);
                    else
                        icon = SystemIcons.Application;
                }
                catch
                {
                    icon = SystemIcons.Application;
                }

                NotifyIcon notifyIcon = new NotifyIcon
                {
                    Icon = icon,
                    Text = item.Name,
                    Visible = true
                };

                // 컨텍스트 메뉴 생성
                ContextMenuStrip menu = new ContextMenuStrip();

                // "Run this program" 또는 "Open folder" 메뉴 항목
                string menuText = Directory.Exists(item.Path) ? "Open folder" : "Run this program";
                ToolStripMenuItem runItem = new ToolStripMenuItem(menuText);
                runItem.Click += (s, e) =>
                {
                    if (Directory.Exists(item.Path))
                    {
                        Process.Start("explorer.exe", item.Path);
                    }
                    else
                    {
                        RunProgram(item.Path);
                    }
                };
                menu.Items.Add(runItem);

                // "Terminate 'QuickLaunchOnTray'" 메뉴 항목
                ToolStripMenuItem exitItem = new ToolStripMenuItem("Terminate 'QuickLaunchOnTray'");
                exitItem.Click += (s, e) =>
                {
                    DialogResult result = MessageBox.Show("Are you sure you want to terminate 'QuickLaunchOnTray'? All quick launch icons will disappear.", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        ExitThread();
                    }
                };
                menu.Items.Add(exitItem);

                notifyIcon.ContextMenuStrip = menu;

                // 클릭 이벤트 처리
                notifyIcon.MouseClick += (s, e) =>
                {
                    if (e.Button == MouseButtons.Left)
                    {
                        if (Directory.Exists(item.Path))
                        {
                            ShowFolderContextMenuWithFiles(item.Path);
                        }
                        else
                        {
                            RunProgram(item.Path);
                        }
                    }
                };

                trayIcons.Add(notifyIcon);
            }

            private void ShowFolderContextMenuWithFiles(string folderPath)
            {
                try
                {
                    ContextMenuStrip folderMenu = new ContextMenuStrip();
                    BuildFolderAndFileMenu(folderPath, folderMenu.Items);
                    // 루트 메뉴에도 '탐색기에서 열기' 추가
                    if (folderMenu.Items.Count > 0)
                    {
                        folderMenu.Items.Add(new ToolStripSeparator());
                        var openInExplorer = new ToolStripMenuItem("탐색기에서 열기");
                        openInExplorer.Click += (s, e) => { Process.Start("explorer.exe", folderPath); };
                        folderMenu.Items.Add(openInExplorer);
                    }
                    if (folderMenu.Items.Count > 0)
                    {
                        POINT cursorPos;
                        GetCursorPos(out cursorPos);
                        folderMenu.Show(hiddenForm, hiddenForm.PointToClient(new Point(cursorPos.X, cursorPos.Y)));
                    }
                    else
                    {
                        Process.Start("explorer.exe", folderPath);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error showing folder menu: {ex.Message}");
                }
            }

            private void BuildFolderAndFileMenu(string folderPath, ToolStripItemCollection items)
            {
                try
                {
                    // 하위 폴더
                    string[] subFolders = Directory.GetDirectories(folderPath);
                    foreach (string subFolder in subFolders)
                    {
                        string folderName = Path.GetFileName(subFolder);
                        ToolStripMenuItem folderItem = new ToolStripMenuItem(folderName, folderIcon?.ToBitmap());
                        if (Directory.GetDirectories(subFolder).Length > 0 || Directory.GetFiles(subFolder).Length > 0)
                        {
                            BuildFolderAndFileMenu(subFolder, folderItem.DropDownItems);
                        }
                        // '탐색기에서 열기' 하위 메뉴 추가
                        var openInExplorer = new ToolStripMenuItem("탐색기에서 열기");
                        openInExplorer.Click += (s, e) => { Process.Start("explorer.exe", subFolder); };
                        folderItem.DropDownItems.Add(new ToolStripSeparator());
                        folderItem.DropDownItems.Add(openInExplorer);
                        items.Add(folderItem);
                    }
                    // 실행파일 및 바로가기 → 모든 파일 표시
                    var files = Directory.GetFiles(folderPath);
                    foreach (string file in files)
                    {
                        string fileName = Path.GetFileName(file);
                        Icon fileIcon = null;
                        try { fileIcon = Icon.ExtractAssociatedIcon(file); } catch { fileIcon = SystemIcons.Application; }
                        ToolStripMenuItem fileItem = new ToolStripMenuItem(fileName, fileIcon?.ToBitmap());
                        fileItem.Click += (s, e) => { RunProgram(file); };
                        items.Add(fileItem);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error building folder menu: {ex.Message}");
                }
            }

            /// <summary>
            /// 지정한 경로의 프로그램을 실행합니다.
            /// </summary>
            private void RunProgram(string programPath)
            {
                try
                {
                    Process.Start(programPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Run this program failed:\n" + ex.Message);
                }
            }

            /// <summary>
            /// ApplicationContext 종료 시 모든 트레이 아이콘을 정리합니다.
            /// </summary>
            protected override void ExitThreadCore()
            {
                // 모든 NotifyIcon을 숨기고 Dispose
                foreach (var icon in trayIcons)
                {
                    icon.Visible = false;
                    icon.Dispose();
                }

                // 폴더 아이콘 해제
                if (folderIcon != null)
                {
                    folderIcon.Dispose();
                }

                if (hiddenForm != null) hiddenForm.Dispose();

                base.ExitThreadCore();
            }
        }
    }
}

