using System;
using System.Threading;
using System.Windows.Forms;
using QuickLaunchOnTray.Services;

namespace QuickLaunchOnTray
{
    internal static class Program
    {
        /// <summary>중복 실행 방지용 뮤텍스 이름(사용자 세션 단위).</summary>
        private const string SingleInstanceMutexName = @"Local\QuickLaunchOnTray.SingleInstance";

        [STAThread]
        private static void Main()
        {
            var localization = LocalizationService.Instance;

            var instanceMutex = new Mutex(false, SingleInstanceMutexName);
            bool ownsMutex = false;

            try
            {
                try
                {
                    ownsMutex = instanceMutex.WaitOne(TimeSpan.Zero, false);
                }
                catch (AbandonedMutexException)
                {
                    // 이전 인스턴스가 비정상 종료한 경우. 소유권을 넘겨받고 계속 실행한다.
                    ownsMutex = true;
                }

                if (!ownsMutex)
                {
                    MessageBox.Show(
                        localization.GetString("AlreadyRunning"),
                        localization.GetString("Warning"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                try
                {
                    Application.Run(new TrayApplicationContext());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        ex.Message,
                        localization.GetString("Error"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    Environment.ExitCode = 1;
                }
            }
            finally
            {
                if (ownsMutex)
                {
                    instanceMutex.ReleaseMutex();
                }

                instanceMutex.Dispose();
            }
        }
    }
}
