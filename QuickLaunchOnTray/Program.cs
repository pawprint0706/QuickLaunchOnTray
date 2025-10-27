using System;
using System.Windows.Forms;
using QuickLaunchOnTray.Services;

namespace QuickLaunchOnTray
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new TrayApplicationContext());
            }
            catch (Exception ex)
            {
                var localization = LocalizationService.Instance;
                MessageBox.Show(
                    localization.GetString("ProgramRunError", ex.Message),
                    localization.GetString("Error"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                Environment.Exit(1);
            }
        }
    }
}
