using System;
using System.Windows.Forms;

namespace SummerPractice
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetHighDpiMode(HighDpiMode.SystemAware);

            Application.ThreadException += (sender, args) =>
            {
                MessageBox.Show(
                    $"Ошибка приложения:\n\n{args.Exception.Message}\n\n{args.Exception.StackTrace}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };

            AppDomain.CurrentDomain.UnhandledException += (sender, args) =>
            {
                string msg = ((Exception)args.ExceptionObject)?.Message ?? "";
                MessageBox.Show($"Критическая ошибка:\n{msg}", "Критическая ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            };

            Application.Run(new MainForm());
        }
    }
}