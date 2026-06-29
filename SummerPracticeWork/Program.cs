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

            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

            Application.ThreadException += (sender, args) =>
            {
                string msg = args.Exception?.Message ?? "";

                // Подавляем ошибки ComboBox-биндинга — они не критичны
                if (args.Exception is ArgumentException ||
                    msg.Contains("DataGridViewComboBoxCell") ||
                    msg.Contains("not valid") ||
                    msg.Contains("value is not valid"))
                {
                    System.Diagnostics.Debug.WriteLine($"[Suppressed] {msg}");
                    return;
                }

                MessageBox.Show(
                    $"Ошибка приложения:\n\n{msg}\n\n{args.Exception?.StackTrace}",
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