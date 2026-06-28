using System;
using System.Windows.Forms;

namespace SummerPractice
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.SetHighDpiMode(HighDpiMode.SystemAware);

                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

                Application.ThreadException += (sender, args) =>
                {
                    var ex = args.Exception;
                    if (ex is AccessViolationException)
                    {
                        System.Diagnostics.Debug.WriteLine($"AccessViolation suppressed: {ex.Message}");
                        return;
                    }

                    MessageBox.Show($"Ошибка:\n{ex?.Message}", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                };

                AppDomain.CurrentDomain.UnhandledException += (sender, args) =>
                {
                    var ex = args.ExceptionObject as Exception;
                    if (ex is AccessViolationException)
                    {
                        System.Diagnostics.Debug.WriteLine($"Unhandled AccessViolation: {ex.Message}");
                        return;
                    }

                    MessageBox.Show($"Критическая ошибка:\n{ex?.Message}", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                };

                // Инициализация БД
                try
                {
                    DatabaseManager.EnsureAdminCreated();
                }
                catch (AccessViolationException)
                {
                    System.Threading.Thread.Sleep(500);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    try
                    {
                        DatabaseManager.EnsureAdminCreated();
                    }
                    catch { }
                }
                catch { }

                Application.Run(new MainForm());
            }
            catch (Exception ex) when (ex is not AccessViolationException)
            {
                MessageBox.Show($"Ошибка запуска:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}