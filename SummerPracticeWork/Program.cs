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

            // поддержка высокого DPI
            Application.SetHighDpiMode(HighDpiMode.SystemAware);

            // проверка на существование админа
            DatabaseManager.EnsureAdminCreated();

            Application.Run(new MainForm());
        }
    }
}