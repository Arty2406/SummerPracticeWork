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

            // Включаем поддержку высокого DPI, чтобы текст не был мыльным
            Application.SetHighDpiMode(HighDpiMode.SystemAware);

            // Запускаем сразу MainForm
            Application.Run(new MainForm());
        }
    }
}