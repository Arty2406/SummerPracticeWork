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

            Application.ThreadException += (sender, e) => {
                MessageBox.Show($"UI Error: {e.Exception.Message}\n{e.Exception.StackTrace}");
            };

            AppDomain.CurrentDomain.UnhandledException += (sender, e) => {
                MessageBox.Show($"Global Error: {((Exception)e.ExceptionObject).Message}");
            };

            Application.Run(new MainForm());
        }
    }
}