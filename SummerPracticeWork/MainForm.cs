using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SummerPractice
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnInput_Click(object? sender, EventArgs e)
        {
            try { DatabaseManager.EnsureAdminCreated(); }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка инициализации БД: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var loginForm = new Autentification();
            if (loginForm.ShowDialog() != DialogResult.OK) return;

            using var projectForm = new ProjectForm();
            this.Hide();
            projectForm.ShowDialog();
            this.Show();
        }

        // кнопка регистрации
        private void btnRegistration_Click(object? sender, EventArgs e)
        {
            using (var regForm = new RegisterForm())
            {
                var result = regForm.ShowDialog();

                if (result == DialogResult.OK)
                {
                    MessageBox.Show("Регистрация успешна! Теперь войдите.", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // после регистрации пользователя сразу перекидывает на форму аутентификации
                    btnInput_Click(sender, e);
                }
            }
        }

        // кнопка выхода из приложения
        private void btnOutput_Click(object? sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "Вы уверены, что хотите выйти из приложения?",
                "Подтверждение выхода",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void labelChoice_Click(object sender, EventArgs e)
        {

        }
    }
}
