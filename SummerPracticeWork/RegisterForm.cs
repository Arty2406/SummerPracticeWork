using System;
using System.Windows.Forms;

namespace SummerPractice
{
    public partial class RegisterForm : Form
    {
        public RegisterForm()
        {
            InitializeComponent();
            this.Text = "Регистрация";
            this.StartPosition = FormStartPosition.CenterParent;
        }

        private void btnRegister_Click(object sender, EventArgs e)
        {
            string login = txtRegLogin.Text.Trim();
            string pass = txtRegPass.Text;
            string confirm = txtRegConfirm.Text;

            if (string.IsNullOrEmpty(login) || string.IsNullOrEmpty(pass))
            {
                MessageBox.Show("Заполните все поля.", "Ошибка");
                return;
            }

            if (pass.Length < 5) // n = 5 символов
            {
                MessageBox.Show("Пароль должен быть не менее 5 символов.", "Ошибка длины");
                return;
            }

            if (pass != confirm)
            {
                MessageBox.Show("Пароли не совпадают.", "Ошибка подтверждения");
                return;
            }

            try
            {
                DatabaseManager.RegisterUser(login, pass);
                MessageBox.Show("Регистрация успешна! Теперь войдите в систему.", "Успех");
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка регистрации: {ex.Message}", "Ошибка БД");
            }
        }

        private void RegisterForm_Load(object sender, EventArgs e)
        {
            btnRegister.Text = "Зарегистрироваться"; // Исправляем название кнопки
                                                     // Здесь можно добавить другие настройки при загрузке, если нужно
        }
    }
}
