using System;
using System.Data;
using System.Windows.Forms;

namespace SummerPractice
{
    public partial class Autentification : Form
    {
        public Autentification()
        {
            InitializeComponent();
            ConfigurePasswordInput();
        }

        private void ConfigurePasswordInput()
        {
            // Маскируем вводимый пароль в текстовом поле
            txtPassword.UseSystemPasswordChar = true;

            // Настройка отправки формы по нажатию Enter
            txtLogin.KeyDown += TextBox_KeyDown;
            txtPassword.KeyDown += TextBox_KeyDown;
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // Отключаем стандартный писк Windows при Enter
                btnAutentification.PerformClick(); // Имитируем клик по кнопке входа
            }
        }

        private void btnAutentification_Click(object sender, EventArgs e)
        {
            string login = txtLogin.Text.Trim();
            string password = txtPassword.Text; // Пароль не триммим, пробелы могут быть частью пароля

            // Проверка на пустые поля
            if (string.IsNullOrWhiteSpace(login) || string.IsNullOrWhiteSpace(password))
            {
                MessageBox.Show("Заполните логин и пароль.", "Ошибка ввода",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (password.Length < 3)
            {
                MessageBox.Show("Пароль должен содержать не менее 3 symbols.", "Ошибка ввода",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Оборачиваем DataTable в using для своевременного освобождения памяти
                using (DataTable userTable = DatabaseManager.GetUserByLogin(login))
                {
                    if (userTable == null || userTable.Rows.Count == 0)
                    {
                        MessageBox.Show("Пользователь не найден. Зарегистрируйтесь.", "Ошибка входа",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    DataRow row = userTable.Rows[0];
                    string storedHash = row["Пароль"]?.ToString();
                    string role = row["Роль"]?.ToString() ?? "Гость";

                    if (DatabaseManager.VerifyPassword(password, storedHash))
                    {
                        // Сохраняем сессию вошедшего пользователя
                        CurrentUser.SetUser(login, role);

                        this.DialogResult = DialogResult.OK;
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Неверный логин или пароль.", "Ошибка входа",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка авторизации:\n{ex.Message}", "Критическая ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Autentification_Load(object sender, EventArgs e)
        {

        }
    }
}