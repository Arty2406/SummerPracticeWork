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
        }

        private void btnAutentification_Click(object sender, EventArgs e)
        {
            string login = txtLogin.Text.Trim();
            string password = txtPassword.Text;

            // проверки
            if (string.IsNullOrWhiteSpace(login) || string.IsNullOrWhiteSpace(password))
            {
                MessageBox.Show("Заполните логин и пароль.", "Ошибка ввода");
                return;
            }

            if (password.Length < 3)
            {
                MessageBox.Show("Пароль должен содержать не менее 3 символов.", "Ошибка");
                return;
            }

            try
            {
                var userTable = DatabaseManager.GetUserByLogin(login);

                if (userTable.Rows.Count == 0)
                {
                    MessageBox.Show("Пользователь не найден. Зарегистрируйтесь.", "Ошибка входа",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                DataRow row = userTable.Rows[0];
                string storedHash = row["Пароль"].ToString();
                // чтение роли
                string role = row["Роль"].ToString();

                if (DatabaseManager.VerifyPassword(password, storedHash))
                {
                    // данные о вошедшем пользователе сохраняются
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
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка подключения к БД:\n{ex.Message}", "Критическая ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoginForm_Load(object sender, EventArgs e)
        {

        }
    }
}
