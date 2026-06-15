using System;
using System.Windows.Forms;

namespace SummerPractice
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();
            this.Text = "Вход в систему";
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            string login = txtLogin.Text.Trim();
            string password = txtPassword.Text;

            // Простая валидация длины (n символов, пусть будет 3)
            if (password.Length < 3)
            {
                MessageBox.Show("Пароль должен содержать не менее 3 символов.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var userTable = DatabaseManager.GetUserByLogin(login);

                if (userTable.Rows.Count == 0)
                {
                    MessageBox.Show("Пользователь не найден. Если вы новый пользователь, зарегистрируйтесь.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Получаем данные из первой строки результата
                string storedHash = userTable.Rows[0]["PasswordHash"].ToString();
                string storedSalt = userTable.Rows[0]["Salt"].ToString();

                if (DatabaseManager.VerifyPassword(password, storedHash, storedSalt))
                {
                    // Успешный вход
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Неверный логин или пароль.", "Ошибка входа", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                // Сюда попадет ошибка драйвера (Microsoft ACE OLEDB)
                MessageBox.Show($"Критическая ошибка подключения к БД:\n{ex.Message}\n\nУбедитесь, что файл CourseWork.accdb лежит в папке bin и установлен драйвер Access.", "Ошибка системы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit(); // Закрываем приложение, так как без БД работать нельзя
            }
        }
    }
}
