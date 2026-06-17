using System;
using System.Windows.Forms;

namespace SummerPractice
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            string login = txtLogin.Text.Trim();
            string password = txtPassword.Text;

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
                    MessageBox.Show("Пользователь не найден. Зарегистрируйтесь.", "Ошибка входа", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Читаем сохранённый хеш из поля "Пароль"
                string storedHash = userTable.Rows[0]["Пароль"].ToString();

                // Сравниваем введённый пароль с хешем через SHA256
                if (DatabaseManager.VerifyPassword(password, storedHash))
                {
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
                MessageBox.Show($"Ошибка подключения к БД:\n{ex.Message}", "Критическая ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void LoginForm_Load(object sender, EventArgs e)
        {

        }
    }
}
