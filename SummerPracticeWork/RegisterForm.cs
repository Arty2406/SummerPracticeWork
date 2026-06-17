using System;
using System.Windows.Forms;

namespace SummerPractice
{
    public partial class RegisterForm : Form
    {
        public RegisterForm()
        {
            InitializeComponent();
        }

        private void RegisterForm_Load(object sender, EventArgs e)
        {
            // фокусировка на поле с логином, звёздочки вместо символов при вводе
            txtRegPass.PasswordChar = '*';
            txtRegConfirm.PasswordChar = '*';
            txtRegLogin.Focus();
        }

        private void btnRegister_Click(object sender, EventArgs e)
        {
            string login = txtRegLogin.Text.Trim();
            string pass = txtRegPass.Text;
            string confirm = txtRegConfirm.Text;

            // проверки
            if (string.IsNullOrWhiteSpace(login))
            {
                MessageBox.Show("Введите логин.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtRegLogin.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(pass))
            {
                MessageBox.Show("Введите пароль.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtRegPass.Focus();
                return;
            }

            if (pass.Length < 6)
            {
                MessageBox.Show("Пароль должен быть не менее 5 символов.", "Ошибка длины", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtRegPass.Focus();
                return;
            }

            if (pass != confirm)
            {
                MessageBox.Show("Пароли не совпадают. Проверьте ввод.", "Ошибка подтверждения", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtRegConfirm.Clear();
                txtRegConfirm.Focus();
                return;
            }

            // запись пароля в БД
            try
            {
                // хеширование пароля через SHA256 и его сохранение в поле "Пароль"
                DatabaseManager.RegisterUser(login, pass);

                MessageBox.Show($"Пользователь '{login}' успешно зарегистрирован!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка регистрации: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtRegPass_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBoxReg_Enter(object sender, EventArgs e)
        {

        }

        private void txtRegConfirm_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
