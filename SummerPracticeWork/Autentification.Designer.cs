using System.Windows.Forms;

namespace SummerPractice
{
    partial class Autentification : Form
    {
        private System.ComponentModel.IContainer components = null;

        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            txtLogin = new TextBox();
            txtPassword = new TextBox();
            btnAutentification = new Button();
            labelLogin = new Label();
            labelPassword = new Label();
            labelHello = new Label();
            SuspendLayout();
            // 
            // txtLogin
            // 
            txtLogin.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtLogin.Location = new Point(316, 182);
            txtLogin.Name = "txtLogin";
            txtLogin.Size = new Size(225, 27);
            txtLogin.TabIndex = 0;
            // 
            // txtPassword
            // 
            txtPassword.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtPassword.Location = new Point(316, 215);
            txtPassword.Name = "txtPassword";
            txtPassword.Size = new Size(225, 27);
            txtPassword.TabIndex = 1;
            // 
            // btnAutentification
            // 
            btnAutentification.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnAutentification.Location = new Point(245, 248);
            btnAutentification.Name = "btnAutentification";
            btnAutentification.Size = new Size(296, 29);
            btnAutentification.TabIndex = 2;
            btnAutentification.Text = "Войти";
            btnAutentification.UseVisualStyleBackColor = true;
            btnAutentification.Click += btnAutentification_Click;
            // 
            // labelLogin
            // 
            labelLogin.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            labelLogin.AutoSize = true;
            labelLogin.Location = new Point(245, 185);
            labelLogin.Name = "labelLogin";
            labelLogin.Size = new Size(55, 20);
            labelLogin.TabIndex = 3;
            labelLogin.Text = "Логин:";
            // 
            // labelPassword
            // 
            labelPassword.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            labelPassword.AutoSize = true;
            labelPassword.Location = new Point(245, 218);
            labelPassword.Name = "labelPassword";
            labelPassword.Size = new Size(65, 20);
            labelPassword.TabIndex = 4;
            labelPassword.Text = "Пароль:";
            // 
            // labelHello
            // 
            labelHello.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            labelHello.AutoSize = true;
            labelHello.Font = new Font("Segoe UI", 15F);
            labelHello.Location = new Point(143, 73);
            labelHello.Name = "labelHello";
            labelHello.Size = new Size(532, 70);
            labelHello.TabIndex = 5;
            labelHello.Text = "Добро пожаловать, пользователь!\r\nВведите логин и пароль для входа в систему\r\n";
            labelHello.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // Autentification
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.LightSkyBlue;
            ClientSize = new Size(800, 450);
            Controls.Add(labelHello);
            Controls.Add(labelPassword);
            Controls.Add(labelLogin);
            Controls.Add(btnAutentification);
            Controls.Add(txtPassword);
            Controls.Add(txtLogin);
            Name = "Autentification";
            Text = "Вход в систему";
            Load += Autentification_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox txtLogin;
        private TextBox txtPassword;
        private Button btnAutentification;
        private Label labelLogin;
        private Label labelPassword;
        private Label labelHello;
    }
}