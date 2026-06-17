using System.Windows.Forms;

namespace SummerPractice
{
    partial class LoginForm : Form
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
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

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            txtLogin = new TextBox();
            txtPassword = new TextBox();
            btnLogin = new Button();
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
            // btnLogin
            // 
            btnLogin.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnLogin.Location = new Point(245, 248);
            btnLogin.Name = "btnLogin";
            btnLogin.Size = new Size(296, 29);
            btnLogin.TabIndex = 2;
            btnLogin.Text = "Войти";
            btnLogin.UseVisualStyleBackColor = true;
            btnLogin.Click += btnLogin_Click;
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
            // LoginForm
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.LightSkyBlue;
            ClientSize = new Size(800, 450);
            Controls.Add(labelHello);
            Controls.Add(labelPassword);
            Controls.Add(labelLogin);
            Controls.Add(btnLogin);
            Controls.Add(txtPassword);
            Controls.Add(txtLogin);
            Name = "LoginForm";
            Text = "Вход в систему";
            Load += LoginForm_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox txtLogin;
        private TextBox txtPassword;
        private Button btnLogin;
        private Label labelLogin;
        private Label labelPassword;
        private Label labelHello;
    }
}