namespace SummerPractice

{
    partial class RegisterForm : Form
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
            labelCorrectPassword = new Label();
            btnRegister = new Button();
            txtRegPass = new TextBox();
            labelLogin = new Label();
            txtRegLogin = new TextBox();
            labelPassword = new Label();
            txtRegConfirm = new TextBox();
            label1 = new Label();
            SuspendLayout();
            // 
            // labelCorrectPassword
            // 
            labelCorrectPassword.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            labelCorrectPassword.AutoSize = true;
            labelCorrectPassword.Location = new Point(86, 228);
            labelCorrectPassword.Name = "labelCorrectPassword";
            labelCorrectPassword.Size = new Size(157, 20);
            labelCorrectPassword.TabIndex = 8;
            labelCorrectPassword.Text = "Подтвердите пароль:";
            // 
            // btnRegister
            // 
            btnRegister.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            btnRegister.Location = new Point(86, 258);
            btnRegister.Name = "btnRegister";
            btnRegister.Size = new Size(628, 66);
            btnRegister.TabIndex = 3;
            btnRegister.Text = "Зарегистрироваться";
            btnRegister.UseVisualStyleBackColor = true;
            btnRegister.Click += btnRegister_Click;
            // 
            // txtRegPass
            // 
            txtRegPass.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            txtRegPass.Location = new Point(249, 225);
            txtRegPass.Name = "txtRegPass";
            txtRegPass.Size = new Size(465, 27);
            txtRegPass.TabIndex = 1;
            // 
            // labelLogin
            // 
            labelLogin.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            labelLogin.AutoSize = true;
            labelLogin.Location = new Point(86, 162);
            labelLogin.Name = "labelLogin";
            labelLogin.Size = new Size(113, 20);
            labelLogin.TabIndex = 6;
            labelLogin.Text = "Введите логин:";
            // 
            // txtRegLogin
            // 
            txtRegLogin.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            txtRegLogin.Location = new Point(249, 159);
            txtRegLogin.Name = "txtRegLogin";
            txtRegLogin.Size = new Size(465, 27);
            txtRegLogin.TabIndex = 0;
            // 
            // labelPassword
            // 
            labelPassword.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            labelPassword.AutoSize = true;
            labelPassword.Location = new Point(86, 195);
            labelPassword.Name = "labelPassword";
            labelPassword.Size = new Size(123, 20);
            labelPassword.TabIndex = 7;
            labelPassword.Text = "Введите пароль:";
            // 
            // txtRegConfirm
            // 
            txtRegConfirm.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            txtRegConfirm.Location = new Point(249, 192);
            txtRegConfirm.Name = "txtRegConfirm";
            txtRegConfirm.Size = new Size(465, 27);
            txtRegConfirm.TabIndex = 2;
            txtRegConfirm.TextChanged += txtRegConfirm_TextChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(86, 129);
            label1.Name = "label1";
            label1.Size = new Size(628, 20);
            label1.TabIndex = 9;
            label1.Text = "Для того чтобы получить доступ к функционалу программы нужно выполнить действия:\r\n";
            // 
            // RegisterForm
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.LightSkyBlue;
            ClientSize = new Size(800, 450);
            Controls.Add(label1);
            Controls.Add(labelCorrectPassword);
            Controls.Add(btnRegister);
            Controls.Add(labelLogin);
            Controls.Add(txtRegPass);
            Controls.Add(txtRegConfirm);
            Controls.Add(labelPassword);
            Controls.Add(txtRegLogin);
            Name = "RegisterForm";
            Text = "Регистрация";
            Load += RegisterForm_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private Label labelCorrectPassword;
        private Button btnRegister;
        private TextBox txtRegPass;
        private Label labelLogin;
        private TextBox txtRegLogin;
        private Label labelPassword;
        private TextBox txtRegConfirm;
        private Label label1;
    }
}