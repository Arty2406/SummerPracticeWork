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
            txtRegLogin = new TextBox();
            txtRegPass = new TextBox();
            txtRegConfirm = new TextBox();
            btnRegister = new Button();
            SuspendLayout();
            // 
            // txtRegLogin
            // 
            txtRegLogin.Location = new Point(91, 88);
            txtRegLogin.Name = "txtRegLogin";
            txtRegLogin.Size = new Size(125, 27);
            txtRegLogin.TabIndex = 0;
            // 
            // txtRegPass
            // 
            txtRegPass.Location = new Point(91, 121);
            txtRegPass.Name = "txtRegPass";
            txtRegPass.Size = new Size(125, 27);
            txtRegPass.TabIndex = 1;
            // 
            // txtRegConfirm
            // 
            txtRegConfirm.Location = new Point(91, 154);
            txtRegConfirm.Name = "txtRegConfirm";
            txtRegConfirm.Size = new Size(125, 27);
            txtRegConfirm.TabIndex = 2;
            // 
            // btnRegister
            // 
            btnRegister.Location = new Point(91, 187);
            btnRegister.Name = "btnRegister";
            btnRegister.Size = new Size(125, 29);
            btnRegister.TabIndex = 3;
            btnRegister.Text = "button1";
            btnRegister.UseVisualStyleBackColor = true;
            // 
            // RegisterForm
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(btnRegister);
            Controls.Add(txtRegConfirm);
            Controls.Add(txtRegPass);
            Controls.Add(txtRegLogin);
            Name = "RegisterForm";
            Text = "RegisterForm";
            Load += this.RegisterForm_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox txtRegLogin;
        private TextBox txtRegPass;
        private TextBox txtRegConfirm;
        private Button btnRegister;
    }
}