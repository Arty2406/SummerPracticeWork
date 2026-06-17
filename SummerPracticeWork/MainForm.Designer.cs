namespace SummerPractice
{
    partial class MainForm
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
            labelTitle = new Label();
            btnRegistration = new Button();
            btnInput = new Button();
            labelChoice = new Label();
            btnOutput = new Button();
            SuspendLayout();
            // 
            // labelTitle
            // 
            labelTitle.Anchor = AnchorStyles.Top;
            labelTitle.AutoSize = true;
            labelTitle.Font = new Font("Segoe UI", 15F);
            labelTitle.Location = new Point(48, 93);
            labelTitle.Name = "labelTitle";
            labelTitle.Size = new Size(702, 35);
            labelTitle.TabIndex = 4;
            labelTitle.Text = "Система работы с таблицами Access файла \"Мероприятие\"";
            // 
            // btnRegistration
            // 
            btnRegistration.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnRegistration.Location = new Point(174, 240);
            btnRegistration.Name = "btnRegistration";
            btnRegistration.Size = new Size(458, 35);
            btnRegistration.TabIndex = 1;
            btnRegistration.Text = "Регистрация";
            btnRegistration.UseVisualStyleBackColor = true;
            btnRegistration.Click += btnRegistration_Click;
            // 
            // btnInput
            // 
            btnInput.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnInput.Location = new Point(174, 199);
            btnInput.Name = "btnInput";
            btnInput.Size = new Size(458, 35);
            btnInput.TabIndex = 0;
            btnInput.Text = "Вход в систему (для зарегистрированных пользователей)";
            btnInput.UseVisualStyleBackColor = true;
            btnInput.Click += btnInput_Click;
            // 
            // labelChoice
            // 
            labelChoice.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            labelChoice.Location = new Point(299, 165);
            labelChoice.Name = "labelChoice";
            labelChoice.Size = new Size(190, 31);
            labelChoice.TabIndex = 5;
            labelChoice.Text = "Выберите действие:";
            labelChoice.TextAlign = ContentAlignment.MiddleCenter;
            labelChoice.Click += labelChoice_Click;
            // 
            // btnOutput
            // 
            btnOutput.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnOutput.Location = new Point(174, 281);
            btnOutput.Name = "btnOutput";
            btnOutput.Size = new Size(458, 35);
            btnOutput.TabIndex = 2;
            btnOutput.Text = "Выйти из приложения";
            btnOutput.UseVisualStyleBackColor = true;
            btnOutput.Click += btnOutput_Click;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.LightSkyBlue;
            ClientSize = new Size(800, 450);
            Controls.Add(btnRegistration);
            Controls.Add(btnOutput);
            Controls.Add(labelChoice);
            Controls.Add(btnInput);
            Controls.Add(labelTitle);
            Name = "MainForm";
            Text = "Главное меню";
            Load += MainForm_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private Label labelTitle;
        private Button btnRegistration;
        private Button btnInput;
        private Label labelChoice;
        private Button btnOutput;
    }
}