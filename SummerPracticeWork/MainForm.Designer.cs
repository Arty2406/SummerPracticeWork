using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace SummerPractice
{
    partial class MainForm : Form
    {
        private IContainer components = null;

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
            lblWelcome = new Label();
            btnLogout = new Button();
            comboBoxTables = new ComboBox();
            dataGridViewMain = new DataGridView();
            ((ISupportInitialize)dataGridViewMain).BeginInit();
            SuspendLayout();
            // 
            // lblWelcome
            // 
            lblWelcome.AutoSize = true;
            lblWelcome.Location = new Point(20, 20);
            lblWelcome.Name = "lblWelcome";
            lblWelcome.Size = new Size(147, 20);
            lblWelcome.TabIndex = 0;
            lblWelcome.Text = "Добро пожаловать!";
            // 
            // btnLogout
            // 
            btnLogout.Location = new Point(20, 50);
            btnLogout.Name = "btnLogout";
            btnLogout.Size = new Size(100, 30);
            btnLogout.TabIndex = 1;
            btnLogout.Text = "Выйти";
            btnLogout.UseVisualStyleBackColor = true;
            // 
            // comboBoxTables
            // 
            comboBoxTables.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxTables.FormattingEnabled = true;
            comboBoxTables.Location = new Point(20, 86);
            comboBoxTables.Name = "comboBoxTables";
            comboBoxTables.Size = new Size(151, 28);
            comboBoxTables.TabIndex = 2;
            comboBoxTables.SelectedIndexChanged += ComboBoxTables_SelectedIndexChanged;
            // 
            // dataGridViewMain
            // 
            dataGridViewMain.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewMain.Location = new Point(20, 120);
            dataGridViewMain.Name = "dataGridViewMain";
            dataGridViewMain.RowHeadersWidth = 51;
            dataGridViewMain.Size = new Size(300, 188);
            dataGridViewMain.TabIndex = 3;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(900, 600);
            Controls.Add(dataGridViewMain);
            Controls.Add(comboBoxTables);
            Controls.Add(lblWelcome);
            Controls.Add(btnLogout);
            Name = "MainForm";
            Text = "Главное меню";
            Load += MainForm_Load;
            ((ISupportInitialize)dataGridViewMain).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }
        #endregion

        // Обязательные поля (иначе элементы не будут доступны в MainForm.cs)
        private System.Windows.Forms.Label lblWelcome;
        private System.Windows.Forms.Button btnLogout;
        private ComboBox comboBoxTables;
        private DataGridView dataGridViewMain;
    }
}
