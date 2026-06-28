using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace SummerPractice
{
    partial class ProjectForm : Form
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
            btnExitToMain = new Button();
            comboBoxTables = new ComboBox();
            btnFilter = new Button();
            btnSort = new Button();
            btnCreateReport = new Button();
            btnSave = new Button();
            dataGridViewMain = new DataGridView();
            btnChange = new Button();
            groupBoxElemProject = new GroupBox();
            btnSearch = new Button();
            ((ISupportInitialize)dataGridViewMain).BeginInit();
            groupBoxElemProject.SuspendLayout();
            SuspendLayout();
            // 
            // btnExitToMain
            // 
            btnExitToMain.Anchor = AnchorStyles.Right;
            btnExitToMain.Location = new Point(18, 244);
            btnExitToMain.Margin = new Padding(3, 2, 3, 2);
            btnExitToMain.Name = "btnExitToMain";
            btnExitToMain.Size = new Size(164, 21);
            btnExitToMain.TabIndex = 1;
            btnExitToMain.Text = "Выйти в главное меню";
            btnExitToMain.UseVisualStyleBackColor = true;
            btnExitToMain.Click += btnExitToMain_Click;
            // 
            // comboBoxTables
            // 
            comboBoxTables.Anchor = AnchorStyles.Right;
            comboBoxTables.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxTables.FormattingEnabled = true;
            comboBoxTables.Location = new Point(18, 66);
            comboBoxTables.Margin = new Padding(3, 2, 3, 2);
            comboBoxTables.Name = "comboBoxTables";
            comboBoxTables.Size = new Size(164, 23);
            comboBoxTables.TabIndex = 2;
            comboBoxTables.SelectedIndexChanged += ComboBoxTables_SelectedIndexChanged;
            // 
            // btnFilter
            // 
            btnFilter.Anchor = AnchorStyles.Right;
            btnFilter.Location = new Point(18, 117);
            btnFilter.Margin = new Padding(3, 2, 3, 2);
            btnFilter.Name = "btnFilter";
            btnFilter.Size = new Size(164, 21);
            btnFilter.TabIndex = 4;
            btnFilter.Text = "Фильтр";
            btnFilter.UseVisualStyleBackColor = true;
            btnFilter.Click += btnFilter_Click;
            // 
            // btnSort
            // 
            btnSort.Anchor = AnchorStyles.Right;
            btnSort.Location = new Point(18, 142);
            btnSort.Margin = new Padding(3, 2, 3, 2);
            btnSort.Name = "btnSort";
            btnSort.Size = new Size(164, 21);
            btnSort.TabIndex = 5;
            btnSort.Text = "Сортировка";
            btnSort.UseVisualStyleBackColor = true;
            btnSort.Click += btnSort_Click;
            // 
            // btnCreateReport
            // 
            btnCreateReport.Anchor = AnchorStyles.Right;
            btnCreateReport.Location = new Point(18, 168);
            btnCreateReport.Margin = new Padding(3, 2, 3, 2);
            btnCreateReport.Name = "btnCreateReport";
            btnCreateReport.Size = new Size(164, 21);
            btnCreateReport.TabIndex = 6;
            btnCreateReport.Text = "Создать отчёт (Word)";
            btnCreateReport.UseVisualStyleBackColor = true;
            btnCreateReport.Click += btnCreateReport_Click;
            // 
            // btnSave
            // 
            btnSave.Anchor = AnchorStyles.Right;
            btnSave.Location = new Point(18, 219);
            btnSave.Margin = new Padding(3, 2, 3, 2);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(164, 21);
            btnSave.TabIndex = 8;
            btnSave.Text = "Сохранить";
            btnSave.UseVisualStyleBackColor = true;
            btnSave.Click += btnSave_Click;
            // 
            // dataGridViewMain
            // 
            dataGridViewMain.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dataGridViewMain.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridViewMain.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridViewMain.BackgroundColor = Color.LightSkyBlue;
            dataGridViewMain.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewMain.Location = new Point(18, 9);
            dataGridViewMain.Margin = new Padding(3, 2, 3, 2);
            dataGridViewMain.Name = "dataGridViewMain";
            dataGridViewMain.RowHeadersWidth = 51;
            dataGridViewMain.Size = new Size(556, 329);
            dataGridViewMain.TabIndex = 3;
            dataGridViewMain.CellValidating += DataGridViewMain_CellValidating;
            dataGridViewMain.UserDeletingRow += DataGridViewMain_UserDeletingRow;
            dataGridViewMain.KeyDown += DataGridViewMain_KeyDown;
            // 
            // btnChange
            // 
            btnChange.Anchor = AnchorStyles.Right;
            btnChange.Location = new Point(18, 194);
            btnChange.Margin = new Padding(3, 2, 3, 2);
            btnChange.Name = "btnChange";
            btnChange.Size = new Size(164, 21);
            btnChange.TabIndex = 9;
            btnChange.Text = "Изменить таблицу";
            btnChange.UseVisualStyleBackColor = true;
            btnChange.Click += btnChange_Click;
            // 
            // groupBoxElemProject
            // 
            groupBoxElemProject.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            groupBoxElemProject.Controls.Add(btnSearch);
            groupBoxElemProject.Controls.Add(comboBoxTables);
            groupBoxElemProject.Controls.Add(btnChange);
            groupBoxElemProject.Controls.Add(btnExitToMain);
            groupBoxElemProject.Controls.Add(btnFilter);
            groupBoxElemProject.Controls.Add(btnSave);
            groupBoxElemProject.Controls.Add(btnSort);
            groupBoxElemProject.Controls.Add(btnCreateReport);
            groupBoxElemProject.Location = new Point(579, 9);
            groupBoxElemProject.Margin = new Padding(3, 2, 3, 2);
            groupBoxElemProject.Name = "groupBoxElemProject";
            groupBoxElemProject.Padding = new Padding(3, 2, 3, 2);
            groupBoxElemProject.Size = new Size(198, 329);
            groupBoxElemProject.TabIndex = 10;
            groupBoxElemProject.TabStop = false;
            // 
            // btnSearch
            // 
            btnSearch.Anchor = AnchorStyles.Right;
            btnSearch.Location = new Point(18, 92);
            btnSearch.Margin = new Padding(3, 2, 3, 2);
            btnSearch.Name = "btnSearch";
            btnSearch.Size = new Size(164, 21);
            btnSearch.TabIndex = 10;
            btnSearch.Text = "Поиск";
            btnSearch.UseVisualStyleBackColor = true;
            btnSearch.Click += btnSearch_Click;
            // 
            // ProjectForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            BackColor = Color.LightSkyBlue;
            ClientSize = new Size(788, 450);
            Controls.Add(groupBoxElemProject);
            Controls.Add(dataGridViewMain);
            Margin = new Padding(3, 2, 3, 2);
            Name = "ProjectForm";
            Text = "Система для работы с таблицами Access";
            Load += ProjectForm_Load;
            ((ISupportInitialize)dataGridViewMain).EndInit();
            groupBoxElemProject.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion
        private System.Windows.Forms.Button btnExitToMain;
        private ComboBox comboBoxTables;
        private Button btnFilter;
        private Button btnSort;
        private Button btnCreateReport;
        private Button btnSave;
        private DataGridView dataGridViewMain;
        private Button btnChange;
        private GroupBox groupBoxElemProject;
        private Button btnSearch;
    }
}
