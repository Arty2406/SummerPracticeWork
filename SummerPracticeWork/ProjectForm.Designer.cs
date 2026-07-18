namespace SummerPractice
{
    partial class ProjectForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            splitContainerMain = new SplitContainer();
            dataGridViewMain = new DataGridView();
            panelRight = new Panel();
            panelTopRight = new Panel();
            lblTable = new Label();
            comboBoxTables = new ComboBox();
            panelButtons = new Panel();
            btnCreateReport = new Button();
            btnFilter = new Button();
            btnSort = new Button();
            btnSearch = new Button();
            btnSave = new Button();
            btnChange = new Button();
            btnExitToMain = new Button();
            ((System.ComponentModel.ISupportInitialize)splitContainerMain).BeginInit();
            splitContainerMain.Panel1.SuspendLayout();
            splitContainerMain.Panel2.SuspendLayout();
            splitContainerMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridViewMain).BeginInit();
            panelRight.SuspendLayout();
            panelTopRight.SuspendLayout();
            panelButtons.SuspendLayout();
            SuspendLayout();
            // 
            // splitContainerMain
            // 
            splitContainerMain.Dock = DockStyle.Fill;
            splitContainerMain.Location = new Point(0, 0);
            splitContainerMain.Name = "splitContainerMain";
            // 
            // splitContainerMain.Panel1
            // 
            splitContainerMain.Panel1.BackColor = Color.LightSkyBlue;
            splitContainerMain.Panel1.Controls.Add(dataGridViewMain);
            splitContainerMain.Panel1.Padding = new Padding(10);
            splitContainerMain.Panel1MinSize = 500;
            // 
            // splitContainerMain.Panel2
            // 
            splitContainerMain.Panel2.BackColor = Color.LightSkyBlue;
            splitContainerMain.Panel2.Controls.Add(panelRight);
            splitContainerMain.Panel2.Padding = new Padding(10);
            splitContainerMain.Panel2MinSize = 300;
            splitContainerMain.Size = new Size(1283, 692);
            splitContainerMain.SplitterDistance = 850;
            splitContainerMain.TabIndex = 0;
            // 
            // dataGridViewMain
            // 
            dataGridViewMain.AllowUserToAddRows = false;
            dataGridViewMain.AllowUserToDeleteRows = false;
            dataGridViewMain.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridViewMain.BackgroundColor = Color.White;
            dataGridViewMain.BorderStyle = BorderStyle.Fixed3D;
            dataGridViewMain.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewMain.Dock = DockStyle.Fill;
            dataGridViewMain.Location = new Point(10, 10);
            dataGridViewMain.Name = "dataGridViewMain";
            dataGridViewMain.ReadOnly = true;
            dataGridViewMain.Size = new Size(830, 672);
            dataGridViewMain.TabIndex = 0;
            // 
            // panelRight
            // 
            panelRight.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            panelRight.BackColor = Color.LightSkyBlue;
            panelRight.Controls.Add(panelTopRight);
            panelRight.Controls.Add(panelButtons);
            panelRight.Controls.Add(btnExitToMain);
            panelRight.Location = new Point(0, 0);
            panelRight.Name = "panelRight";
            panelRight.Size = new Size(413, 672);
            panelRight.TabIndex = 0;
            // 
            // panelTopRight
            // 
            panelTopRight.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            panelTopRight.BackColor = Color.LightSkyBlue;
            panelTopRight.Controls.Add(lblTable);
            panelTopRight.Controls.Add(comboBoxTables);
            panelTopRight.Location = new Point(5, 5);
            panelTopRight.Name = "panelTopRight";
            panelTopRight.Size = new Size(403, 80);
            panelTopRight.TabIndex = 0;
            // 
            // lblTable
            // 
            lblTable.AutoSize = true;
            lblTable.Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold);
            lblTable.Location = new Point(10, 10);
            lblTable.Name = "lblTable";
            lblTable.Size = new Size(153, 17);
            lblTable.TabIndex = 0;
            lblTable.Text = "Выберите таблицу:";
            // 
            // comboBoxTables
            // 
            comboBoxTables.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            comboBoxTables.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxTables.Font = new Font("Microsoft Sans Serif", 10F);
            comboBoxTables.FormattingEnabled = true;
            comboBoxTables.Location = new Point(10, 40);
            comboBoxTables.Name = "comboBoxTables";
            comboBoxTables.Size = new Size(383, 24);
            comboBoxTables.TabIndex = 1;
            comboBoxTables.SelectedIndexChanged += ComboBoxTables_SelectedIndexChanged;
            // 
            // panelButtons
            // 
            panelButtons.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            panelButtons.BackColor = Color.LightSkyBlue;
            panelButtons.Controls.Add(btnCreateReport);
            panelButtons.Controls.Add(btnFilter);
            panelButtons.Controls.Add(btnSort);
            panelButtons.Controls.Add(btnSearch);
            panelButtons.Controls.Add(btnSave);
            panelButtons.Controls.Add(btnChange);
            panelButtons.Location = new Point(5, 90);
            panelButtons.Name = "panelButtons";
            panelButtons.Size = new Size(403, 527);
            panelButtons.TabIndex = 1;
            // 
            // btnCreateReport
            // 
            btnCreateReport.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnCreateReport.BackColor = Color.FromArgb(52, 152, 219);
            btnCreateReport.FlatStyle = FlatStyle.Flat;
            btnCreateReport.Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold);
            btnCreateReport.ForeColor = Color.White;
            btnCreateReport.Location = new Point(10, 265);
            btnCreateReport.Name = "btnCreateReport";
            btnCreateReport.Size = new Size(383, 50);
            btnCreateReport.TabIndex = 5;
            btnCreateReport.Text = "Создать отчёт";
            btnCreateReport.UseVisualStyleBackColor = false;
            btnCreateReport.Click += btnCreateReport_Click;
            // 
            // btnFilter
            // 
            btnFilter.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnFilter.BackColor = Color.FromArgb(52, 152, 219);
            btnFilter.FlatStyle = FlatStyle.Flat;
            btnFilter.Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold);
            btnFilter.ForeColor = Color.White;
            btnFilter.Location = new Point(10, 209);
            btnFilter.Name = "btnFilter";
            btnFilter.Size = new Size(383, 50);
            btnFilter.TabIndex = 4;
            btnFilter.Text = "Фильтр";
            btnFilter.UseVisualStyleBackColor = false;
            btnFilter.Click += btnFilter_Click;
            // 
            // btnSort
            // 
            btnSort.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnSort.BackColor = Color.FromArgb(52, 152, 219);
            btnSort.FlatStyle = FlatStyle.Flat;
            btnSort.Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold);
            btnSort.ForeColor = Color.White;
            btnSort.Location = new Point(10, 153);
            btnSort.Name = "btnSort";
            btnSort.Size = new Size(383, 50);
            btnSort.TabIndex = 3;
            btnSort.Text = "Сортировка";
            btnSort.UseVisualStyleBackColor = false;
            btnSort.Click += btnSort_Click;
            // 
            // btnSearch
            // 
            btnSearch.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnSearch.BackColor = Color.FromArgb(52, 152, 219);
            btnSearch.FlatStyle = FlatStyle.Flat;
            btnSearch.Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold);
            btnSearch.ForeColor = Color.White;
            btnSearch.Location = new Point(10, 97);
            btnSearch.Name = "btnSearch";
            btnSearch.Size = new Size(383, 50);
            btnSearch.TabIndex = 2;
            btnSearch.Text = "Поиск";
            btnSearch.UseVisualStyleBackColor = false;
            btnSearch.Click += btnSearch_Click;
            // 
            // btnSave
            // 
            btnSave.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnSave.BackColor = Color.FromArgb(52, 152, 219);
            btnSave.Enabled = false;
            btnSave.FlatStyle = FlatStyle.Flat;
            btnSave.Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold);
            btnSave.ForeColor = Color.White;
            btnSave.Location = new Point(10, 321);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(383, 50);
            btnSave.TabIndex = 1;
            btnSave.Text = "Сохранить";
            btnSave.UseVisualStyleBackColor = false;
            btnSave.Click += btnSave_Click;
            // 
            // btnChange
            // 
            btnChange.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnChange.BackColor = Color.FromArgb(52, 152, 219);
            btnChange.FlatStyle = FlatStyle.Flat;
            btnChange.Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold);
            btnChange.ForeColor = Color.White;
            btnChange.Location = new Point(10, 377);
            btnChange.Name = "btnChange";
            btnChange.Size = new Size(383, 50);
            btnChange.TabIndex = 0;
            btnChange.Text = "Изменить";
            btnChange.UseVisualStyleBackColor = false;
            btnChange.Click += btnChange_Click;
            // 
            // btnExitToMain
            // 
            btnExitToMain.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            btnExitToMain.BackColor = Color.FromArgb(220, 53, 69);
            btnExitToMain.FlatStyle = FlatStyle.Flat;
            btnExitToMain.Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold);
            btnExitToMain.ForeColor = Color.White;
            btnExitToMain.Location = new Point(5, 622);
            btnExitToMain.Name = "btnExitToMain";
            btnExitToMain.Size = new Size(403, 45);
            btnExitToMain.TabIndex = 2;
            btnExitToMain.Text = "Выйти в меню";
            btnExitToMain.UseVisualStyleBackColor = false;
            btnExitToMain.Click += btnExitToMain_Click;
            // 
            // ProjectForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1283, 692);
            Controls.Add(splitContainerMain);
            MinimumSize = new Size(900, 500);
            Name = "ProjectForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Работа с таблицами Access";
            Load += ProjectForm_Load;
            splitContainerMain.Panel1.ResumeLayout(false);
            splitContainerMain.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainerMain).EndInit();
            splitContainerMain.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridViewMain).EndInit();
            panelRight.ResumeLayout(false);
            panelTopRight.ResumeLayout(false);
            panelTopRight.PerformLayout();
            panelButtons.ResumeLayout(false);
            ResumeLayout(false);
        }

        #region Компоненты
        private System.Windows.Forms.SplitContainer splitContainerMain;
        private System.Windows.Forms.DataGridView dataGridViewMain;
        private System.Windows.Forms.Panel panelRight;
        private System.Windows.Forms.Panel panelTopRight;
        private System.Windows.Forms.Label lblTable;
        private System.Windows.Forms.ComboBox comboBoxTables;
        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Button btnChange;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.Button btnSort;
        private System.Windows.Forms.Button btnFilter;
        private System.Windows.Forms.Button btnCreateReport;
        private System.Windows.Forms.Button btnExitToMain;
        #endregion
    }
}