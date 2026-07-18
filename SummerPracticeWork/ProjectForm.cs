using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SummerPractice
{
    public partial class ProjectForm : Form
    {
        private ProjectDataManager dataManager;
        private bool isAdminMode = false;
        private List<string> primaryKeyColumns = new List<string>();
        private List<LookupInfo> currentLookups = new List<LookupInfo>();
        private DataTable currentDataTable;
        private string currentTableName;
        private bool _isShowingValidationError = false;
        private DataView currentDataView;

        public ProjectForm()
        {
            InitializeComponent();
            InitializeDataManager();
            SetupDataGridViewEvents();
        }

        #region Инициализация

        private void InitializeDataManager()
        {
            string dbName = "CourseWork.accdb";
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string dbPath = Path.Combine(basePath, dbName);

            if (!File.Exists(dbPath))
            {
                MessageBox.Show(
                    $"Файл базы данных не найден по пути:\n{dbPath}",
                    "Ошибка БД", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
                return;
            }

            string connStr = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";
            dataManager = new ProjectDataManager(connStr);
        }

        private void SetupDataGridViewEvents()
        {
            dataGridViewMain.DefaultValuesNeeded += DataGridViewMain_DefaultValuesNeeded;
            dataGridViewMain.CellValidating += DataGridViewMain_CellValidating;
            dataGridViewMain.UserDeletingRow += DataGridViewMain_UserDeletingRow;
            dataGridViewMain.KeyDown += DataGridViewMain_KeyDown;
            dataGridViewMain.DataError += DataGridViewMain_DataError;
            dataGridViewMain.CellBeginEdit += DataGridViewMain_CellBeginEdit;
            dataGridViewMain.CellEndEdit += DataGridViewMain_CellEndEdit;
        }

        #endregion

        #region Загрузка формы

        private void ProjectForm_Load(object sender, EventArgs e)
        {
            try
            {
                LoadTablesList();

                if (CurrentUser.IsLoggedIn)
                    this.Text = $"Система для работы с таблицами Access — {CurrentUser.GetDisplayName()}";

                UpdateUIForUserRole();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось подключиться к базе данных.\n\nОшибка: {ex.Message}",
                    "Критическая ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        private void LoadTablesList()
        {
            var tableNames = dataManager.GetTableNames(IsAdmin());

            if (tableNames.Count == 0)
            {
                MessageBox.Show("В базе данных нет доступных вам таблиц.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            comboBoxTables.DataSource = tableNames;
            if (tableNames.Count > 0)
                comboBoxTables.SelectedIndex = 0;
        }

        private void UpdateUIForUserRole()
        {
            bool isAdmin = IsAdmin();

            btnChange.Visible = isAdmin;
            btnSave.Visible = isAdmin;

            if (!isAdmin)
            {
                btnChange.Text = "Редактирование недоступно";
                btnChange.Enabled = false;
                btnSave.Enabled = false;
            }
            else
            {
                btnChange.Text = "Изменить";
                btnChange.Enabled = true;
                btnSave.Enabled = false;
            }
        }

        #endregion

        #region Свойства и вспомогательные методы

        private bool IsAdmin() => CurrentUser.IsLoggedIn && CurrentUser.IsAdmin;

        private bool HasUnsavedChanges() => currentDataTable?.GetChanges() != null;

        private void SafeEndEdit()
        {
            try
            {
                if (dataGridViewMain != null && dataGridViewMain.IsCurrentCellInEditMode)
                    dataGridViewMain.EndEdit();
            }
            catch { }
        }

        private DataRow GetCurrentDataRow(DataGridViewRow dgvRow)
        {
            if (dgvRow.DataBoundItem is DataRowView drv)
                return drv.Row;
            return null;
        }

        private bool IsNumericType(Type type)
        {
            return type == typeof(int) || type == typeof(long) ||
                   type == typeof(short) || type == typeof(byte) ||
                   type == typeof(decimal) || type == typeof(double) ||
                   type == typeof(float);
        }

        #endregion

        #region Работа с таблицами

        private void ComboBoxTables_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetterOrDigit(e.KeyChar) || char.IsWhiteSpace(e.KeyChar))
            {
                string searchChar = e.KeyChar.ToString().ToLower();

                int startIndex = comboBoxTables.SelectedIndex + 1;
                if (startIndex >= comboBoxTables.Items.Count)
                    startIndex = 0;

                for (int i = startIndex; i < comboBoxTables.Items.Count; i++)
                {
                    if (comboBoxTables.Items[i].ToString().ToLower().StartsWith(searchChar))
                    {
                        comboBoxTables.SelectedIndex = i;
                        e.Handled = true;
                        return;
                    }
                }

                for (int i = 0; i < startIndex; i++)
                {
                    if (comboBoxTables.Items[i].ToString().ToLower().StartsWith(searchChar))
                    {
                        comboBoxTables.SelectedIndex = i;
                        e.Handled = true;
                        return;
                    }
                }
            }
        }

        private void ComboBoxTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxTables.SelectedItem == null) return;

            string newTableName = comboBoxTables.SelectedItem.ToString();
            if (newTableName == currentTableName) return;

            SafeEndEdit();

            if (HasUnsavedChanges())
            {
                var result = MessageBox.Show(
                    "Есть несохранённые изменения. Сохранить их?",
                    "Несохранённые изменения",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    if (!TrySaveChanges()) return;
                }
                else if (result == DialogResult.Cancel)
                {
                    comboBoxTables.SelectedIndexChanged -= ComboBoxTables_SelectedIndexChanged;
                    comboBoxTables.SelectedItem = currentTableName;
                    comboBoxTables.SelectedIndexChanged += ComboBoxTables_SelectedIndexChanged;
                    return;
                }
                else
                {
                    currentDataTable?.RejectChanges();
                }
            }

            LoadTable(newTableName);
        }

        private void LoadTable(string tableName)
        {
            try
            {
                if (isAdminMode)
                    ExitEditModeSilently();

                dataGridViewMain.DataSource = null;
                dataGridViewMain.Columns.Clear();
                dataGridViewMain.Rows.Clear();

                dataManager.SelectTable(tableName);
                currentDataTable = dataManager.OriginalTable;
                currentTableName = tableName;
                currentDataView = currentDataTable.DefaultView;

                primaryKeyColumns = dataManager.GetPrimaryKeyColumns();
                currentLookups = dataManager.GetLookupsForCurrentTable();

                SetupColumns();
                dataGridViewMain.DataSource = currentDataView;
                ApplyColumnStyles();
                SetReadOnlyMode(!isAdminMode);
                UpdateButtonsState();
                ResetAllFilters();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке таблицы '{tableName}':\n{ex.Message}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetupColumns()
        {
            dataGridViewMain.AutoGenerateColumns = false;
            dataGridViewMain.Columns.Clear();

            foreach (DataColumn dc in currentDataTable.Columns)
            {
                DataGridViewColumn col = CreateColumn(dc);
                dataGridViewMain.Columns.Add(col);
            }
        }

        private DataGridViewColumn CreateColumn(DataColumn dc)
        {
            var lookup = currentLookups.FirstOrDefault(l =>
                l.FkColumn.Equals(dc.ColumnName, StringComparison.OrdinalIgnoreCase));

            DataGridViewColumn col;

            if (lookup != null && lookup.LookupTable != null)
            {
                var comboCol = new DataGridViewComboBoxColumn
                {
                    DataSource = lookup.LookupTable,
                    ValueMember = "ValueMember",
                    DisplayMember = "DisplayMember",
                    DisplayStyleForCurrentCellOnly = true,
                    FlatStyle = FlatStyle.Flat,
                    ValueType = dc.DataType
                };
                comboCol.DefaultCellStyle.NullValue = DBNull.Value;
                col = comboCol;
            }
            else if (dc.DataType == typeof(DateTime))
            {
                col = new DataGridViewTextBoxColumn();
                col.DefaultCellStyle.Format = "dd.MM.yyyy";
            }
            else
            {
                col = new DataGridViewTextBoxColumn();
            }

            col.Name = dc.ColumnName;
            col.HeaderText = dc.ColumnName;
            col.DataPropertyName = dc.ColumnName;
            col.ReadOnly = primaryKeyColumns.Contains(dc.ColumnName);

            return col;
        }

        private void ApplyColumnStyles()
        {
            foreach (string pkCol in primaryKeyColumns)
            {
                if (dataGridViewMain.Columns.Contains(pkCol))
                {
                    dataGridViewMain.Columns[pkCol].ReadOnly = true;
                    dataGridViewMain.Columns[pkCol].DefaultCellStyle.BackColor = Color.LightGray;
                    dataGridViewMain.Columns[pkCol].DefaultCellStyle.ForeColor = Color.Gray;
                }
            }

            foreach (DataGridViewColumn col in dataGridViewMain.Columns)
            {
                if (col is DataGridViewComboBoxColumn) continue;
                if (col.ValueType == typeof(DateTime))
                    col.DefaultCellStyle.Format = "dd.MM.yyyy";
            }
        }

        private void SetReadOnlyMode(bool readOnly)
        {
            dataGridViewMain.ReadOnly = readOnly;
            dataGridViewMain.AllowUserToAddRows = !readOnly && isAdminMode;
            dataGridViewMain.AllowUserToDeleteRows = !readOnly && isAdminMode;

            foreach (DataGridViewColumn col in dataGridViewMain.Columns)
            {
                if (col is DataGridViewComboBoxColumn)
                {
                    bool isPk = primaryKeyColumns?.Contains(col.Name) ?? false;
                    col.ReadOnly = isPk || readOnly;
                }
            }
        }

        private void UpdateButtonsState()
        {
            btnSave.Enabled = isAdminMode && HasUnsavedChanges();

            if (IsAdmin())
            {
                btnChange.Enabled = true;
                btnChange.Text = isAdminMode ? "Завершить редактирование" : "Изменить";
                btnChange.BackColor = isAdminMode ? Color.OrangeRed : SystemColors.Control;
                btnChange.ForeColor = isAdminMode ? Color.White : SystemColors.ControlText;
            }
        }

        private void ResetAllFilters()
        {
            if (currentDataView != null)
            {
                currentDataView.RowFilter = "";
                currentDataView.Sort = "";
            }

            foreach (DataGridViewColumn col in dataGridViewMain.Columns)
                col.Visible = true;

            foreach (DataGridViewRow row in dataGridViewMain.Rows)
                row.Visible = true;
        }

        #endregion

        #region Режим редактирования

        private void btnChange_Click(object sender, EventArgs e)
        {
            if (!IsAdmin())
            {
                MessageBox.Show(
                    "Эта функция доступна только администратору.",
                    "Доступ запрещён", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (isAdminMode)
            {
                ExitEditMode();
            }
            else
            {
                EnterEditMode();
            }
        }

        private void EnterEditMode()
        {
            if (currentDataTable == null)
            {
                MessageBox.Show("Данные таблицы не загружены.", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                SafeEndEdit();
                ResetAllFilters();

                isAdminMode = true;
                SetReadOnlyMode(false);
                UpdateButtonsState();

                MessageBox.Show(
                    "Режим редактирования включён.\n\n" +
                    "• Для изменения значения — кликните по ячейке.\n" +
                    "• Для FK-столбцов — выбор из выпадающего списка.\n" +
                    "• Для добавления строки — перейдите в последнюю пустую строку.\n" +
                    "• Для удаления строки — выделите её и нажмите Delete.\n" +
                    "• После изменений нажмите «Сохранить».",
                    "Режим редактирования", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при включении режима редактирования:\n{ex.Message}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExitEditMode()
        {
            SafeEndEdit();

            if (HasUnsavedChanges())
            {
                var result = MessageBox.Show(
                    "Есть несохранённые изменения. Сохранить их?",
                    "Несохранённые изменения",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    if (!TrySaveChanges()) return;
                }
                else if (result == DialogResult.Cancel)
                {
                    return;
                }
                else
                {
                    currentDataTable?.RejectChanges();
                }
            }

            ExitEditModeSilently();
        }

        private void ExitEditModeSilently()
        {
            SafeEndEdit();

            isAdminMode = false;
            SetReadOnlyMode(true);
            UpdateButtonsState();

            if (currentDataTable != null)
            {
                dataGridViewMain.DataSource = currentDataTable.DefaultView;
            }
        }

        #endregion

        #region Сохранение

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!isAdminMode)
            {
                MessageBox.Show("Сначала включите режим редактирования.",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            TrySaveChanges();
        }

        private bool TrySaveChanges()
        {
            SafeEndEdit();

            if (!ValidateRelationships())
                return false;

            var result = dataManager.TrySaveChanges();

            if (result.Success)
            {
                MessageBox.Show("Изменения успешно сохранены.",
                    "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                UpdateButtonsState();
                return true;
            }
            else
            {
                MessageBox.Show($"Ошибка сохранения:\n{result.ErrorMessage}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private bool ValidateRelationships()
        {
            if (currentDataTable == null) return true;

            DataTable changes = currentDataTable.GetChanges();
            if (changes == null) return true;

            foreach (DataRow row in changes.Rows)
            {
                if (row.RowState == DataRowState.Deleted)
                {
                    var violations = dataManager.CheckForeignKeyViolations(row);
                    if (violations.Count > 0)
                    {
                        MessageBox.Show(
                            "Нельзя удалить запись — существуют связанные данные:\n\n" +
                            string.Join("\n", violations),
                            "Нарушение связей", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }

            return true;
        }

        #endregion

        #region Поиск

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (currentDataTable == null)
            {
                MessageBox.Show("Сначала выберите таблицу.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SafeEndEdit();

            using var searchDialog = CreateSearchDialog();

            if (searchDialog.ShowDialog() == DialogResult.OK)
            {
                var txtSearch = searchDialog.Controls.Find("txtSearch", true).FirstOrDefault() as TextBox;
                var cmbColumn = searchDialog.Controls.Find("cmbColumn", true).FirstOrDefault() as ComboBox;

                string searchText = txtSearch?.Text.Trim() ?? string.Empty;
                string columnName = cmbColumn?.SelectedItem?.ToString() ?? "Все столбцы";

                if (string.IsNullOrWhiteSpace(searchText))
                {
                    MessageBox.Show("Введите текст для поиска.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                PerformSearch(searchText, columnName);
            }
        }

        private Form CreateSearchDialog()
        {
            var dialog = new Form
            {
                Text = "Поиск по таблице",
                Size = new Size(400, 230),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var lblSearch = new Label
            {
                Text = "Введите текст или число:",
                Location = new Point(15, 15),
                AutoSize = true
            };

            var txtSearch = new TextBox
            {
                Name = "txtSearch",
                Location = new Point(15, 40),
                Width = 350
            };

            var lblCol = new Label
            {
                Text = "Искать в:",
                Location = new Point(15, 75),
                AutoSize = true
            };

            var cmbColumn = new ComboBox
            {
                Name = "cmbColumn",
                Location = new Point(15, 95),
                Width = 350,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            cmbColumn.Items.Add("Все столбцы");
            foreach (DataGridViewColumn col in dataGridViewMain.Columns)
            {
                if (col.Visible && currentDataTable.Columns.Contains(col.Name))
                    cmbColumn.Items.Add(col.Name);
            }
            cmbColumn.SelectedIndex = 0;

            var btnOk = new Button
            {
                Text = "Найти",
                DialogResult = DialogResult.OK,
                Location = new Point(60, 140),
                Width = 90,
                Height = 30
            };

            var btnReset = new Button
            {
                Text = "Сбросить",
                Location = new Point(160, 140),
                Width = 110,
                Height = 30
            };

            var btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new Point(280, 140),
                Width = 90,
                Height = 30
            };

            dialog.Controls.AddRange(new Control[] {
                lblSearch, txtSearch, lblCol, cmbColumn,
                btnOk, btnReset, btnCancel
            });

            dialog.AcceptButton = btnOk;
            dialog.CancelButton = btnCancel;

            btnReset.Click += (s, ev) =>
            {
                ResetSearch();
                dialog.Close();
            };

            return dialog;
        }

        private void PerformSearch(string searchText, string columnName)
        {
            if (currentDataView == null) return;

            var conditions = new List<string>();
            string escapedText = searchText.Replace("'", "''");

            bool isNumber = decimal.TryParse(searchText,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out decimal numValue);

            bool isDate = DateTime.TryParse(searchText, out DateTime dateValue);

            IEnumerable<DataColumn> colsToSearch = columnName == "Все столбцы"
                ? currentDataTable.Columns.Cast<DataColumn>()
                : currentDataTable.Columns.Cast<DataColumn>()
                    .Where(c => c.ColumnName.Equals(columnName, StringComparison.OrdinalIgnoreCase));

            foreach (DataColumn col in colsToSearch)
            {
                if (col.DataType == typeof(string))
                {
                    conditions.Add($"[{col.ColumnName}] LIKE '*{escapedText}*'");
                }
                else if (IsNumericType(col.DataType) && isNumber)
                {
                    string numStr = numValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    conditions.Add($"[{col.ColumnName}] = {numStr}");
                }
                else if (col.DataType == typeof(DateTime) && isDate)
                {
                    conditions.Add($"[{col.ColumnName}] = #{dateValue:yyyy-MM-dd}#");
                }
            }

            if (conditions.Count == 0)
            {
                MessageBox.Show("Поиск не дал результатов.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            currentDataView.RowFilter = string.Join(" OR ", conditions);
            int count = currentDataView.Count;

            MessageBox.Show($"Найдено записей: {count}", "Результаты поиска",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ResetSearch()
        {
            try
            {
                if (currentDataView != null)
                    currentDataView.RowFilter = "";

                MessageBox.Show("Поиск сброшен.", "Готово",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сброса поиска: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region Сортировка

        private void btnSort_Click(object sender, EventArgs e)
        {
            if (currentDataTable == null)
            {
                MessageBox.Show("Сначала выберите таблицу.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SafeEndEdit();

            using var sortDialog = CreateSortDialog();

            if (sortDialog.ShowDialog() == DialogResult.OK)
            {
                var cmbColumn = sortDialog.Controls.Find("cmbColumn", true).FirstOrDefault() as ComboBox;
                var rbAscending = sortDialog.Controls.Find("rbAscending", true).FirstOrDefault() as RadioButton;

                if (cmbColumn?.SelectedItem == null)
                {
                    MessageBox.Show("Выберите столбец для сортировки.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string columnName = cmbColumn.SelectedItem.ToString();
                bool ascending = rbAscending?.Checked ?? true;

                try
                {
                    string sortDirection = ascending ? "ASC" : "DESC";
                    currentDataView.Sort = $"[{columnName}] {sortDirection}";

                    string dir = ascending ? "по возрастанию" : "по убыванию";
                    MessageBox.Show($"Таблица отсортирована по столбцу '{columnName}' {dir}.",
                        "Сортировка применена", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при сортировке:\n{ex.Message}", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private Form CreateSortDialog()
        {
            var dialog = new Form
            {
                Text = "Сортировка таблицы",
                Size = new Size(460, 300),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var grpColumn = new GroupBox
            {
                Text = "Выберите столбец для сортировки:",
                Location = new Point(15, 10),
                Size = new Size(415, 80)
            };

            var cmbColumn = new ComboBox
            {
                Name = "cmbColumn",
                Location = new Point(15, 30),
                Width = 385,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            foreach (DataGridViewColumn dgvCol in dataGridViewMain.Columns)
            {
                if (dgvCol.Visible && currentDataTable.Columns.Contains(dgvCol.Name))
                    cmbColumn.Items.Add(dgvCol.Name);
            }
            if (cmbColumn.Items.Count > 0) cmbColumn.SelectedIndex = 0;

            grpColumn.Controls.Add(cmbColumn);

            var grpSortType = new GroupBox
            {
                Text = "Тип сортировки:",
                Location = new Point(15, 100),
                Size = new Size(415, 90)
            };

            var rbAscending = new RadioButton
            {
                Name = "rbAscending",
                Text = "По возрастанию (А→Я, 0→9)",
                Location = new Point(15, 30),
                AutoSize = true,
                Checked = true
            };

            var rbDescending = new RadioButton
            {
                Text = "По убыванию (Я→А, 9→0)",
                Location = new Point(15, 60),
                AutoSize = true
            };

            grpSortType.Controls.AddRange(new Control[] { rbAscending, rbDescending });

            var btnApply = new Button
            {
                Text = "Применить",
                DialogResult = DialogResult.OK,
                Location = new Point(70, 210),
                Size = new Size(110, 35)
            };

            var btnReset = new Button
            {
                Text = "Сбросить сортировку",
                Location = new Point(190, 210),
                Size = new Size(130, 35)
            };

            var btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new Point(330, 210),
                Size = new Size(110, 35)
            };

            dialog.Controls.AddRange(new Control[] { grpColumn, grpSortType, btnApply, btnReset, btnCancel });
            dialog.AcceptButton = btnApply;
            dialog.CancelButton = btnCancel;

            btnReset.Click += (s, ev) => { ResetSort(); dialog.Close(); };

            return dialog;
        }

        private void ResetSort()
        {
            try
            {
                if (currentDataView != null)
                    currentDataView.Sort = "";

                MessageBox.Show("Сортировка сброшена.", "Готово",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сброса сортировки: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region Фильтр

        private void btnFilter_Click(object sender, EventArgs e)
        {
            if (currentDataTable == null)
            {
                MessageBox.Show("Сначала выберите таблицу.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SafeEndEdit();

            using var filterDialog = CreateFilterDialog();

            if (filterDialog.ShowDialog() == DialogResult.OK)
            {
                var checkedListBox = filterDialog.Controls.Find("checkedListBoxColumns", true)[0] as CheckedListBox;
                var txtFrom = filterDialog.Controls.Find("txtFrom", true)[0] as TextBox;
                var txtTo = filterDialog.Controls.Find("txtTo", true)[0] as TextBox;

                if (checkedListBox.CheckedItems.Count == 0)
                {
                    MessageBox.Show("Выберите хотя бы один столбец для отображения.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!int.TryParse(txtFrom.Text, out int fromRow) ||
                    !int.TryParse(txtTo.Text, out int toRow))
                {
                    MessageBox.Show("Введите корректные номера строк (числа).", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                fromRow = Math.Max(1, fromRow);
                toRow = Math.Min(currentDataTable.Rows.Count, toRow);

                if (fromRow > toRow)
                {
                    MessageBox.Show("Начальная строка не может быть больше конечной.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                ApplyFilter(checkedListBox, fromRow, toRow);
            }
        }

        private Form CreateFilterDialog()
        {
            var dialog = new Form
            {
                Text = "Фильтр таблицы",
                Size = new Size(480, 520),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var grpColumns = new GroupBox
            {
                Text = "Выберите столбцы для отображения:",
                Location = new Point(15, 10),
                Size = new Size(435, 300)
            };

            var checkedListBox = new CheckedListBox
            {
                Name = "checkedListBoxColumns",
                Location = new Point(15, 25),
                Size = new Size(405, 220),
                CheckOnClick = true
            };

            foreach (DataColumn col in currentDataTable.Columns)
                checkedListBox.Items.Add(col.ColumnName, true);

            var btnSelectAll = new Button
            {
                Text = "Выбрать все",
                Location = new Point(15, 255),
                Size = new Size(130, 30)
            };
            btnSelectAll.Click += (s, ev) =>
            {
                for (int i = 0; i < checkedListBox.Items.Count; i++)
                    checkedListBox.SetItemChecked(i, true);
            };

            var btnDeselectAll = new Button
            {
                Text = "Снять все",
                Location = new Point(155, 255),
                Size = new Size(130, 30)
            };
            btnDeselectAll.Click += (s, ev) =>
            {
                for (int i = 0; i < checkedListBox.Items.Count; i++)
                    checkedListBox.SetItemChecked(i, false);
            };

            grpColumns.Controls.AddRange(new Control[] { checkedListBox, btnSelectAll, btnDeselectAll });

            var grpRows = new GroupBox
            {
                Text = "Диапазон строк:",
                Location = new Point(15, 320),
                Size = new Size(435, 80)
            };

            var lblFrom = new Label { Text = "С:", Location = new Point(15, 30), AutoSize = true };
            var txtFrom = new TextBox { Name = "txtFrom", Text = "1", Location = new Point(40, 27), Width = 70 };
            var lblTo = new Label { Text = "По:", Location = new Point(120, 30), AutoSize = true };
            var txtTo = new TextBox { Name = "txtTo", Text = currentDataTable.Rows.Count.ToString(), Location = new Point(160, 27), Width = 70 };
            var lblInfo = new Label
            {
                Text = $"Всего строк: {currentDataTable.Rows.Count}",
                Location = new Point(260, 30),
                AutoSize = true,
                ForeColor = Color.Gray
            };

            grpRows.Controls.AddRange(new Control[] { lblFrom, txtFrom, lblTo, txtTo, lblInfo });

            var btnApply = new Button
            {
                Text = "Применить",
                DialogResult = DialogResult.OK,
                Location = new Point(70, 420),
                Size = new Size(100, 35)
            };
            var btnReset = new Button
            {
                Text = "Сбросить фильтр",
                Location = new Point(190, 420),
                Size = new Size(120, 35)
            };
            btnReset.Click += (s, ev) => { ResetFilter(); dialog.Close(); };
            var btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new Point(320, 420),
                Size = new Size(100, 35)
            };

            dialog.Controls.AddRange(new Control[] { grpColumns, grpRows, btnApply, btnReset, btnCancel });
            dialog.AcceptButton = btnApply;
            dialog.CancelButton = btnCancel;
            return dialog;
        }

        private void ApplyFilter(CheckedListBox checkedListBox, int fromRow, int toRow)
        {
            try
            {
                foreach (DataGridViewColumn col in dataGridViewMain.Columns)
                    col.Visible = true;
                foreach (DataGridViewRow row in dataGridViewMain.Rows)
                    row.Visible = true;

                var selectedColumns = new HashSet<string>();
                foreach (var item in checkedListBox.CheckedItems)
                    selectedColumns.Add(item.ToString());

                foreach (DataGridViewColumn col in dataGridViewMain.Columns)
                    col.Visible = selectedColumns.Contains(col.Name);

                int visibleCount = 0;
                for (int i = 0; i < dataGridViewMain.Rows.Count; i++)
                {
                    var row = dataGridViewMain.Rows[i];
                    if (row.IsNewRow) continue;
                    bool isVisible = (i + 1 >= fromRow && i + 1 <= toRow);
                    row.Visible = isVisible;
                    if (isVisible) visibleCount++;
                }

                MessageBox.Show(
                    $"Применён фильтр:\n• Столбцов: {selectedColumns.Count}\n• Строк: {visibleCount} (с {fromRow} по {toRow})",
                    "Фильтр применён", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при применении фильтра:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ResetFilter()
        {
            try
            {
                foreach (DataGridViewColumn col in dataGridViewMain.Columns)
                    col.Visible = true;
                foreach (DataGridViewRow row in dataGridViewMain.Rows)
                    row.Visible = true;

                MessageBox.Show("Фильтр сброшен.", "Готово",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сброса фильтра: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region Отчёт в Word

        private void btnCreateReport_Click(object sender, EventArgs e)
        {
            if (dataGridViewMain.DataSource == null)
            {
                MessageBox.Show("Нет данных для создания отчёта.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SafeEndEdit();

            string userName = CurrentUser.IsLoggedIn ? CurrentUser.Login : "Неизвестно";
            string tableName = comboBoxTables.SelectedItem?.ToString() ?? "Таблица";

            var result = WordReportGenerator.CreateReport(dataGridViewMain, tableName, userName);

            if (result.Success)
                MessageBox.Show($"Отчёт успешно создан!\n\nФайл сохранён:\n{result.FilePath}",
                    "Отчёт готов", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(result.ErrorMessage, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        #endregion

        #region Валидация данных

        private void DataGridViewMain_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            if (dataManager == null || primaryKeyColumns == null) return;

            try
            {
                foreach (string pkCol in primaryKeyColumns)
                {
                    if (string.IsNullOrEmpty(pkCol) || !dataGridViewMain.Columns.Contains(pkCol))
                        continue;

                    object nextValue = dataManager.GetNextPrimaryKeyValue(pkCol);
                    if (e.Row.Cells[pkCol] != null)
                        e.Row.Cells[pkCol].Value = nextValue;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DefaultValuesNeeded error: {ex.Message}");
            }
        }

        private void DataGridViewMain_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (!isAdminMode || e.RowIndex < 0 || e.RowIndex >= dataGridViewMain.Rows.Count) return;
            if (dataGridViewMain.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn) return;
            if (_isShowingValidationError) return;

            try
            {
                var row = dataGridViewMain.Rows[e.RowIndex];
                if (!(row.DataBoundItem is DataRowView drv)) return;

                object newValue = e.FormattedValue;
                string colName = dataGridViewMain.Columns[e.ColumnIndex].Name;

                if (newValue == null || string.IsNullOrWhiteSpace(newValue.ToString()))
                {
                    bool hasOtherValues = false;
                    foreach (DataColumn col in drv.Row.Table.Columns)
                    {
                        if (col.ColumnName == colName) continue;
                        object v = drv.Row[col.ColumnName];
                        if (v != null && v != DBNull.Value && !string.IsNullOrWhiteSpace(v.ToString()))
                        {
                            hasOtherValues = true;
                            break;
                        }
                    }

                    object origVal = drv.Row[colName];
                    bool wasFilled = origVal != null && origVal != DBNull.Value &&
                                     !string.IsNullOrWhiteSpace(origVal.ToString());

                    if (hasOtherValues || wasFilled)
                    {
                        e.Cancel = true;
                        string header = dataGridViewMain.Columns[e.ColumnIndex].HeaderText;
                        row.Cells[e.ColumnIndex].ErrorText = $"Поле «{header}» не может быть пустым.";

                        _isShowingValidationError = true;
                        MessageBox.Show(
                            $"Нельзя оставить поле «{header}» пустым.\n\n" +
                            "Строка содержит данные — либо заполните поле, либо удалите строку целиком.",
                            "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        _isShowingValidationError = false;
                    }
                }
                else
                {
                    row.Cells[e.ColumnIndex].ErrorText = "";
                }
            }
            catch (Exception ex)
            {
                _isShowingValidationError = false;
                System.Diagnostics.Debug.WriteLine($"CellValidating error: {ex.Message}");
            }
        }

        private void DataGridViewMain_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            try
            {
                if (!isAdminMode) return;

                DataRow row = GetCurrentDataRow(e.Row);
                if (row == null || row.RowState == DataRowState.Deleted) return;

                var violations = dataManager.CheckForeignKeyViolations(row);
                if (violations.Count > 0)
                {
                    e.Cancel = true;
                    MessageBox.Show(
                        "Нельзя удалить строку — существуют связанные данные:\n\n" +
                        string.Join("\n", violations) +
                        "\n\nСначала удалите связанные записи из дочерних таблиц.",
                        "Нарушение связей", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var result = MessageBox.Show(
                    "Удалить выбранную строку?\nЭто действие нельзя отменить после сохранения.",
                    "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.No) e.Cancel = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении строки:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
            }
        }

        private void DataGridViewMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (!isAdminMode || e.KeyCode != Keys.Delete || dataGridViewMain.SelectedRows.Count == 0)
                return;

            var rowsToDelete = new List<DataRow>();

            foreach (DataGridViewRow dgvRow in dataGridViewMain.SelectedRows)
            {
                if (dgvRow.IsNewRow) continue;
                DataRow row = GetCurrentDataRow(dgvRow);
                if (row == null) continue;

                var violations = dataManager.CheckForeignKeyViolations(row);
                if (violations.Count > 0)
                {
                    MessageBox.Show(
                        "Нельзя удалить выбранные строки — есть связанные данные:\n\n" +
                        string.Join("\n", violations),
                        "Нарушение связей", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                rowsToDelete.Add(row);
            }

            if (rowsToDelete.Count == 0) return;

            var result = MessageBox.Show(
                $"Удалить выбранные строки ({rowsToDelete.Count})?",
                "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                foreach (DataRow dataRow in rowsToDelete)
                {
                    if (dataRow.RowState != DataRowState.Deleted)
                        dataRow.Delete();
                }
                UpdateButtonsState();
            }
        }

        private void DataGridViewMain_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;
            System.Diagnostics.Debug.WriteLine($"DataError: {e.Exception?.Message}");
        }

        private void DataGridViewMain_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (!isAdminMode)
                e.Cancel = true;
        }

        private void DataGridViewMain_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            UpdateButtonsState();
        }

        #endregion

        #region Выход

        private void btnExitToMain_Click(object sender, EventArgs e)
        {
            if (HasUnsavedChanges())
            {
                var result = MessageBox.Show(
                    "Есть несохранённые изменения. Выйти без сохранения?",
                    "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.No)
                    return;
            }

            CurrentUser.Logout();
            this.Close();
        }

        #endregion
    }
}