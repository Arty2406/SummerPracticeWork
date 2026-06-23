using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SummerPractice
{
    public partial class ProjectForm : Form
    {
        private ProjectDataManager dataManager;
        private bool isAdminMode = false;

        public ProjectForm()
        {
            InitializeComponent();

            string dbName = "CourseWork.accdb";
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string dbPath = Path.Combine(basePath, dbName);

            if (!File.Exists(dbPath))
            {
                MessageBox.Show(
                    $"Файл базы данных не найден по пути:\n{dbPath}\n\nУбедитесь, что файл скопирован в папку bin\\Debug\\netX.X",
                    "Ошибка БД", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
                return;
            }

            string connStr = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";
            dataManager = new ProjectDataManager(connStr);

            this.Load += ProjectForm_Load;
        }

        private void ProjectForm_Load(object sender, EventArgs e)
        {
            try
            {
                var tableNames = dataManager.GetTableNames(IsAdmin());

                if (tableNames.Count == 0)
                {
                    MessageBox.Show("В базе данных нет доступных вам таблиц.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                comboBoxTables.DataSource = tableNames;
                comboBoxTables.SelectedIndex = 0;

                if (CurrentUser.IsLoggedIn)
                    this.Text = $"Система для работы с таблицами Access — {CurrentUser.GetDisplayName()}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось подключиться к базе данных.\n\nПроверьте:\n1) Файл БД лежит рядом с программой (.exe).\n2) Установлен ли Access Database Engine.\n\nОшибка: {ex.Message}",
                    "Критическая ошибка подключения", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        private bool IsAdmin()
        {
            return CurrentUser.IsLoggedIn && CurrentUser.IsAdmin;
        }

        private void ComboBoxTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxTables.SelectedItem == null) return;

            // Проверка несохранённых изменений
            if (dataManager.HasUnsavedChanges)
            {
                var result = MessageBox.Show(
                    "Есть несохранённые изменения в текущей таблице. Сохранить их?",
                    "Несохранённые изменения", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    if (!TrySaveChanges()) return;
                }
                else if (result == DialogResult.Cancel)
                {
                    return;
                }
            }

            string tableName = comboBoxTables.SelectedItem.ToString();

            try
            {
                dataManager.SelectTable(tableName);
                dataGridViewMain.DataSource = dataManager.OriginalTable;

                // Форматирование столбцов с датами
                foreach (DataGridViewColumn col in dataGridViewMain.Columns)
                {
                    if (col.ValueType == typeof(DateTime))
                        col.DefaultCellStyle.Format = "dd.MM.yyyy";
                }

                dataGridViewMain.ReadOnly = !isAdminMode;
                dataGridViewMain.AllowUserToAddRows = isAdminMode;
                dataGridViewMain.AllowUserToDeleteRows = isAdminMode;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении таблицы '{tableName}':\n{ex.Message}",
                    "Ошибка данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExitToMain_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "Вернуться в главное меню?\n\nВы будете разлогинены и вернётесь к форме входа.",
                "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                CurrentUser.Logout();
                this.Close();
            }
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            if (!IsAdmin())
            {
                MessageBox.Show(
                    "Эта функция доступна только администратору.\n\nВойдите в систему под учётной записью администратора.",
                    "Доступ запрещён", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (isAdminMode)
            {
                ExitEditMode();
                return;
            }

            // Сброс всех состояний перед редактированием
            ResetAllViewStates();
            dataManager.ResetAll();

            dataManager.LoadForeignKeys();

            isAdminMode = true;
            dataGridViewMain.ReadOnly = false;
            dataGridViewMain.AllowUserToAddRows = true;
            dataGridViewMain.AllowUserToDeleteRows = true;

            btnChange.Text = "Завершить редактирование";
            btnChange.BackColor = System.Drawing.Color.OrangeRed;
            btnChange.ForeColor = System.Drawing.Color.White;

            MessageBox.Show(
                "Режим редактирования включён.\n\n" +
                "• Для изменения значения — кликните по ячейке и введите новое.\n" +
                "• Для добавления строки — перейдите в последнюю пустую строку.\n" +
                "• Для удаления строки — выделите её и нажмите Delete.\n" +
                "• После изменений нажмите «Сохранить».\n\n" +
                "Внимание: учитываются связи между таблицами!",
                "Режим редактирования", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ExitEditMode()
        {
            if (dataManager.HasUnsavedChanges)
            {
                var result = MessageBox.Show(
                    "Есть несохранённые изменения. Сохранить их перед выходом?",
                    "Несохранённые изменения", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    if (!TrySaveChanges()) return;
                }
                else if (result == DialogResult.Cancel)
                {
                    return;
                }
            }

            isAdminMode = false;
            dataGridViewMain.ReadOnly = true;
            dataGridViewMain.AllowUserToAddRows = false;
            dataGridViewMain.AllowUserToDeleteRows = false;

            btnChange.Text = "Изменить";
            btnChange.BackColor = System.Drawing.SystemColors.Control;
            btnChange.ForeColor = System.Drawing.SystemColors.ControlText;

            MessageBox.Show("Режим редактирования завершён.", "Готово",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #region Фильтр

        private void btnFilter_Click(object sender, EventArgs e)
        {
            if (dataManager.OriginalTable == null)
            {
                MessageBox.Show("Сначала выберите таблицу.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

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

                if (!int.TryParse(txtFrom.Text, out int fromRow) || !int.TryParse(txtTo.Text, out int toRow))
                {
                    MessageBox.Show("Введите корректные номера строк (числа).", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                fromRow = Math.Max(1, fromRow);
                toRow = Math.Min(dataManager.OriginalTable.Rows.Count, toRow);

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
                Size = new System.Drawing.Size(480, 520),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var grpColumns = new GroupBox
            {
                Text = "Выберите столбцы для отображения:",
                Location = new System.Drawing.Point(15, 10),
                Size = new System.Drawing.Size(435, 300)
            };

            var checkedListBox = new CheckedListBox
            {
                Name = "checkedListBoxColumns",
                Location = new System.Drawing.Point(15, 25),
                Size = new System.Drawing.Size(405, 220),
                CheckOnClick = true
            };

            foreach (DataColumn col in dataManager.OriginalTable.Columns)
                checkedListBox.Items.Add(col.ColumnName, true);

            var btnSelectAll = new Button { Text = "Выбрать все", Location = new System.Drawing.Point(15, 255), Size = new System.Drawing.Size(130, 30) };
            btnSelectAll.Click += (s, ev) => { for (int i = 0; i < checkedListBox.Items.Count; i++) checkedListBox.SetItemChecked(i, true); };

            var btnDeselectAll = new Button { Text = "Снять все", Location = new System.Drawing.Point(155, 255), Size = new System.Drawing.Size(130, 30) };
            btnDeselectAll.Click += (s, ev) => { for (int i = 0; i < checkedListBox.Items.Count; i++) checkedListBox.SetItemChecked(i, false); };

            grpColumns.Controls.AddRange(new Control[] { checkedListBox, btnSelectAll, btnDeselectAll });

            var grpRows = new GroupBox
            {
                Text = "Диапазон строк:",
                Location = new System.Drawing.Point(15, 320),
                Size = new System.Drawing.Size(435, 80)
            };

            var lblFrom = new Label { Text = "С:", Location = new System.Drawing.Point(15, 30), AutoSize = true };
            var txtFrom = new TextBox { Name = "txtFrom", Text = "1", Location = new System.Drawing.Point(40, 27), Width = 70 };
            var lblTo = new Label { Text = "По:", Location = new System.Drawing.Point(120, 30), AutoSize = true };
            var txtTo = new TextBox { Name = "txtTo", Text = dataManager.OriginalTable.Rows.Count.ToString(), Location = new System.Drawing.Point(160, 27), Width = 70 };
            var lblInfo = new Label { Text = $"Всего строк: {dataManager.OriginalTable.Rows.Count}", Location = new System.Drawing.Point(260, 30), AutoSize = true, ForeColor = System.Drawing.Color.Gray };

            grpRows.Controls.AddRange(new Control[] { lblFrom, txtFrom, lblTo, txtTo, lblInfo });

            var btnApply = new Button { Text = "Применить", DialogResult = DialogResult.OK, Location = new System.Drawing.Point(70, 420), Size = new System.Drawing.Size(100, 35) };
            var btnReset = new Button { Text = "Сбросить фильтр", Location = new System.Drawing.Point(190, 420), Size = new System.Drawing.Size(120, 35) };
            btnReset.Click += (s, ev) => { ResetFilter(); dialog.Close(); };
            var btnCancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, Location = new System.Drawing.Point(320, 420), Size = new System.Drawing.Size(100, 35) };

            dialog.Controls.AddRange(new Control[] { grpColumns, grpRows, btnApply, btnReset, btnCancel });
            dialog.AcceptButton = btnApply;
            dialog.CancelButton = btnCancel;

            return dialog;
        }

        private void ApplyFilter(CheckedListBox checkedListBox, int fromRow, int toRow)
        {
            try
            {
                dataGridViewMain.EndEdit();

                // Если DataSource — DataView, возвращаем к originalTable
                if (dataGridViewMain.DataSource is DataView)
                    dataGridViewMain.DataSource = dataManager.OriginalTable;

                // Показываем все строки и столбцы
                foreach (DataGridViewColumn col in dataGridViewMain.Columns)
                    try { col.Visible = true; } catch { }
                foreach (DataGridViewRow row in dataGridViewMain.Rows)
                    try { row.Visible = true; } catch { }

                var selectedColumns = new HashSet<string>();
                foreach (var item in checkedListBox.CheckedItems)
                    selectedColumns.Add(item.ToString());

                foreach (DataGridViewColumn col in dataGridViewMain.Columns)
                    col.Visible = selectedColumns.Contains(col.Name);

                int visibleCount = 0;
                for (int i = 0; i < dataGridViewMain.Rows.Count; i++)
                {
                    DataGridViewRow row = dataGridViewMain.Rows[i];
                    if (row.IsNewRow) continue;

                    int rowNumber = i + 1;
                    bool isVisible = (rowNumber >= fromRow && rowNumber <= toRow);
                    try { row.Visible = isVisible; } catch { }
                    if (isVisible) visibleCount++;
                }

                dataManager.MarkFiltered();

                MessageBox.Show(
                    $"Применён фильтр:\n• Столбцов отображается: {selectedColumns.Count}\n• Строк отображается: {visibleCount} (с {fromRow} по {toRow})",
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
                    try { col.Visible = true; } catch { }
                foreach (DataGridViewRow row in dataGridViewMain.Rows)
                    try { row.Visible = true; } catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сброса фильтра: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region Сортировка

        private void btnSort_Click(object sender, EventArgs e)
        {
            if (dataGridViewMain.DataSource == null)
            {
                MessageBox.Show("Сначала выберите таблицу.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (dataManager.OriginalTable == null)
            {
                MessageBox.Show("Нет данных для сортировки.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using var sortDialog = CreateSortDialog();

            if (sortDialog.ShowDialog() == DialogResult.OK)
            {
                var cmbColumn = sortDialog.Controls.Find("cmbColumn", true)[0] as ComboBox;
                var rbAscending = sortDialog.Controls.Find("rbAscending", true)[0] as RadioButton;

                if (cmbColumn.SelectedItem == null)
                {
                    MessageBox.Show("Выберите столбец для сортировки.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string columnName = cmbColumn.SelectedItem.ToString();
                bool ascending = rbAscending.Checked;

                try
                {
                    DataView dv = dataManager.ApplySort(columnName, ascending);
                    dataGridViewMain.DataSource = dv;

                    string directionText = ascending ? "по возрастанию" : "по убыванию";
                    MessageBox.Show($"Таблица отсортирована по столбцу '{columnName}' {directionText}.",
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
                Size = new System.Drawing.Size(460, 340),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var grpColumn = new GroupBox
            {
                Text = "Выберите столбец для сортировки:",
                Location = new System.Drawing.Point(15, 10),
                Size = new System.Drawing.Size(415, 110)
            };

            var cmbColumn = new ComboBox
            {
                Name = "cmbColumn",
                Location = new System.Drawing.Point(15, 30),
                Width = 385,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            var lblType = new Label
            {
                Text = "Тип данных: (не выбрано)",
                Location = new System.Drawing.Point(15, 70),
                AutoSize = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.DarkBlue
            };

            // Только видимые столбцы
            foreach (DataGridViewColumn dgvCol in dataGridViewMain.Columns)
                if (dgvCol.Visible) cmbColumn.Items.Add(dgvCol.Name);
            if (cmbColumn.Items.Count > 0) cmbColumn.SelectedIndex = 0;

            grpColumn.Controls.AddRange(new Control[] { cmbColumn, lblType });

            var grpSortType = new GroupBox
            {
                Text = "Тип сортировки:",
                Location = new System.Drawing.Point(15, 130),
                Size = new System.Drawing.Size(415, 100)
            };

            var rbAscending = new RadioButton { Name = "rbAscending", Text = "По возрастанию (А→Я, 0→9, старые→новые)", Location = new System.Drawing.Point(15, 30), AutoSize = true, Checked = true };
            var rbDescending = new RadioButton { Text = "По убыванию (Я→А, 9→0, новые→старые)", Location = new System.Drawing.Point(15, 60), AutoSize = true };

            grpSortType.Controls.AddRange(new Control[] { rbAscending, rbDescending });

            var btnApply = new Button { Text = "Применить", DialogResult = DialogResult.OK, Location = new System.Drawing.Point(70, 250), Size = new System.Drawing.Size(110, 35) };
            var btnReset = new Button { Text = "Сбросить сортировку", Location = new System.Drawing.Point(190, 250), Size = new System.Drawing.Size(130, 35) };
            btnReset.Click += (s, ev) => { ResetSort(); dialog.Close(); };
            var btnCancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, Location = new System.Drawing.Point(330, 250), Size = new System.Drawing.Size(110, 35) };

            dialog.Controls.AddRange(new Control[] { grpColumn, grpSortType, btnApply, btnReset, btnCancel });
            dialog.AcceptButton = btnApply;
            dialog.CancelButton = btnCancel;

            cmbColumn.SelectedIndexChanged += (s, ev) =>
            {
                if (cmbColumn.SelectedItem == null) return;
                string colName = cmbColumn.SelectedItem.ToString();
                if (dataManager.OriginalTable.Columns.Contains(colName))
                {
                    DataColumn col = dataManager.OriginalTable.Columns[colName];
                    lblType.Text = $"Тип данных: {GetDataTypeDescription(col.DataType)}";
                }
            };

            if (cmbColumn.Items.Count > 0)
            {
                string colName = cmbColumn.SelectedItem.ToString();
                if (dataManager.OriginalTable.Columns.Contains(colName))
                {
                    DataColumn col = dataManager.OriginalTable.Columns[colName];
                    lblType.Text = $"Тип данных: {GetDataTypeDescription(col.DataType)}";
                }
            }

            return dialog;
        }

        private void ResetSort()
        {
            try
            {
                if (dataGridViewMain.DataSource is DataView dv)
                    dv.Sort = "";
                dataManager.ResetSort();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сброса сортировки: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region Поиск

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (dataManager.OriginalTable == null)
            {
                MessageBox.Show("Сначала выберите таблицу.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using var searchDialog = CreateSearchDialog();

            if (searchDialog.ShowDialog() == DialogResult.OK)
            {
                var txtSearch = searchDialog.Controls.Find("txtSearch", true)[0] as TextBox;
                string searchText = txtSearch.Text.Trim();

                if (string.IsNullOrWhiteSpace(searchText))
                {
                    MessageBox.Show("Введите текст для поиска.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                try
                {
                    var visibleColumnNames = dataGridViewMain.Columns.Cast<DataGridViewColumn>()
                        .Where(c => c.Visible)
                        .Select(c => c.Name);

                    DataView dv = dataManager.PerformSearch(searchText, visibleColumnNames);

                    if (dv == null)
                    {
                        MessageBox.Show("Нет подходящих столбцов для поиска.", "Внимание",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    dataGridViewMain.DataSource = dv;

                    MessageBox.Show($"Найдено записей: {dv.Count}", "Результаты поиска",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при поиске:\n{ex.Message}", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private Form CreateSearchDialog()
        {
            var dialog = new Form
            {
                Text = "Поиск по таблице",
                Size = new System.Drawing.Size(400, 160),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var lblSearch = new Label { Text = "Введите текст или число:", Location = new System.Drawing.Point(15, 15), AutoSize = true };
            var txtSearch = new TextBox { Name = "txtSearch", Location = new System.Drawing.Point(15, 40), Width = 350 };
            var btnOk = new Button { Text = "Найти", DialogResult = DialogResult.OK, Location = new System.Drawing.Point(60, 80), Width = 90, Height = 30 };
            var btnReset = new Button { Text = "Сбросить поиск", Location = new System.Drawing.Point(160, 80), Width = 110, Height = 30 };
            btnReset.Click += (s, ev) => { ResetSearch(); dialog.Close(); };
            var btnCancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, Location = new System.Drawing.Point(280, 80), Width = 90, Height = 30 };

            dialog.Controls.AddRange(new Control[] { lblSearch, txtSearch, btnOk, btnReset, btnCancel });
            dialog.AcceptButton = btnOk;
            dialog.CancelButton = btnCancel;

            return dialog;
        }

        private void ResetSearch()
        {
            try
            {
                if (dataGridViewMain.DataSource is DataView dv)
                    dv.RowFilter = "";
                dataManager.ResetSearch();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сброса поиска: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region Полный сброс

        private void ResetAllViewStates()
        {
            try
            {
                if (dataGridViewMain.DataSource is DataView)
                    dataGridViewMain.DataSource = dataManager.OriginalTable;

                foreach (DataGridViewColumn col in dataGridViewMain.Columns)
                    try { col.Visible = true; } catch { }
                foreach (DataGridViewRow row in dataGridViewMain.Rows)
                    try { row.Visible = true; } catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сброса: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region Сохранение

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!isAdminMode)
            {
                MessageBox.Show("Сначала включите режим редактирования (кнопка «Изменить»).",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            TrySaveChanges();
        }

        private bool TrySaveChanges()
        {
            try { dataGridViewMain.EndEdit(); } catch { }

            var result = dataManager.TrySaveChanges();

            if (result.Success)
            {
                MessageBox.Show("Изменения успешно сохранены в базу данных.",
                    "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }
            else
            {
                MessageBox.Show($"Ошибка сохранения:\n{result.ErrorMessage}",
                    "Ошибка базы данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        #endregion

        #region Отчёт

        private void btnCreateReport_Click(object sender, EventArgs e)
        {
            if (dataGridViewMain.DataSource == null)
            {
                MessageBox.Show("Нет данных для создания отчёта. Выберите таблицу.",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string userName = CurrentUser.IsLoggedIn ? CurrentUser.Login : "Неизвестно";
            string tableName = comboBoxTables.SelectedItem?.ToString() ?? "Таблица";

            var result = WordReportGenerator.CreateReport(dataGridViewMain, tableName, userName);

            if (result.Success)
            {
                MessageBox.Show($"Отчёт успешно создан!\n\nФайл сохранён:\n{result.FilePath}",
                    "Отчёт готов", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBoxIcon icon = result.IsWordNotInstalled ? MessageBoxIcon.Error : MessageBoxIcon.Error;
                MessageBox.Show(result.ErrorMessage, "Ошибка", MessageBoxButtons.OK, icon);
            }
        }

        #endregion

        #region Валидация и удаление (UI-логика)

        private void DataGridViewMain_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (!isAdminMode || e.RowIndex < 0) return;

            try
            {
                DataGridViewRow row = dataGridViewMain.Rows[e.RowIndex];
                object newValue = e.FormattedValue;

                if (newValue == null || string.IsNullOrWhiteSpace(newValue.ToString()))
                {
                    bool hasOtherValues = false;
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        if (i == e.ColumnIndex) continue;
                        object cellValue = row.Cells[i].Value;
                        if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        {
                            hasOtherValues = true;
                            break;
                        }
                    }

                    bool wasFilled = row.Cells[e.ColumnIndex].Value != null &&
                                    !string.IsNullOrWhiteSpace(row.Cells[e.ColumnIndex].Value.ToString());

                    if (hasOtherValues || wasFilled)
                    {
                        string colName = dataGridViewMain.Columns[e.ColumnIndex].HeaderText;
                        e.Cancel = true;
                        row.Cells[e.ColumnIndex].ErrorText = $"Нельзя очистить поле «{colName}»: строка содержит данные.";

                        MessageBox.Show(
                            $"Нельзя оставить поле «{colName}» пустым.\n\nСтрока содержит данные — либо заполните поле, либо удалите всю строку.",
                            "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    row.Cells[e.ColumnIndex].ErrorText = "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка валидации: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DataGridViewMain_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            if (!isAdminMode) return;

            DataRow row = GetCurrentDataRow(e.Row);
            if (row == null) return;

            var violations = dataManager.CheckForeignKeyViolations(row);

            if (violations.Count > 0)
            {
                e.Cancel = true;
                MessageBox.Show(
                    "Нельзя удалить строку — существуют связанные данные:\n\n" + string.Join("\n", violations) +
                    "\n\nСначала удалите связанные записи из дочерних таблиц.",
                    "Нарушение связей", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var result = MessageBox.Show(
                "Удалить выбранную строку?\nЭто действие нельзя отменить после сохранения.",
                "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.No)
                e.Cancel = true;
        }

        private void DataGridViewMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (!isAdminMode || e.KeyCode != Keys.Delete || dataGridViewMain.SelectedRows.Count == 0)
                return;

            foreach (DataGridViewRow dgvRow in dataGridViewMain.SelectedRows)
            {
                if (dgvRow.IsNewRow) continue;
                DataRow row = GetCurrentDataRow(dgvRow);
                if (row == null) continue;

                var violations = dataManager.CheckForeignKeyViolations(row);
                if (violations.Count > 0)
                {
                    MessageBox.Show(
                        "Нельзя удалить выбранные строки — есть связанные данные:\n\n" + string.Join("\n", violations) +
                        "\n\nСначала удалите связанные записи из дочерних таблиц.",
                        "Нарушение связей", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            var result = MessageBox.Show(
                $"Удалить выбранные строки ({dataGridViewMain.SelectedRows.Count})?",
                "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                foreach (DataGridViewRow dgvRow in dataGridViewMain.SelectedRows)
                {
                    if (dgvRow.IsNewRow) continue;
                    DataRow dataRow = GetCurrentDataRow(dgvRow);
                    if (dataRow != null && dataRow.RowState != DataRowState.Deleted)
                        dataRow.Delete();
                }
            }
        }

        private DataRow GetCurrentDataRow(DataGridViewRow dgvRow)
        {
            if (dgvRow.DataBoundItem is DataRowView drv)
                return drv.Row;
            return null;
        }

        #endregion

        #region Вспомогательные методы

        private string GetDataTypeDescription(Type type)
        {
            if (type == typeof(string)) return "Текст (А-Я / Я-А)";
            else if (IsNumericType(type)) return "Число (0-9 / 9-0)";
            else if (type == typeof(DateTime)) return "Дата (старые-новые / новые-старые)";
            else if (type == typeof(bool)) return "Логический (Да/Нет)";
            else return type.Name;
        }

        private bool IsNumericType(Type type)
        {
            return type == typeof(int) || type == typeof(long) || type == typeof(short) ||
                   type == typeof(byte) || type == typeof(decimal) || type == typeof(double) ||
                   type == typeof(float);
        }

        #endregion
    }
}