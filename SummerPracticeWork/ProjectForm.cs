using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SummerPractice
{
    public partial class ProjectForm : Form
    {
        private DataTable originalTable;
        private OleDbDataAdapter dataAdapter;
        private OleDbCommandBuilder commandBuilder;
        private string currentTableName;
        private bool isAdminMode = false;

        private bool isFiltered = false;   // применён ли фильтр (видимость строк/столбцов)
        private bool isSorted = false;     // применена ли сортировка
        private bool isSearched = false;   // применён ли поиск
        private string currentSortColumn = null;
        private bool currentSortAscending = true;

        private Dictionary<string, List<ForeignKeyInfo>> foreignKeys = new Dictionary<string, List<ForeignKeyInfo>>();

        private class ForeignKeyInfo
        {
            public string ParentTable { get; set; }
            public string ParentColumn { get; set; }
            public string ChildTable { get; set; }
            public string ChildColumn { get; set; }
        }
        private readonly string connStr;

        private bool IsAdmin()
        {
            if (!CurrentUser.IsLoggedIn)
                return false;
            return CurrentUser.IsAdmin;
        }

        private void LoadForeignKeys()
        {
            foreignKeys.Clear();

            try
            {
                using var conn = new OleDbConnection(connStr);
                conn.Open();

                var fkSchema = conn.GetOleDbSchemaTable(
                    OleDbSchemaGuid.Foreign_Keys,
                    new object[] { null, null, null });

                if (fkSchema == null) return;

                foreach (DataRow row in fkSchema.Rows)
                {
                    var fk = new ForeignKeyInfo
                    {
                        ParentTable = row["PK_TABLE_NAME"]?.ToString(),
                        ParentColumn = row["PK_COLUMN_NAME"]?.ToString(),
                        ChildTable = row["FK_TABLE_NAME"]?.ToString(),
                        ChildColumn = row["FK_COLUMN_NAME"]?.ToString()
                    };

                    if (string.IsNullOrEmpty(fk.ParentTable) || string.IsNullOrEmpty(fk.ChildTable))
                        continue;

                    if (!foreignKeys.ContainsKey(fk.ParentTable))
                        foreignKeys[fk.ParentTable] = new List<ForeignKeyInfo>();
                    foreignKeys[fk.ParentTable].Add(fk);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось загрузить связи между таблицами:\n{ex.Message}\n\nПроверки связей будут отключены.",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private List<string> CheckForeignKeyViolations(DataRow row)
        {
            var violations = new List<string>();

            if (!foreignKeys.ContainsKey(currentTableName))
                return violations;

            foreach (var fk in foreignKeys[currentTableName])
            {
                if (fk.ParentTable != currentTableName)
                    continue;

                if (!row.Table.Columns.Contains(fk.ParentColumn))
                    continue;

                object pkValue;
                try
                {
                    pkValue = row[fk.ParentColumn, DataRowVersion.Original];
                }
                catch
                {
                    try { pkValue = row[fk.ParentColumn]; }
                    catch { continue; }
                }

                if (pkValue == null || pkValue == DBNull.Value)
                    continue;

                string checkSql = $"SELECT COUNT(*) FROM [{fk.ChildTable}] WHERE [{fk.ChildColumn}] = ?";

                try
                {
                    using var conn = new OleDbConnection(connStr);
                    using var cmd = new OleDbCommand(checkSql, conn);
                    cmd.Parameters.AddWithValue("@p1", pkValue);
                    conn.Open();

                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    if (count > 0)
                    {
                        violations.Add(
                            $"• В таблице «{fk.ChildTable}» есть {count} записей, " +
                            $"ссылающихся на эту (поле «{fk.ChildColumn}» = {pkValue}).");
                    }
                }
                catch (Exception ex)
                {
                    violations.Add($"• Ошибка проверки таблицы «{fk.ChildTable}»: {ex.Message}");
                }
            }

            return violations;
        }

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
                    "Ошибка БД", MessageBoxButtons.OK, MessageBoxIcon.Error
                    );
                this.Close();
                return;
            }

            connStr = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";

            this.Load += ProjectForm_Load;
        }

        private void ProjectForm_Load(object sender, EventArgs e)
        {
            try
            {
                var tableNames = GetTableNames(connStr);

                if (!IsAdmin())
                {
                    tableNames = tableNames.Where(t => !t.Equals("Пользователи", StringComparison.OrdinalIgnoreCase)).ToList();
                }

                if (tableNames.Count == 0)
                {
                    MessageBox.Show("В базе данных нет доступных вам таблиц.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                comboBoxTables.DataSource = tableNames;
                comboBoxTables.SelectedIndex = 0;

                if (CurrentUser.IsLoggedIn)
                {
                    this.Text = $"Система для работы с таблицами Access — {CurrentUser.GetDisplayName()}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось подключиться к базе данных.\n\nПроверьте:\n1) Файл БД лежит рядом с программой (.exe).\n2) Установлен ли Access Database Engine (разрядность x86/x64 должна совпадать с проектом).\n\nОшибка: {ex.Message}",
                    "Критическая ошибка подключения",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                    );

                this.Close();
            }
        }

        private List<string> GetTableNames(string connStr)
        {
            var list = new List<string>();
            using var conn = new OleDbConnection(connStr);

            try
            {
                conn.Open();

                string[] restrictions = new string[4] { null, null, null, "TABLE" };
                var schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, restrictions);

                if (schema != null)
                {
                    foreach (DataRow row in schema.Rows)
                    {
                        string name = row["TABLE_NAME"].ToString();

                        if (!name.StartsWith("MSys") && !name.StartsWith("USys"))
                            list.Add(name);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }

            return list.OrderBy(x => x).ToList();
        }

        private void ComboBoxTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxTables.SelectedItem == null) return;

            if (originalTable != null && originalTable.GetChanges() != null)
            {
                var result = MessageBox.Show(
                    "Есть несохранённые изменения в текущей таблице. Сохранить их?",
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
            }

            currentTableName = comboBoxTables.SelectedItem.ToString();
            string sql = $"SELECT * FROM [{currentTableName}]";

            try
            {
                var conn = new OleDbConnection(connStr);
                dataAdapter = new OleDbDataAdapter(sql, conn);
                commandBuilder = new OleDbCommandBuilder(dataAdapter);

                var dt = new DataTable();
                dataAdapter.Fill(dt);

                originalTable = dt;
                dataGridViewMain.DataSource = dt;

                foreach (DataGridViewColumn col in dataGridViewMain.Columns)
                {
                    if (col.ValueType == typeof(DateTime))
                    {
                        col.DefaultCellStyle.Format = "dd.MM.yyyy";
                        col.HeaderCell.Style.Format = "dd.MM.yyyy";
                    }
                }

                isFiltered = false;
                isSorted = false;
                isSearched = false;
                currentSortColumn = null;

                dataGridViewMain.ReadOnly = !isAdminMode;
                dataGridViewMain.AllowUserToAddRows = isAdminMode;
                dataGridViewMain.AllowUserToDeleteRows = isAdminMode;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при чтении таблицы '{currentTableName}':\n{ex.Message}",
                    "Ошибка данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExitToMain_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "Вернуться в главное меню?\n\n" +
                "Вы будете разлогинены и вернётесь к форме входа.",
                "Подтверждение",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
                );

            if (result == DialogResult.Yes)
            {
                CurrentUser.Logout();
                this.Close();
            }
        }

        private void OpenMainMenu()
        {
            var mainForm = Application.OpenForms.OfType<MainForm>().FirstOrDefault();

            if (mainForm != null)
            {
                mainForm.Show();
                mainForm.BringToFront();
                mainForm.Focus();
            }
            else
            {
                mainForm = new MainForm();
                mainForm.Show();
            }
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            if (!IsAdmin())
            {
                MessageBox.Show(
                    "Эта функция доступна только администратору.\n\n" +
                    "Войдите в систему под учётной записью администратора.",
                    "Доступ запрещён",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Stop);
                return;
            }

            if (isAdminMode)
            {
                ExitEditMode();
                return;
            }

            ResetAllViewStates();

            if (foreignKeys.Count == 0)
                LoadForeignKeys();

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
                "• После изменений нажмите «Сохранить» (кнопка рядом).\n\n" +
                "Внимание: учитываются связи между таблицами!",
                "Режим редактирования",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void ResetFilter()
        {
            try
            {
                foreach (DataGridViewColumn col in dataGridViewMain.Columns)
                {
                    try { col.Visible = true; } catch { }
                }

                foreach (DataGridViewRow row in dataGridViewMain.Rows)
                {
                    try { row.Visible = true; } catch { }
                }

                isFiltered = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сброса фильтра: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ResetSort()
        {
            try
            {
                if (dataGridViewMain.DataSource is DataView dv)
                {
                    dv.Sort = "";
                }

                isSorted = false;
                currentSortColumn = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сброса сортировки: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ResetSearch()
        {
            try
            {
                if (dataGridViewMain.DataSource is DataView dv)
                {
                    dv.RowFilter = "";
                }

                isSearched = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сброса поиска: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        
        private void ResetAllViewStates()
        {
            try
            {
                if (dataGridViewMain.DataSource is DataView)
                {
                    dataGridViewMain.DataSource = originalTable;
                }

                foreach (DataGridViewColumn col in dataGridViewMain.Columns)
                {
                    try { col.Visible = true; } catch { }
                }

                foreach (DataGridViewRow row in dataGridViewMain.Rows)
                {
                    try { row.Visible = true; } catch { }
                }

                isFiltered = false;
                isSorted = false;
                isSearched = false;
                currentSortColumn = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сброса: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ExitEditMode()
        {
            if (originalTable != null && originalTable.GetChanges() != null)
            {
                var result = MessageBox.Show(
                    "Есть несохранённые изменения. Сохранить их перед выходом?",
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

        private void btnFilter_Click(object sender, EventArgs e)
        {
            if (originalTable == null)
            {
                MessageBox.Show("Сначала выберите таблицу.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using var filterDialog = new Form
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
                Location = new System.Drawing.Point(15, 25),
                Size = new System.Drawing.Size(405, 220),
                CheckOnClick = true
            };

            foreach (DataColumn col in originalTable.Columns)
            {
                checkedListBox.Items.Add(col.ColumnName, true);
            }

            var btnSelectAll = new Button
            {
                Text = "Выбрать все",
                Location = new System.Drawing.Point(15, 255),
                Size = new System.Drawing.Size(130, 30)
            };
            btnSelectAll.Click += (s, ev) =>
            {
                for (int i = 0; i < checkedListBox.Items.Count; i++)
                    checkedListBox.SetItemChecked(i, true);
            };

            var btnDeselectAll = new Button
            {
                Text = "Снять все",
                Location = new System.Drawing.Point(155, 255),
                Size = new System.Drawing.Size(130, 30)
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
                Location = new System.Drawing.Point(15, 320),
                Size = new System.Drawing.Size(435, 80)
            };

            var lblFrom = new Label
            {
                Text = "С:",
                Location = new System.Drawing.Point(15, 30),
                AutoSize = true
            };
            var txtFrom = new TextBox
            {
                Text = "1",
                Location = new System.Drawing.Point(40, 27),
                Width = 70
            };

            var lblTo = new Label
            {
                Text = "По:",
                Location = new System.Drawing.Point(120, 30),
                AutoSize = true
            };
            var txtTo = new TextBox
            {
                Text = originalTable.Rows.Count.ToString(),
                Location = new System.Drawing.Point(160, 27),
                Width = 70
            };

            var lblInfo = new Label
            {
                Text = $"Всего строк: {originalTable.Rows.Count}",
                Location = new System.Drawing.Point(260, 30),
                AutoSize = true,
                ForeColor = System.Drawing.Color.Gray
            };

            grpRows.Controls.AddRange(new Control[] { lblFrom, txtFrom, lblTo, txtTo, lblInfo });

            var btnApply = new Button
            {
                Text = "Применить",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(70, 420),
                Size = new System.Drawing.Size(100, 35)
            };

            var btnReset = new Button
            {
                Text = "Сбросить фильтр",
                Location = new System.Drawing.Point(190, 420),
                Size = new System.Drawing.Size(120, 35)
            };
            btnReset.Click += (s, ev) =>
            {
                ResetFilter();
                filterDialog.Close();
            };

            var btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(320, 420),
                Size = new System.Drawing.Size(100, 35)
            };

            filterDialog.Controls.AddRange(new Control[] { grpColumns, grpRows, btnApply, btnReset, btnCancel });
            filterDialog.AcceptButton = btnApply;
            filterDialog.CancelButton = btnCancel;

            if (filterDialog.ShowDialog() == DialogResult.OK)
            {
                if (checkedListBox.CheckedItems.Count == 0)
                {
                    MessageBox.Show(
                        "Выберите хотя бы один столбец для отображения.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning
                        );
                    return;
                }

                if (!int.TryParse(txtFrom.Text, out int fromRow) || !int.TryParse(txtTo.Text, out int toRow))
                {
                    MessageBox.Show(
                        "Введите корректные номера строк (числа).", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning
                        );
                    return;
                }

                fromRow = Math.Max(1, fromRow);
                toRow = Math.Min(originalTable.Rows.Count, toRow);

                if (fromRow > toRow)
                {
                    MessageBox.Show(
                        "Начальная строка не может быть больше конечной.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning
                        );
                    return;
                }

                ApplyFilter(checkedListBox, fromRow, toRow);
            }
        }

        private void ApplyFilter(CheckedListBox checkedListBox, int fromRow, int toRow)
        {
            try
            {
                dataGridViewMain.EndEdit();

                if (dataGridViewMain.DataSource is DataView)
                {
                    dataGridViewMain.DataSource = originalTable;
                    isSorted = false;
                    isSearched = false;
                }

                foreach (DataGridViewColumn col in dataGridViewMain.Columns)
                {
                    try { col.Visible = true; } catch { }
                }
                foreach (DataGridViewRow row in dataGridViewMain.Rows)
                {
                    try { row.Visible = true; } catch { }
                }

                var selectedColumns = new HashSet<string>();
                foreach (var item in checkedListBox.CheckedItems)
                {
                    selectedColumns.Add(item.ToString());
                }

                foreach (DataGridViewColumn col in dataGridViewMain.Columns)
                {
                    col.Visible = selectedColumns.Contains(col.Name);
                }

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

                isFiltered = true;

                MessageBox.Show(
                    $"Применён фильтр:\n• Столбцов отображается: {selectedColumns.Count}\n• Строк отображается: {visibleCount} (с {fromRow} по {toRow})",
                    "Фильтр применён",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                    );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при применении фильтра:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error
                    );
            }
        }

        private void btnSort_Click(object sender, EventArgs e)
        {
            if (dataGridViewMain.DataSource == null)
            {
                MessageBox.Show(
                    "Сначала выберите таблицу.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning
                    );
                return;
            }

            DataTable currentTable = originalTable;
            if (currentTable == null)
            {
                MessageBox.Show("Нет данных для сортировки.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using var sortDialog = new Form
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

            foreach (DataGridViewColumn dgvCol in dataGridViewMain.Columns)
            {
                if (dgvCol.Visible)
                {
                    cmbColumn.Items.Add(dgvCol.Name);
                }
            }
            if (cmbColumn.Items.Count > 0) cmbColumn.SelectedIndex = 0;

            grpColumn.Controls.AddRange(new Control[] { cmbColumn, lblType });

            var grpSortType = new GroupBox
            {
                Text = "Тип сортировки:",
                Location = new System.Drawing.Point(15, 130),
                Size = new System.Drawing.Size(415, 100)
            };

            var rbAscending = new RadioButton
            {
                Text = "По возрастанию (А→Я, 0→9, старые→новые)",
                Location = new System.Drawing.Point(15, 30),
                AutoSize = true,
                Checked = true
            };

            var rbDescending = new RadioButton
            {
                Text = "По убыванию (Я→А, 9→0, новые→старые)",
                Location = new System.Drawing.Point(15, 60),
                AutoSize = true
            };

            grpSortType.Controls.AddRange(new Control[] { rbAscending, rbDescending });

            var btnApply = new Button
            {
                Text = "Применить",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(70, 250),
                Size = new System.Drawing.Size(110, 35)
            };

            var btnReset = new Button
            {
                Text = "Сбросить сортировку",
                Location = new System.Drawing.Point(190, 250),
                Size = new System.Drawing.Size(130, 35)
            };
            btnReset.Click += (s, ev) =>
            {
                ResetSort();
                sortDialog.Close();
            };

            var btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(330, 250),
                Size = new System.Drawing.Size(110, 35)
            };

            sortDialog.Controls.AddRange(new Control[] { grpColumn, grpSortType, btnApply, btnReset, btnCancel });
            sortDialog.AcceptButton = btnApply;
            sortDialog.CancelButton = btnCancel;

            cmbColumn.SelectedIndexChanged += (s, ev) =>
            {
                if (cmbColumn.SelectedItem == null) return;
                string colName = cmbColumn.SelectedItem.ToString();

                if (currentTable.Columns.Contains(colName))
                {
                    DataColumn col = currentTable.Columns[colName];
                    lblType.Text = $"Тип данных: {GetDataTypeDescription(col.DataType)}";
                }
            };

            if (cmbColumn.Items.Count > 0)
            {
                string colName = cmbColumn.SelectedItem.ToString();
                if (currentTable.Columns.Contains(colName))
                {
                    DataColumn col = currentTable.Columns[colName];
                    lblType.Text = $"Тип данных: {GetDataTypeDescription(col.DataType)}";
                }
            }

            if (sortDialog.ShowDialog() == DialogResult.OK)
            {
                if (cmbColumn.SelectedItem == null)
                {
                    MessageBox.Show(
                        "Выберите столбец для сортировки.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning
                        );
                    return;
                }

                string columnName = cmbColumn.SelectedItem.ToString();
                bool ascending = rbAscending.Checked;

                ApplySort(columnName, ascending);
            }
        }

        private void ApplySort(string columnName, bool ascending)
        {
            try
            {
                if (originalTable == null)
                {
                    MessageBox.Show("Нет данных для сортировки.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!originalTable.Columns.Contains(columnName))
                {
                    MessageBox.Show(
                        $"Столбец '{columnName}' отсутствует в таблице.",
                        "Ошибка сортировки",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (dataGridViewMain.DataSource is DataTable)
                {
                    dataGridViewMain.DataSource = new DataView(originalTable);
                }

                if (dataGridViewMain.DataSource is DataView dv)
                {
                    string sortDirection = ascending ? "ASC" : "DESC";
                    dv.Sort = $"[{columnName}] {sortDirection}";
                }

                isSorted = true;
                currentSortColumn = columnName;
                currentSortAscending = ascending;

                string directionText = ascending ? "по возрастанию" : "по убыванию";
                MessageBox.Show(
                    $"Таблица отсортирована по столбцу '{columnName}' {directionText}.",
                    "Сортировка применена",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                    );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при сортировке:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error
                    );
            }
        }

        private string GetDataTypeDescription(Type type)
        {
            if (type == typeof(string))
                return "Текст (А-Я / Я-А)";
            else if (IsNumericType(type))
                return "Число (0-9 / 9-0)";
            else if (type == typeof(DateTime))
                return "Дата (старые-новые / новые-старые)";
            else if (type == typeof(bool))
                return "Логический (Да/Нет)";
            else
                return type.Name;
        }

        private void btnCreateReport_Click(object sender, EventArgs e)
        {
            if (dataGridViewMain.DataSource == null)
            {
                MessageBox.Show(
                    "Нет данных для создания отчёта. Выберите таблицу.",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning
                    );
                return;
            }

            string userName = CurrentUser.IsLoggedIn ? CurrentUser.Login : "Неизвестно";
            DateTime now = DateTime.Now;
            string dateNow = now.ToString("dd.MM.yyyy");
            string timeNow = now.ToString("HH:mm:ss");
            string tableName = comboBoxTables.SelectedItem?.ToString() ?? "Таблица";

            string fileName = $"{userName}_Отчёт_{now:yyyyMMdd_HH-mm-ss}.docx";
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string fullPath = Path.Combine(documentsPath, fileName);

            Word.Application wordApp = null;
            Word.Document doc = null;
            bool wordVisible = false;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;

                doc = wordApp.Documents.Add();
                Word.Selection selection = wordApp.Selection;

                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                selection.Font.Size = 16;
                selection.Font.Bold = 1;
                selection.TypeText($"Отчёт по таблице «{tableName}»");
                selection.TypeParagraph();

                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                selection.Font.Size = 11;
                selection.Font.Bold = 0;
                selection.TypeText($"Пользователь: {userName}");
                selection.TypeParagraph();
                selection.TypeText($"Дата создания: {dateNow}");
                selection.TypeParagraph();
                selection.TypeText($"Время создания: {timeNow}");
                selection.TypeParagraph();
                selection.TypeText($"Источник: {tableName}");
                selection.TypeParagraph();
                selection.TypeParagraph();

                var visibleRows = dataGridViewMain.Rows.Cast<DataGridViewRow>()
                    .Where(r => !r.IsNewRow && r.Visible)
                    .ToList();
                var visibleCols = dataGridViewMain.Columns.Cast<DataGridViewColumn>()
                    .Where(c => c.Visible)
                    .ToList();

                int rowCount = visibleRows.Count;
                int colCount = visibleCols.Count;

                if (rowCount == 0 || colCount == 0)
                {
                    MessageBox.Show("Таблица пуста. Нечего включать в отчёт.",
                        "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    doc.Close(false);
                    wordApp.Quit(false);
                    wordVisible = true;
                    return;
                }

                Word.Table wordTable = doc.Tables.Add(
                    selection.Range, rowCount + 1, colCount);

                wordTable.Borders.Enable = 1;

                for (int c = 0; c < colCount; c++)
                {
                    Word.Cell cell = wordTable.Cell(1, c + 1);
                    cell.Range.Text = visibleCols[c].HeaderText;
                    cell.Range.Font.Bold = 1;
                    cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray15;
                }

                int currentRow = 2;
                foreach (DataGridViewRow dgvRow in visibleRows)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        Word.Cell cell = wordTable.Cell(currentRow, c + 1);
                        object cellValue = dgvRow.Cells[visibleCols[c].Index].Value;

                        if (cellValue is DateTime dt)
                        {
                            cell.Range.Text = dt.ToString("dd.MM.yyyy");
                        }
                        else
                        {
                            cell.Range.Text = cellValue?.ToString() ?? "";
                        }
                    }
                    currentRow++;
                }

                wordTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

                doc.SaveAs2(fullPath, Word.WdSaveFormat.wdFormatXMLDocument);

                wordApp.Visible = true;
                wordApp.Activate();
                wordVisible = true;

                MessageBox.Show(
                    $"Отчёт успешно создан!\n\nФайл сохранён:\n{fullPath}",
                    "Отчёт готов",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                    );
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show(
                    $"Не удалось создать отчёт в Word.\n\n" +
                    $"Возможно, Microsoft Word не установлен на этом компьютере.\n\n" +
                    $"Ошибка: {ex.Message}",
                    "Ошибка Word",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                    );

                try { doc?.Close(false); } catch { }

                if (wordApp != null && !wordVisible)
                {
                    try { wordApp.Quit(false); } catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при создании отчёта:\n{ex.Message}",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                    );

                try { doc?.Close(false); } catch { }

                if (wordApp != null && !wordVisible)
                {
                    try { wordApp.Quit(false); } catch { }
                }
            }
            finally
            {
                if (!wordVisible)
                {
                    if (doc != null)
                    {
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(doc); } catch { }
                        doc = null;
                    }

                    if (wordApp != null)
                    {
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp); } catch { }
                        wordApp = null;
                    }
                }
                else
                {
                    doc = null;
                    wordApp = null;
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!isAdminMode)
            {
                MessageBox.Show(
                    "Сначала включите режим редактирования (кнопка «Изменить»).",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            TrySaveChanges();
        }

        private bool TrySaveChanges()
        {
            if (dataAdapter == null || commandBuilder == null || originalTable == null)
                return false;

            try
            {
                try { dataGridViewMain.EndEdit(); } catch { }

                DataTable changes = originalTable.GetChanges();
                if (changes != null)
                {
                    foreach (DataRow row in changes.Rows)
                    {
                        if (row.RowState == DataRowState.Deleted)
                        {
                            var violations = CheckForeignKeyViolations(row);
                            if (violations.Count > 0)
                            {
                                string msg = "Нельзя удалить запись — существуют связанные данные в других таблицах:\n\n"
                                    + string.Join("\n", violations)
                                    + "\n\nСначала удалите связанные записи из дочерних таблиц.";

                                MessageBox.Show(msg, "Нарушение связей",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);

                                originalTable.RejectChanges();
                                return false;
                            }
                        }
                    }
                }

                dataAdapter.Update(originalTable);
                originalTable.AcceptChanges();

                MessageBox.Show("Изменения успешно сохранены в базу данных.",
                    "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }
            catch (OleDbException ex)
            {
                string errorMsg = ex.Message;

                if (errorMsg.Contains("relationship") || errorMsg.Contains("связ") || ex.ErrorCode == -2147217900)
                {
                    errorMsg = "Нарушена целостность связей между таблицами.\n\n" +
                               "Возможно, вы пытаетесь удалить запись, на которую ссылаются из другой таблицы, " +
                               "или изменить значение первичного ключа.\n\n" +
                               "Сначала удалите/измените связанные записи в дочерних таблицах.";
                }
                else if (errorMsg.Contains("Null") || errorMsg.Contains("null"))
                {
                    errorMsg = "Нельзя оставить обязательное поле пустым.\n\n" +
                               "Проверьте, все ли обязательные поля заполнены.";
                }

                MessageBox.Show($"Ошибка сохранения:\n{errorMsg}",
                    "Ошибка базы данных", MessageBoxButtons.OK, MessageBoxIcon.Error);

                try { originalTable.RejectChanges(); } catch { }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения:\n{ex.Message}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                try { originalTable.RejectChanges(); } catch { }
                return false;
            }
        }

        private void DataGridViewMain_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (!isAdminMode) return;
            if (e.RowIndex < 0) return;

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
                            $"Нельзя оставить поле «{colName}» пустым.\n\n" +
                            "Строка содержит данные — либо заполните поле, либо удалите всю строку.",
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

            var violations = CheckForeignKeyViolations(row);

            if (violations.Count > 0)
            {
                e.Cancel = true;

                string msg = "Нельзя удалить строку — существуют связанные данные:\n\n"
                    + string.Join("\n", violations)
                    + "\n\nСначала удалите связанные записи из дочерних таблиц.";

                MessageBox.Show(msg, "Нарушение связей",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var result = MessageBox.Show(
                "Удалить выбранную строку?\nЭто действие нельзя отменить после сохранения.",
                "Подтверждение удаления",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.No)
                e.Cancel = true;
        }

        private void DataGridViewMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (!isAdminMode) return;

            if (e.KeyCode == Keys.Delete && dataGridViewMain.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow dgvRow in dataGridViewMain.SelectedRows)
                {
                    if (dgvRow.IsNewRow) continue;

                    DataRow row = GetCurrentDataRow(dgvRow);
                    if (row == null) continue;

                    var violations = CheckForeignKeyViolations(row);
                    if (violations.Count > 0)
                    {
                        string msg = "Нельзя удалить выбранные строки — есть связанные данные:\n\n"
                            + string.Join("\n", violations)
                            + "\n\nСначала удалите связанные записи из дочерних таблиц.";

                        MessageBox.Show(msg, "Нарушение связей",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                var result = MessageBox.Show(
                    $"Удалить выбранные строки ({dataGridViewMain.SelectedRows.Count})?",
                    "Подтверждение удаления",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    foreach (DataGridViewRow dgvRow in dataGridViewMain.SelectedRows)
                    {
                        if (dgvRow.IsNewRow) continue;

                        DataRow dataRow = GetCurrentDataRow(dgvRow);
                        if (dataRow != null && dataRow.RowState != DataRowState.Deleted)
                        {
                            dataRow.Delete();
                        }
                    }
                }
            }
        }

        private DataRow GetCurrentDataRow(DataGridViewRow dgvRow)
        {
            if (dgvRow.DataBoundItem is DataRowView drv)
                return drv.Row;
            return null;
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (originalTable == null)
            {
                MessageBox.Show("Сначала выберите таблицу.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using var searchDialog = new Form
            {
                Text = "Поиск по таблице",
                Size = new System.Drawing.Size(400, 160),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var lblSearch = new Label
            {
                Text = "Введите текст или число:",
                Location = new System.Drawing.Point(15, 15),
                AutoSize = true
            };

            var txtSearch = new TextBox
            {
                Location = new System.Drawing.Point(15, 40),
                Width = 350
            };

            var btnOk = new Button
            {
                Text = "Найти",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(60, 80),
                Width = 90,
                Height = 30
            };

            var btnReset = new Button
            {
                Text = "Сбросить поиск",
                Location = new System.Drawing.Point(160, 80),
                Width = 110,
                Height = 30
            };
            btnReset.Click += (s, ev) =>
            {
                ResetSearch();
                searchDialog.Close();
            };

            var btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(280, 80),
                Width = 90,
                Height = 30
            };

            searchDialog.Controls.AddRange(new Control[] { lblSearch, txtSearch, btnOk, btnReset, btnCancel });
            searchDialog.AcceptButton = btnOk;
            searchDialog.CancelButton = btnCancel;

            if (searchDialog.ShowDialog() == DialogResult.OK)
            {
                string searchText = txtSearch.Text.Trim();

                if (string.IsNullOrWhiteSpace(searchText))
                {
                    MessageBox.Show(
                        "Введите текст для поиска.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning
                        );
                    return;
                }

                PerformSearch(searchText);
            }
        }

        private void PerformSearch(string searchText)
        {
            try
            {
                if (originalTable == null) return;

                var conditions = new List<string>();
                string escapedText = searchText.Replace("'", "''");
                bool isNumber = decimal.TryParse(searchText, out decimal numValue);
                bool isDate = DateTime.TryParse(searchText, out DateTime dateValue);

                foreach (DataGridViewColumn dgvCol in dataGridViewMain.Columns)
                {
                    if (!dgvCol.Visible) continue;

                    string colName = dgvCol.Name;
                    if (!originalTable.Columns.Contains(colName)) continue;

                    DataColumn col = originalTable.Columns[colName];

                    if (col.DataType == typeof(string))
                    {
                        conditions.Add($"[{colName}] LIKE '*{escapedText}*'");
                    }
                    else if (IsNumericType(col.DataType) && isNumber)
                    {
                        string numStr = numValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        conditions.Add($"[{colName}] = {numStr}");
                    }
                    else if (col.DataType == typeof(DateTime) && isDate)
                    {
                        conditions.Add($"[{colName}] = #{dateValue:MM/dd/yyyy}#");
                    }
                }

                if (conditions.Count == 0)
                {
                    MessageBox.Show(
                        "Нет подходящих столбцов для поиска.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Information
                        );
                    return;
                }

                string filter = string.Join(" OR ", conditions);

                if (dataGridViewMain.DataSource is DataTable)
                {
                    dataGridViewMain.DataSource = new DataView(originalTable);
                }

                if (dataGridViewMain.DataSource is DataView dv)
                {
                    dv.RowFilter = filter;
                }

                isSearched = true;

                MessageBox.Show(
                    $"Найдено записей: {(dataGridViewMain.DataSource as DataView)?.Count ?? 0}",
                    "Результаты поиска",
                    MessageBoxButtons.OK, MessageBoxIcon.Information
                    );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при поиске:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error
                    );
            }
        }

        private bool IsNumericType(Type type)
        {
            return type == typeof(int) || type == typeof(long) || type == typeof(short) ||
                   type == typeof(byte) || type == typeof(decimal) || type == typeof(double) ||
                   type == typeof(float);
        }
    }
}