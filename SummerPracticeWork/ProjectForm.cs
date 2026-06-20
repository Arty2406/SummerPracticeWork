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

        // Словарь связей: ключ = дочерняя таблица, значение = список (связанная таблица, поле FK)
        private Dictionary<string, List<ForeignKeyInfo>> foreignKeys = new Dictionary<string, List<ForeignKeyInfo>>();

        private class ForeignKeyInfo
        {
            public string ParentTable { get; set; }   // Таблица, на которую ссылаются
            public string ParentColumn { get; set; }  // Поле PK в родительской таблице
            public string ChildTable { get; set; }    // Таблица с внешним ключом
            public string ChildColumn { get; set; }   // Поле FK в дочерней таблице
        }
        private readonly string connStr;

        /// <summary>
        /// Проверка, является ли текущий пользователь администратором.
        /// Использует данные из CurrentUser (устанавливаются при входе в систему).
        /// </summary>
        private bool IsAdmin()
        {
            // Если пользователь не залогинен — доступ запрещён
            if (!CurrentUser.IsLoggedIn)
                return false;

            // Проверяем роль
            return CurrentUser.IsAdmin;
        }

        /// <summary>
        /// Загрузка информации о связях между таблицами из БД Access.
        /// </summary>
        private void LoadForeignKeys()
        {
            foreignKeys.Clear();

            try
            {
                using var conn = new OleDbConnection(connStr);
                conn.Open();

                // Получаем схему внешних ключей
                var fkSchema = conn.GetOleDbSchemaTable(
                    OleDbSchemaGuid.Foreign_Keys,
                    new object[] { null, null, null });

                if (fkSchema == null) return;

                foreach (DataRow row in fkSchema.Rows)
                {
                    var fk = new ForeignKeyInfo
                    {
                        // PK_TABLE - родительская таблица (на которую ссылаются)
                        ParentTable = row["PK_TABLE_NAME"]?.ToString(),
                        ParentColumn = row["PK_COLUMN_NAME"]?.ToString(),
                        // FK_TABLE - дочерняя таблица (содержит внешний ключ)
                        ChildTable = row["FK_TABLE_NAME"]?.ToString(),
                        ChildColumn = row["FK_COLUMN_NAME"]?.ToString()
                    };

                    if (string.IsNullOrEmpty(fk.ParentTable) || string.IsNullOrEmpty(fk.ChildTable))
                        continue;

                    // Сохраняем: для родительской таблицы — список её "детей"
                    if (!foreignKeys.ContainsKey(fk.ParentTable))
                        foreignKeys[fk.ParentTable] = new List<ForeignKeyInfo>();
                    foreignKeys[fk.ParentTable].Add(fk);
                }
            }
            catch (Exception ex)
            {
                // Если не удалось получить связи — работаем без проверок
                MessageBox.Show(
                    $"Не удалось загрузить связи между таблицами:\n{ex.Message}\n\nПроверки связей будут отключены.",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Проверяет, есть ли в дочерних таблицах записи, ссылающиеся на указанную строку.
        /// Возвращает список сообщений о нарушениях.
        /// </summary>
        private List<string> CheckForeignKeyViolations(DataRow row)
        {
            var violations = new List<string>();

            if (!foreignKeys.ContainsKey(currentTableName))
                return violations;

            foreach (var fk in foreignKeys[currentTableName])
            {
                // Получаем значение первичного ключа удаляемой строки
                if (!row.Table.Columns.Contains(fk.ParentColumn))
                    continue;

                object pkValue = row[fk.ParentColumn];
                if (pkValue == null || pkValue == DBNull.Value)
                    continue;

                // Проверяем, есть ли в дочерней таблице записи с таким FK
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

            // формирование строки подключения
            string dbName = "CourseWork.accdb";
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string dbPath = Path.Combine(basePath, dbName);

            // проверка существования файла перед созданием строки
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

                if (tableNames.Count == 0)
                {
                    MessageBox.Show("В базе данных нет пользовательских таблиц (не считая системных).", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // привязка списка таблиц к ComboBox
                comboBoxTables.DataSource = tableNames;

                // выбор первой таблицы
                comboBoxTables.SelectedIndex = 0;

                // Показываем в заголовке, кто вошёл
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

                // получение схемы таблиц
                string[] restrictions = new string[4] { null, null, null, "TABLE" };
                var schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, restrictions);

                if (schema != null)
                {
                    foreach (DataRow row in schema.Rows)
                    {
                        string name = row["TABLE_NAME"].ToString();

                        // исключение системных таблиц Access
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

            currentTableName = comboBoxTables.SelectedItem.ToString();
            string sql = $"SELECT * FROM [{currentTableName}]";

            try
            {
                var conn = new OleDbConnection(connStr);
                dataAdapter = new OleDbDataAdapter(sql, conn);
                commandBuilder = new OleDbCommandBuilder(dataAdapter);

                var dt = new DataTable();
                dataAdapter.Fill(dt);

                originalTable = dt.Copy(); // копия для сброса
                dataGridViewMain.DataSource = dt;

                // Настраиваем DataGridView для режима редактирования
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

        private void dataGridViewMain_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        // функционал кнопки перехода в главное меню
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
                // ✅ Выходим из системы — очищаем данные пользователя
                CurrentUser.Logout();

                // Закрываем текущую форму (это завершит приложение, т.к. FormClosed вызывает Application.Exit)
                this.Close();
            }
        }

        private void OpenMainMenu()
        {
            // поик существующей MainForm среди открытых форм
            var mainForm = Application.OpenForms.OfType<MainForm>().FirstOrDefault();

            if (mainForm != null)
            {
                // если форма уже открыта, то просто нужно показать её пользователю
                mainForm.Show();
                mainForm.BringToFront();
                mainForm.Focus();
            }
            else
            {
                // если MainForm ещё не было, то создаётся новая
                mainForm = new MainForm();
                mainForm.Show();
            }
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            // Проверка прав администратора
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

            // Если уже в режиме редактирования — выходим из него
            if (isAdminMode)
            {
                ExitEditMode();
                return;
            }

            // Загружаем связи между таблицами (один раз)
            if (foreignKeys.Count == 0)
                LoadForeignKeys();

            // Включаем режим редактирования
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

        /// <summary>
        /// Выход из режима редактирования.
        /// </summary>
        private void ExitEditMode()
        {
            // Проверяем несохранённые изменения
            DataTable current = GetCurrentDataTable();
            if (current != null && current.GetChanges() != null)
            {
                var result = MessageBox.Show(
                    "Есть несохранённые изменения. Сохранить их перед выходом?",
                    "Несохранённые изменения",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    if (!TrySaveChanges()) return; // если не удалось — остаёмся в режиме
                }
                else if (result == DialogResult.Cancel)
                {
                    return; // не выходим
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

        // функционал фильтра
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

            // Группа выбора столбцов
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

            // Заполняем список столбцов
            foreach (DataColumn col in originalTable.Columns)
            {
                checkedListBox.Items.Add(col.ColumnName, true); // все выбраны по умолчанию
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
                Text = "Сбросить",
                Location = new System.Drawing.Point(190, 420),
                Size = new System.Drawing.Size(100, 35)
            };
            btnReset.Click += (s, ev) =>
            {
                dataGridViewMain.DataSource = originalTable;
                filterDialog.Close();
            };

            var btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(310, 420),
                Size = new System.Drawing.Size(100, 35)
            };

            filterDialog.Controls.AddRange(new Control[] { grpColumns, grpRows, btnApply, btnReset, btnCancel });
            filterDialog.AcceptButton = btnApply;
            filterDialog.CancelButton = btnCancel;

            if (filterDialog.ShowDialog() == DialogResult.OK)
            {
                // проверки
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

                // нормализация диапазона
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
                // создание новой таблицы с выбранными столбцами
                var filteredTable = new DataTable();

                // добавление только выбранных столбцов
                var selectedColumns = new List<string>();
                foreach (var item in checkedListBox.CheckedItems)
                {
                    string colName = item.ToString();
                    selectedColumns.Add(colName);
                    filteredTable.Columns.Add(colName, originalTable.Columns[colName].DataType);
                }

                // копирование строк из выбранного диапазона
                for (int i = fromRow - 1; i < toRow; i++) // -1 потому что индекс 0-based, а пользователь вводит 1-based
                {
                    DataRow sourceRow = originalTable.Rows[i];
                    DataRow newRow = filteredTable.NewRow();

                    foreach (string colName in selectedColumns)
                    {
                        newRow[colName] = sourceRow[colName];
                    }

                    filteredTable.Rows.Add(newRow);
                }

                dataGridViewMain.DataSource = filteredTable;

                MessageBox.Show(
                    $"Применён фильтр:\n• Столбцов: {selectedColumns.Count}\n• Строк: {filteredTable.Rows.Count} (с {fromRow} по {toRow})",
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

        // функционал сортировки
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

            // Получаем текущую таблицу (с учётом фильтра)
            DataTable currentTable = GetCurrentDataTable();
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

            // ВАЖНО: берём столбцы из ТЕКУЩЕЙ таблицы, а не из originalTable
            foreach (DataColumn col in currentTable.Columns)
            {
                cmbColumn.Items.Add(col.ColumnName);
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
                Text = "Сбросить",
                Location = new System.Drawing.Point(190, 250),
                Size = new System.Drawing.Size(110, 35)
            };
            btnReset.Click += (s, ev) =>
            {
                dataGridViewMain.DataSource = originalTable;
                sortDialog.Close();
            };

            var btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(310, 250),
                Size = new System.Drawing.Size(110, 35)
            };

            sortDialog.Controls.AddRange(new Control[] { grpColumn, grpSortType, btnApply, btnReset, btnCancel });
            sortDialog.AcceptButton = btnApply;
            sortDialog.CancelButton = btnCancel;

            // обновление текста типа данных при изменении выбора столбца
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

            // инициализация текста типа данных
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
                // Получаем текущую таблицу (с учётом фильтра)
                DataTable currentTable = GetCurrentDataTable();
                if (currentTable == null)
                {
                    MessageBox.Show("Нет данных для сортировки.", "Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Проверяем, что столбец существует в текущей таблице
                if (!currentTable.Columns.Contains(columnName))
                {
                    MessageBox.Show(
                        $"Столбец '{columnName}' отсутствует в текущем представлении таблицы.\n" +
                        "Возможно, он был скрыт фильтром.",
                        "Ошибка сортировки",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                DataView dv;

                // Если уже есть DataView — используем его (сохраняем фильтр)
                if (dataGridViewMain.DataSource is DataView existingDv)
                {
                    dv = existingDv;
                }
                else
                {
                    // Создаём новый DataView на основе ТЕКУЩЕЙ таблицы (а не originalTable!)
                    dv = new DataView(currentTable);
                }

                // Применяем сортировку
                string sortDirection = ascending ? "ASC" : "DESC";
                dv.Sort = $"[{columnName}] {sortDirection}";

                dataGridViewMain.DataSource = dv;

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

        private DataTable GetCurrentDataTable()
        {
            if (dataGridViewMain.DataSource is DataView dv)
                return dv.Table;
            else if (dataGridViewMain.DataSource is DataTable dt)
                return dt;
            return null;
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

        // функционал создания отчёта в Word
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

            // Получаем логин пользователя
            string userName = CurrentUser.IsLoggedIn ? CurrentUser.Login : "Неизвестно";
            string dateNow = DateTime.Now.ToString("dd.MM.yyyy");
            string tableName = comboBoxTables.SelectedItem?.ToString() ?? "Таблица";

            // Формируем имя файла
            string fileName = $"{userName}_Отчёт_{DateTime.Now:yyyyMMdd}.docx";
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string fullPath = Path.Combine(documentsPath, fileName);

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                // Создаём приложение Word
                wordApp = new Word.Application();
                wordApp.Visible = false; // работаем в фоне

                doc = wordApp.Documents.Add();
                Word.Selection selection = wordApp.Selection;

                // ===== Заголовок отчёта =====
                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                selection.Font.Size = 16;
                selection.Font.Bold = 1;
                selection.TypeText($"Отчёт по таблице «{tableName}»");
                selection.TypeParagraph();

                // ===== Информация об отчёте =====
                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                selection.Font.Size = 11;
                selection.Font.Bold = 0;
                selection.TypeText($"Пользователь: {userName}");
                selection.TypeParagraph();
                selection.TypeText($"Дата создания: {dateNow}");
                selection.TypeParagraph();
                selection.TypeText($"Источник: {tableName}");
                selection.TypeParagraph();
                selection.TypeParagraph();

                // ===== Таблица =====
                // Получаем реальные данные (учитываем DataView, если есть)
                DataView dv = null;
                DataTable dt = null;

                if (dataGridViewMain.DataSource is DataView dvSource)
                    dv = dvSource;
                else if (dataGridViewMain.DataSource is DataTable dtSource)
                    dt = dtSource;

                int rowCount = dataGridViewMain.Rows.Cast<DataGridViewRow>().Count(r => !r.IsNewRow);
                int colCount = dataGridViewMain.Columns.Count;

                if (rowCount == 0 || colCount == 0)
                {
                    MessageBox.Show("Таблица пуста. Нечего включать в отчёт.",
                        "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    doc.Close(false);
                    return;
                }

                // Создаём таблицу в Word (+1 строка для заголовков)
                Word.Table wordTable = doc.Tables.Add(
                    selection.Range, rowCount + 1, colCount);

                // Стиль таблицы
                wordTable.Borders.Enable = 1;

                // Заголовки столбцов
                for (int c = 0; c < colCount; c++)
                {
                    Word.Cell cell = wordTable.Cell(1, c + 1);
                    cell.Range.Text = dataGridViewMain.Columns[c].HeaderText;
                    cell.Range.Font.Bold = 1;
                    cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray15;
                }

                // Данные строк
                int currentRow = 2;
                foreach (DataGridViewRow dgvRow in dataGridViewMain.Rows)
                {
                    if (dgvRow.IsNewRow) continue;

                    for (int c = 0; c < colCount; c++)
                    {
                        Word.Cell cell = wordTable.Cell(currentRow, c + 1);
                        object cellValue = dgvRow.Cells[c].Value;
                        cell.Range.Text = cellValue?.ToString() ?? "";
                    }
                    currentRow++;
                }

                // Автоподбор ширины столбцов
                wordTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

                // Сохраняем документ
                doc.SaveAs2(fullPath, Word.WdSaveFormat.wdFormatXMLDocument);

                // Показываем документ пользователю
                wordApp.Visible = true;
                wordApp.Activate();

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

                // Закрываем Word, если он открылся
                try
                {
                    doc?.Close(false);
                    wordApp?.Quit(false);
                }
                catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при создании отчёта:\n{ex.Message}",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                    );

                try
                {
                    doc?.Close(false);
                    wordApp?.Quit(false);
                }
                catch { }
            }
            finally
            {
                // Освобождаем COM-объекты
                if (doc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                    doc = null;
                }

                // Если Word не был показан пользователю (ошибка или что-то пошло не так) — закрываем его
                if (wordApp != null)
                {
                    // Если Word невидим — значит, мы не показали его пользователю, нужно закрыть
                    if (!wordApp.Visible)
                    {
                        wordApp.Quit(false);
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
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

        /// <summary>
        /// Попытка сохранить изменения в БД. Возвращает true при успехе.
        /// </summary>
        private bool TrySaveChanges()
        {
            if (dataAdapter == null || commandBuilder == null)
                return false;

            DataTable current = GetCurrentDataTable();
            if (current == null) return false;

            try
            {
                // Завершаем редактирование текущей ячейки
                dataGridViewMain.EndEdit();

                // Получаем изменённые строки для проверки FK
                DataTable changes = current.GetChanges();
                if (changes != null)
                {
                    // Проверяем удалённые строки на нарушение связей
                    foreach (DataRow row in changes.Rows)
                    {
                        if (row.RowState == DataRowState.Deleted)
                        {
                            // Для удалённой строки нужно получить оригинальные значения
                            var violations = CheckForeignKeyViolations(row);
                            if (violations.Count > 0)
                            {
                                string msg = "Нельзя удалить запись — существуют связанные данные в других таблицах:\n\n"
                                    + string.Join("\n", violations)
                                    + "\n\nСначала удалите связанные записи из дочерних таблиц.";

                                MessageBox.Show(msg, "Нарушение связей",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);

                                current.RejectChanges();
                                return false;
                            }
                        }
                    }
                }

                // Сохраняем в БД
                dataAdapter.Update(current);
                current.AcceptChanges();

                // Обновляем originalTable
                originalTable = current.Copy();

                MessageBox.Show("Изменения успешно сохранены в базу данных.",
                    "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }
            catch (OleDbException ex)
            {
                // Специальная обработка ошибок Access
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

                try { current.RejectChanges(); } catch { }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения:\n{ex.Message}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                try { current.RejectChanges(); } catch { }
                return false;
            }
        }

        /// <summary>
        /// Проверка ячейки перед завершением редактирования.
        /// Нельзя оставить ячейку пустой, если другие ячейки строки заполнены.
        /// </summary>
        private void DataGridViewMain_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (!isAdminMode) return;
            if (e.RowIndex < 0) return;

            DataGridViewRow row = dataGridViewMain.Rows[e.RowIndex];
            object newValue = e.FormattedValue;

            // Если пользователь пытается очистить ячейку
            if (newValue == null || string.IsNullOrWhiteSpace(newValue.ToString()))
            {
                // Проверяем, есть ли в строке другие заполненные ячейки
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

                // Проверяем, была ли ячейка ранее заполнена (т.е. это не новая строка)
                bool wasFilled = row.Cells[e.ColumnIndex].Value != null &&
                                !string.IsNullOrWhiteSpace(row.Cells[e.ColumnIndex].Value.ToString());

                // Если в строке есть другие значения ИЛИ ячейка была заполнена — запрещаем очистку
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
                // Очищаем ошибку, если значение валидно
                row.Cells[e.ColumnIndex].ErrorText = "";
            }
        }

        /// <summary>
        /// Проверка перед удалением строки — учитывает связующие ключи.
        /// </summary>
        private void DataGridViewMain_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            if (!isAdminMode) return;

            DataRow row = GetCurrentDataRow(e.Row);
            if (row == null) return;

            // Проверяем связующие ключи
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

            // Подтверждение удаления
            var result = MessageBox.Show(
                "Удалить выбранную строку?\nЭто действие нельзя отменить после сохранения.",
                "Подтверждение удаления",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.No)
                e.Cancel = true;
        }

        /// <summary>
        /// Обработка клавиши Delete для удаления строк.
        /// </summary>
        private void DataGridViewMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (!isAdminMode) return;

            if (e.KeyCode == Keys.Delete && dataGridViewMain.SelectedRows.Count > 0)
            {
                // Проверяем все выбранные строки на связующие ключи
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

                        // ПРАВИЛЬНОЕ удаление через DataRow
                        DataRow dataRow = GetCurrentDataRow(dgvRow);
                        if (dataRow != null)
                        {
                            dataRow.Delete(); // помечает строку как удалённую в DataTable
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Получает DataRow из DataGridViewRow.
        /// </summary>
        private DataRow GetCurrentDataRow(DataGridViewRow dgvRow)
        {
            if (dgvRow.DataBoundItem is DataRowView drv)
                return drv.Row;
            return null;
        }

        // функционал поиска
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
                Text = "Сбросить",
                Location = new System.Drawing.Point(160, 80),
                Width = 90,
                Height = 30
            };

            var btnCancel = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(260, 80),
                Width = 90,
                Height = 30
            };

            searchDialog.Controls.AddRange(new Control[] { lblSearch, txtSearch, btnOk, btnReset, btnCancel });
            searchDialog.AcceptButton = btnOk;
            searchDialog.CancelButton = btnCancel;

            btnReset.Click += (s, ev) =>
            {
                dataGridViewMain.DataSource = originalTable;
                searchDialog.Close();
            };

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

                foreach (DataColumn col in originalTable.Columns)
                {
                    // строковые поля - поиск через LIKE
                    if (col.DataType == typeof(string))
                    {
                        conditions.Add($"[{col.ColumnName}] LIKE '*{escapedText}*'");
                    }
                    // числовые поля - только точное совпадение
                    else if (IsNumericType(col.DataType) && isNumber)
                    {
                        string numStr = numValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        conditions.Add($"[{col.ColumnName}] = {numStr}");
                    }
                    // даты - точное совпадение
                    else if (col.DataType == typeof(DateTime) && isDate)
                    {
                        conditions.Add($"[{col.ColumnName}] = #{dateValue:MM/dd/yyyy}#");
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
                DataView dv = new DataView(originalTable) { RowFilter = filter };
                dataGridViewMain.DataSource = dv;

                MessageBox.Show(
                    $"Найдено записей: {dv.Count}", "Результаты поиска",
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
