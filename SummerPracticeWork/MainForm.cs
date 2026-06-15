using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SummerPractice
{
    public partial class MainForm : Form
    {
        private readonly string _connStr;

        public MainForm()
        {
            // 1. Сначала инициализируем компоненты (кнопки, поля), иначе будет ошибка NullReference
            InitializeComponent();

            // 2. Формируем строку подключения
            string dbName = "CourseWork.accdb";
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string dbPath = Path.Combine(basePath, dbName);

            // Проверка существования файла перед созданием строки (хорошая практика)
            if (!File.Exists(dbPath))
            {
                MessageBox.Show($"Файл базы данных не найден по пути:\n{dbPath}\n\nУбедитесь, что файл скопирован в папку bin\\Debug\\netX.X",
                    "Ошибка БД", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Можно даже закрыть форму здесь, если БД обязательна
                this.Close();
                return;
            }

            _connStr = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";

            // 3. Настраиваем форму
            this.Text = "Главное меню";
            this.StartPosition = FormStartPosition.CenterScreen;

            // Подписываемся на событие загрузки ТОЛЬКО ЗДЕСЬ, не дублируем в дизайнере
            this.Load += MainForm_Load;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                var tableNames = GetTableNames(_connStr);

                if (tableNames.Count == 0)
                {
                    MessageBox.Show("В базе данных нет пользовательских таблиц (не считая системных).", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Привязываем список таблиц к ComboBox
                comboBoxTables.DataSource = tableNames;

                // Выбираем первую таблицу
                comboBoxTables.SelectedIndex = 0;

                // Подписываемся на изменение выбора. 
                // ВАЖНО: Если ты уже подписался на это событие в дизайнере (двойным кликом), 
                // то эту строку нужно удалить, иначе событие сработает дважды!
                // Для надежности лучше удалять подписку в коде, если она есть в дизайне.
                // Но пока оставим так, предполагая, что в дизайне подписки нет.
                comboBoxTables.SelectedIndexChanged += ComboBoxTables_SelectedIndexChanged;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось подключиться к базе данных.\n\nПроверьте:\n1) Файл БД лежит рядом с программой (.exe).\n2) Установлен ли Access Database Engine (разрядность x86/x64 должна совпадать с проектом).\n\nОшибка: {ex.Message}",
                    "Критическая ошибка подключения",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                this.Close(); // Закрываем форму, так как без БД работать нельзя
            }
        }

        private List<string> GetTableNames(string connStr)
        {
            var list = new List<string>();
            using var conn = new OleDbConnection(connStr);

            try
            {
                conn.Open();
                // Получаем схему таблиц
                string[] restrictions = new string[4] { null, null, null, "TABLE" };
                var schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, restrictions);

                if (schema != null)
                {
                    foreach (DataRow row in schema.Rows)
                    {
                        string name = row["TABLE_NAME"].ToString();
                        // Исключаем системные таблицы Access
                        if (!name.StartsWith("MSys") && !name.StartsWith("USys"))
                            list.Add(name);
                    }
                }
            }
            catch (Exception)
            {
                throw; // Пробрасываем ошибку выше, чтобы MainForm_Load её поймал
            }

            return list.OrderBy(x => x).ToList();
        }

        private void ComboBoxTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxTables.SelectedItem == null) return;

            string tableName = comboBoxTables.SelectedItem.ToString();

            // Защита от SQL-инъекций здесь не нужна, так как имя таблицы нельзя параметризовать в OleDb,
            // но мы фильтруем имена в GetTableNames, так что это безопасно.
            string sql = $"SELECT * FROM [{tableName}]";

            try
            {
                using var conn = new OleDbConnection(_connStr);
                using var cmd = new OleDbCommand(sql, conn);
                conn.Open();

                var adapter = new OleDbDataAdapter(cmd);
                var dt = new DataTable();
                adapter.Fill(dt);

                dataGridViewMain.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении таблицы '{tableName}':\n{ex.Message}",
                                "Ошибка данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Этот метод УДАЛИ, если он пустой. Он остался от дизайнера.
        // private void MainForm_Load_1(object sender, EventArgs e) { } 

        // Обработчик клика по ячейке (можно оставить пустым или использовать для редактирования)
        private void dataGridViewMain_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Пока ничего не делаем. Если нужно редактировать прямо в сетке, 
            // DataGridView сам это позволяет, если таблица не только для чтения.
        }
    }
}
