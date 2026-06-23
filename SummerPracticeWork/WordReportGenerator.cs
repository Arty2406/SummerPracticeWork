using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace SummerPractice
{
    public class ReportResult
    {
        public bool Success { get; set; }
        public string FilePath { get; set; }
        public string ErrorMessage { get; set; }
        public bool IsWordNotInstalled { get; set; }
    }

    public static class WordReportGenerator
    {
        public static ReportResult CreateReport(DataGridView dgv, string tableName, string userName)
        {
            if (dgv == null || dgv.DataSource == null)
                return new ReportResult { Success = false, ErrorMessage = "Нет данных для создания отчёта." };

            DateTime now = DateTime.Now;
            string dateNow = now.ToString("dd.MM.yyyy");
            string timeNow = now.ToString("HH:mm:ss");

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

                // Заголовок
                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                selection.Font.Size = 16;
                selection.Font.Bold = 1;
                selection.TypeText($"Отчёт по таблице «{tableName}»");
                selection.TypeParagraph();

                // Информация
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

                // Получаем видимые строки и столбцы
                var visibleRows = dgv.Rows.Cast<DataGridViewRow>()
                    .Where(r => !r.IsNewRow && r.Visible)
                    .ToList();
                var visibleCols = dgv.Columns.Cast<DataGridViewColumn>()
                    .Where(c => c.Visible)
                    .ToList();

                int rowCount = visibleRows.Count;
                int colCount = visibleCols.Count;

                if (rowCount == 0 || colCount == 0)
                {
                    doc.Close(false);
                    wordApp.Quit(false);
                    wordVisible = true;
                    return new ReportResult { Success = false, ErrorMessage = "Таблица пуста. Нечего включать в отчёт." };
                }

                // Создаём таблицу в Word
                Word.Table wordTable = doc.Tables.Add(selection.Range, rowCount + 1, colCount);
                wordTable.Borders.Enable = 1;

                // Заголовки
                for (int c = 0; c < colCount; c++)
                {
                    Word.Cell cell = wordTable.Cell(1, c + 1);
                    cell.Range.Text = visibleCols[c].HeaderText;
                    cell.Range.Font.Bold = 1;
                    cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray15;
                }

                // Данные
                int currentRow = 2;
                foreach (DataGridViewRow dgvRow in visibleRows)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        Word.Cell cell = wordTable.Cell(currentRow, c + 1);
                        object cellValue = dgvRow.Cells[visibleCols[c].Index].Value;

                        // Форматируем даты без времени
                        if (cellValue is DateTime dt)
                            cell.Range.Text = dt.ToString("dd.MM.yyyy");
                        else
                            cell.Range.Text = cellValue?.ToString() ?? "";
                    }
                    currentRow++;
                }

                wordTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                doc.SaveAs2(fullPath, Word.WdSaveFormat.wdFormatXMLDocument);

                wordApp.Visible = true;
                wordApp.Activate();
                wordVisible = true;

                return new ReportResult { Success = true, FilePath = fullPath };
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                try { doc?.Close(false); } catch { }
                if (wordApp != null && !wordVisible)
                    try { wordApp.Quit(false); } catch { }

                return new ReportResult
                {
                    Success = false,
                    IsWordNotInstalled = true,
                    ErrorMessage = $"Не удалось создать отчёт в Word.\n\nВозможно, Microsoft Word не установлен.\n\nОшибка: {ex.Message}"
                };
            }
            catch (Exception ex)
            {
                try { doc?.Close(false); } catch { }
                if (wordApp != null && !wordVisible)
                    try { wordApp.Quit(false); } catch { }

                return new ReportResult { Success = false, ErrorMessage = ex.Message };
            }
            finally
            {
                if (!wordVisible)
                {
                    if (doc != null)
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(doc); } catch { }
                    if (wordApp != null)
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp); } catch { }
                }
            }
        }
    }
}