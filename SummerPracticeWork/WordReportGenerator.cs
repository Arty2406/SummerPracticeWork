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
            Word.Table wordTable = null;
            bool wordVisible = false;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;

                doc = wordApp.Documents.Add();
                Word.Selection selection = wordApp.Selection;

                // заголовок
                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                selection.Font.Size = 16;
                selection.Font.Bold = 1;
                selection.TypeText($"Отчёт по таблице «{tableName}»");
                selection.TypeParagraph();

                // информация
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

                // получение вилимих строк и столбцов
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
                    return new ReportResult { Success = false, ErrorMessage = "Таблица пуста. Нечего включать в отчёт." };
                }

                Word.Range tableRange = doc.Content;
                tableRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                // создание таблицы в Word
                wordTable = doc.Tables.Add(tableRange, rowCount + 1, colCount);
                wordTable.Borders.Enable = 1;

                // заголовки
                for (int i = 0; i < colCount; i++)
                {
                    Word.Cell cell = wordTable.Cell(1, i + 1);
                    cell.Range.Text = visibleCols[i].HeaderText;
                    cell.Range.Font.Bold = 1;
                    cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray15;
                }

                // данные
                int currentRow = 2;
                foreach (DataGridViewRow dgvRow in visibleRows)
                {
                    for (int i = 0; i < colCount; i++)
                    {
                        Word.Cell cell = wordTable.Cell(currentRow, i + 1);
                        object cellValue = dgvRow.Cells[visibleCols[i].Name].Value;

                        if (cellValue is DateTime dt)
                            cell.Range.Text = dt.ToString("dd.MM.yyyy");
                        else if (cellValue != null && cellValue != DBNull.Value)
                            cell.Range.Text = cellValue.ToString();
                        else
                            cell.Range.Text = "";
                    }
                    currentRow++;
                }

                wordTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                
                if (File.Exists(fullPath))
                {
                    try { File.Delete(fullPath); } catch { }
                }

                doc.SaveAs2(fullPath, Word.WdSaveFormat.wdFormatXMLDocument);

                wordApp.Visible = true;
                wordApp.Activate();
                wordVisible = true;

                return new ReportResult { Success = true, FilePath = fullPath };
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                return new ReportResult
                {
                    Success = false,
                    IsWordNotInstalled = true,
                    ErrorMessage = $"Не удалось создать отчёт в Word.\n\nВозможно, Microsoft Word не установлен.\n\nОшибка: {ex.Message}"
                };
            }
            catch (Exception ex)
            {
                return new ReportResult { Success = false, ErrorMessage = ex.Message };
            }
            finally
            {
                if (wordVisible)
                {
                    if (wordTable != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordTable);
                        wordTable = null;
                    }
                    if (doc != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                        doc = null;
                    }
                    if (wordApp != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                        wordApp = null;
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                else
                {
                    try
                    {
                        if (doc != null)
                        {
                            doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                            doc = null;
                        }
                    }
                    catch { }

                    try
                    {
                        if (wordApp != null)
                        {
                            wordApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                            wordApp = null;
                        }
                    }
                    catch { }
                }
            }
        }
    }
}