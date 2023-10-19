using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExelParser
{
    class ExcelHelper
    {
        private Excel.Application _excel;
        private Excel.Workbook _workbook;
        private Excel.Sheets _sheets;
        private string _filePath;
        private int[] requiredColumns = { 4, 10, 18, 19, 26, 27, 28, 30, 32, 44, 49, 51, 151 };

        public ExcelHelper(string filePath)
        {
            this._filePath = filePath;
            this._excel = new Excel.Application();
        }
        public void StartParser()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Работа с файлом {_filePath} началась!");
            Console.ResetColor();
            try
            {
                EHOpen();

                // Получение всех страниц докуента
                _sheets = _workbook.Sheets;

                foreach (Excel.Worksheet worksheet in _sheets)
                {
                    // Получаем диапазон используемых на странице ячеек
                    Excel.Range UsedRange = worksheet.UsedRange;
                    // Получаем строки в используемом диапазоне
                    Excel.Range urRows = UsedRange.Rows;
                    // Получаем столбцы в используемом диапазоне
                    Excel.Range urColums = UsedRange.Columns;

                    // Количества строк и столбцов
                    int RowsCount = urRows.Count;
                    int ColumnsCount = urColums.Count;
                    for (int i = 1; i <= RowsCount; i++)
                    {
                        foreach (var column in requiredColumns)
                        {
                            Excel.Range CellRange = UsedRange.Cells[i, column];
                            // Получение текста ячейки
                            string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                                (CellRange as Excel.Range).Value2.ToString();

                            if (CellText != null)
                            {
                                EHSet(1, 1, "qqqqq");
                            }
                        }
                        if (i % 2 == 0)
                        {
                            //Console.Clear();
                            Console.WriteLine($"Прогресс: {i}/{RowsCount}");
                        }
                    }
                    // Очистка неуправляемых ресурсов на каждой итерации
                    if (urRows != null) Marshal.ReleaseComObject(urRows);
                    if (urColums != null) Marshal.ReleaseComObject(urColums);
                    if (UsedRange != null) Marshal.ReleaseComObject(UsedRange);
                    if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                /* Очистка оставшихся неуправляемых ресурсов */
                if (_sheets != null)
                {
                    Marshal.ReleaseComObject(_sheets);
                }
                if (_workbook != null)
                {
                    EHSave();
                    _workbook.Close(true);
                    Marshal.ReleaseComObject(_workbook);
                    _workbook = null;
                }
                //if (workbooks != null)
                //{
                //    workbooks.Close();
                //    Marshal.ReleaseComObject(workbooks);
                //    workbooks = null;
                //}
                if (_excel != null)
                {
                    _excel.Quit();
                    Marshal.ReleaseComObject(_excel);
                    _excel = null;
                }
            }
        }
        internal bool EHOpen()
        {
            try
            {
                if (File.Exists(_filePath))
                {
                    _workbook = _excel.Workbooks.Open(_filePath);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Ошибка при открытии! Файла {_filePath} не существует!");
                    Console.ResetColor();
                    return false;
                }

                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

        internal void EHSave()
        {
            if (string.IsNullOrEmpty(_filePath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Ошибка при сохранении! Файла {_filePath} не существует!");
                Console.ResetColor();
                return;
            }
            else
            {
                _workbook.Save();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Файл {_filePath} успешно сохранен!");
                Console.ResetColor();
            }
        }

        internal bool EHSet(int column, int row, object data)
        {
            try
            {
                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column] = data;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

        internal object Get(string column, int row)
        {
            try
            {
                return ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column].Value2;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return null;
        }
    }
}