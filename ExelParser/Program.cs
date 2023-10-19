using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExelParser
{
    class Program
    {
        private static readonly string _file = "Test.xlsx";
        private static readonly int _numberRowsForProgress = 100;
        /// <summary>
        /// Найденные колонки в файле, которые соответствуют необходимым
        /// </summary>
        private static Dictionary<string, int> _foundColumns;
        /// <summary>
        /// Необходимые колонки
        /// </summary>
        private static List<string> _requiredColumns = new List<string> {"ETA", "TOTA" };
        //private static List<int> _currentNumbersColumns = new List<int> { 4, 10, 18, 19, 26, 27, 28, 30, 32, 44, 49, 51, 151 };

        static void Main(string[] args)
        {
            Console.WriteLine($"Работа с файлом {_file}");
            Console.WriteLine("Чтобы запустить работу нажмите \"Enter\"");
            Console.ReadLine();

            StartParser();
        }
        private static void StartParser()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Работа с файлом {_file} началась!");
            Console.ResetColor();

            Excel.Sheets _sheets;
            _foundColumns = new Dictionary<string, int>();

            try
            {
                using (ExcelHelper helper = new ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, _file)))
                    {
                        _sheets = helper._workbook.Sheets;

                        foreach (Excel.Worksheet worksheet in _sheets)
                        {  
                            Excel.Range UsedRange = worksheet.UsedRange; // Получаем диапазон используемых на странице ячеек        
                            Excel.Range urRows = UsedRange.Rows; // Получаем строки в используемом диапазоне
                            Excel.Range urColums = UsedRange.Columns; // Получаем столбцы в используемом диапазоне

                            int RowsCount = urRows.Count;
                            int ColumnsCount = urColums.Count;

                            //Получение нужных столбцов
                            for (int i = 1; i < 2; i++) 
                            {
                                for (int j = 1; j <= ColumnsCount; j++)
                                {
                                    Excel.Range CellRange = UsedRange.Cells[i, j];
                                    string cellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                                        (CellRange as Excel.Range).Value2.ToString();

                                    if(_requiredColumns.Contains(cellText)) { _foundColumns.Add(cellText, j); }
                                }
                            }

                            StatusColumns();
                            Console.WriteLine($"Началась обработка! Необходимо обработать: {RowsCount} строк");

                            for (int i = 2; i <= RowsCount; i++)
                            {
                                foreach (var column in _foundColumns)
                                {
                                    Excel.Range CellRange = UsedRange.Cells[i, column.Value];
                                    string cellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                                        (CellRange as Excel.Range).Value2.ToString();

                                    if (cellText != null)
                                    {
                                        if(column.Key == _requiredColumns[0]) // если столбец под индексом 0, тоесть "ETA"
                                        {
                                            //helper.Set(i, 2, data: "MYTEST2"); //устанавливаем значение в нужную строку и колонку (строка автоматический берется и i)
                                        }
                                    }
                                }
                                if (i % _numberRowsForProgress == 0)
                                {
                                    Console.Clear();
                                    Console.WriteLine($"Прогресс: {i}/{RowsCount}");
                                }
                            }
                            #region 
                            // Очистка неуправляемых ресурсов на каждой итерации
                            //if (urRows != null) Marshal.ReleaseComObject(urRows);
                            //if (urColums != null) Marshal.ReleaseComObject(urColums);
                            //if (UsedRange != null) Marshal.ReleaseComObject(UsedRange);
                            //if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                            #endregion
                        }
                        helper.Save();
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
        private static void StatusColumns()
        {
            string foundColumns = "";
            foreach (var column in _foundColumns)
            {
                foundColumns += $"{column.Key}\t";
            }
            Console.WriteLine($"Найдено столбцов: {_foundColumns.Count} из {_requiredColumns.Count}");
            Console.WriteLine(foundColumns);

            if(_foundColumns.Count != _requiredColumns.Count)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"ВНИМАНИЕ! Несовпадение найденных и необходимых столбцов!\n" +
                    $"Чтобы продолжить работу нажмите \"Enter\"");
                Console.ResetColor();
                Console.ReadKey();
            }
        }
    }
}


