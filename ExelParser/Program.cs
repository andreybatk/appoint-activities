using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExelParser
{
    class Program
    {
        private static Stopwatch _watchTimer = new Stopwatch();
        private static readonly string _file = "Test.xlsx";
        /// <summary>
        /// показывать прогресс каждые _numberRowsForProgress строк
        /// </summary>
        private static readonly int _numberRowsForProgress = 100;
        /// <summary>
        /// Найденные колонки в файле, которые соответствуют необходимым
        /// </summary>
        private static Dictionary<string, int> _foundColumns;
        /// <summary>
        /// Необходимые колонки
        /// </summary>
        private static List<string> _requiredColumns = new List<string>
        {   "Катег. защитн.",
            "Катег. Земель",
            "Мер. 1", "% выборки",
            "Группа А",
            "Класс А",
            "Преобл. Порода",
            "Бонитет",
            "ТЛУ",
            "A1",
            "Полнота1",
            "Запас1",
            "Густота подр."
        };
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
            Console.WriteLine($"Работа с файлом {_file} началась!");

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
                            for (int j = 1; j <= ColumnsCount; j++)
                            {
                                Excel.Range CellRange = UsedRange.Cells[1, j]; //row: 1 column: j
                                string cellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                                    (CellRange as Excel.Range).Value2.ToString();

                                if(_requiredColumns.Contains(cellText)) { _foundColumns.Add(cellText, j); }
                            }

                            StatusColumns();
                            Console.WriteLine($"Началась обработка! Необходимо обработать: {RowsCount} строк");

                            int startRowsCount1, endRowCount1;
                            startRowsCount1 = 2; endRowCount1 = RowsCount / 4;

                            int startRowsCount2, endRowCount2;
                            startRowsCount2 = endRowCount1 + 1; endRowCount2 = endRowCount1 + endRowCount1;

                            int startRowsCount3, endRowCount3;
                            startRowsCount3 = endRowCount2 + 1; endRowCount3 = endRowCount2 + endRowCount1;

                            int startRowsCount4, endRowCount4;
                            startRowsCount4 = endRowCount3 + 1; endRowCount4 = RowsCount;

                            Parallel.Invoke(
                                () =>
                                {
                                    ParseElement(startRowsCount1, endRowCount1, UsedRange, 1);
                                },
                                () =>
                                {
                                    ParseElement(startRowsCount2, endRowCount2, UsedRange, 2);
                                },
                                () =>
                                {
                                    ParseElement(startRowsCount3, endRowCount3, UsedRange, 3);
                                },
                                () =>
                                {
                                    ParseElement(startRowsCount4, endRowCount4, UsedRange, 4);
                                });
                        }
                        helper.Save();
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
        private static void ParseElement(int StartRowsCount, int EndRowsCount, Excel.Range UsedRange, int thread)
        {
            bool isCheckTimeLeft = false;

            _watchTimer.Start();
            for (int i = StartRowsCount; i <= EndRowsCount; i++)
            {
                foreach (var column in _foundColumns)
                {
                    Excel.Range CellRange = UsedRange.Cells[i, column.Value];
                    string cellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                        (CellRange as Excel.Range).Value2.ToString();

                    if (cellText != null)
                    {
                        if (column.Key == _requiredColumns[0]) // если столбец под индексом 0
                        {
                            //helper.Set(i, 2, data: "MYTEST2"); //устанавливаем значение в нужную строку и колонку (строка автоматический берется и i)
                        }
                    }
                }

                #region PROGRESS BAR
                if (i % _numberRowsForProgress == 0)
                {
                    Console.WriteLine($"Поток: #{thread}. Прогресс: {i}/{EndRowsCount} ");
                }
                if (!isCheckTimeLeft && thread == 1)
                {
                    if (i == 1000)
                    {
                        TimeSpan tempTimeSpan = _watchTimer.Elapsed;
                        double coeff = EndRowsCount / 1000;
                        var ts = TimeSpan.FromSeconds(tempTimeSpan.TotalSeconds * coeff);
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"Все потоки завершили 1000 строк за {tempTimeSpan.Minutes}m. {tempTimeSpan.Seconds}s.");
                        Console.WriteLine($"Примерное время завершения потоков: {ts.Minutes}m. {ts.Seconds}s.");
                        Console.ResetColor();
                        isCheckTimeLeft = true;
                    }
                }
                #endregion
            }
            #region PROGRESS BAR ENDTIME
            TimeSpan timeSpan = _watchTimer.Elapsed;
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Поток #{thread} завершил работу. Время: {timeSpan.Hours}h {timeSpan.Minutes}m. {timeSpan.Seconds}s.");
            Console.ResetColor();
            _watchTimer.Stop();
            #endregion
        }
        private static void StatusColumns()
        {
            Console.WriteLine($"Найдено столбцов: {_foundColumns.Count} из необходимых {_requiredColumns.Count}");

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


