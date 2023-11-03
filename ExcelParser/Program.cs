using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelParser
{
    // TESTS MULTITHREADING:
    // Примерное время завершения потоков: 13m. 34s.
    // Работа завершена. Время: 0h 13m. 25s.
    // TESTS SINGLETHREADING:
    // Примерное время завершения потоков: 25m. 12s.

    class Program
    {
        private static Stopwatch _watchTimer = new Stopwatch();
        private static readonly string _file = "MainTest.xlsx";
        private static int _numberRowsExecuted;
        private static int _processorCount;
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
            _processorCount = Environment.ProcessorCount;
        }
        private static void StartParser()
        {
            Console.WriteLine($"Работа с файлом {_file} началась!");

            Excel.Sheets _sheets;
            _foundColumns = new Dictionary<string, int>();
            bool showProgress = true;
            try
            {
                using (ExcelHelper helper = new ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, _file)))
                    {
                        _sheets = helper.Workbook.Sheets;

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
                            _watchTimer.Start();
                            var options = new ParallelOptions() { MaxDegreeOfParallelism = Environment.ProcessorCount - 1 };
                            Parallel.For(2, RowsCount, options, i =>
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
                                    _numberRowsExecuted += _numberRowsForProgress;
                                    Console.WriteLine($"Поток: #{Thread.CurrentThread.ManagedThreadId}. Прогресс: {_numberRowsExecuted}/{RowsCount} ");
                                }

                                if (showProgress && _numberRowsExecuted == 1000)
                                {
                                    TimeSpan tempTimeSpan = _watchTimer.Elapsed;
                                    double coeff = RowsCount / 1000;
                                    var ts = TimeSpan.FromSeconds(tempTimeSpan.TotalSeconds * coeff);
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Console.WriteLine($"Используемые потоки завершили 1000 строк за {tempTimeSpan.Minutes}m. {tempTimeSpan.Seconds}s.");
                                    Console.WriteLine($"Примерное время завершения потоков: {ts.Minutes}m. {ts.Seconds}s.");
                                    Console.ResetColor();
                                    showProgress = false;
                                }
                                #endregion
                            });
                            #region PROGRESS BAR ENDTIME
                            TimeSpan timeSpan = _watchTimer.Elapsed;
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine($"Работа завершена. Время: {timeSpan.Hours}h {timeSpan.Minutes}m. {timeSpan.Seconds}s.");
                            Console.ResetColor();
                            _watchTimer.Stop();
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
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"Найдено столбцов: {_foundColumns.Count} из необходимых {_requiredColumns.Count}");

            if(_foundColumns.Count != _requiredColumns.Count)
            {
                Console.WriteLine($"ВНИМАНИЕ! Несовпадение найденных и необходимых столбцов!\n" +
                    $"Чтобы продолжить работу нажмите \"Enter\"");
                Console.ReadKey();
            }

            Console.ResetColor();
        }
    }
}


