using System;
using System.Collections.Generic;
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
        private static readonly int[] _requiredColumns = { 4, 10, 18, 19, 26, 27, 28, 30, 32, 44, 49, 51, 151 };
        private static readonly int _numberRowsForProgress = 100;
        static void Main(string[] args)
        {
            Console.WriteLine($"Работа с файлом {_file}");
            Console.WriteLine("Чтобы запустить работу нажмити \"Enter\"");
            Console.ReadLine();

            StartParser();
        }
        private static void StartParser()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Работа с файлом {_file} началась!");
            Console.ResetColor();

            Excel.Sheets _sheets;
            try
            {
                using (ExcelHelper helper = new ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "Test.xlsx")))
                    {
                        _sheets = helper._workbook.Sheets;

                        foreach (Excel.Worksheet worksheet in _sheets)
                        {  
                            Excel.Range UsedRange = worksheet.UsedRange; // Получаем диапазон используемых на странице ячеек        
                            Excel.Range urRows = UsedRange.Rows; // Получаем строки в используемом диапазоне
                            Excel.Range urColums = UsedRange.Columns; // Получаем столбцы в используемом диапазоне

                            int RowsCount = urRows.Count;
                            int ColumnsCount = urColums.Count;
                            for (int i = 1; i <= RowsCount; i++)
                            {
                                foreach (var column in _requiredColumns)
                                {
                                    Excel.Range CellRange = UsedRange.Cells[i, column];

                                    string cellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                                        (CellRange as Excel.Range).Value2.ToString();

                                    if (cellText != null)
                                    {
                                        //helper.Set(11, 2, data: "MYTEST2");
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
    }
}


