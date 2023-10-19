using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Tests
{
    class Program
    {
        //private Excel.Sheets _sheets;
        static void Main(string[] args)
        {
            Excel.Sheets _sheets;
            int[] requiredColumns = { 4, 10, 18, 19, 26, 27, 28, 30, 32, 44, 49, 51, 151 };
            try
            {
                using (Class1 helper = new Class1())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "Test.xlsx")))
                    {
                        _sheets = helper._workbook.Sheets;

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
                                        //UsedRange.Cells[1, 1] = "qqq";
                                        helper.Set(10, 1, data: "MYTEST");
                                    }
                                }
                                if (i % 2 == 0)
                                {
                                    //Console.Clear();
                                    Console.WriteLine($"Прогресс: {i}/{RowsCount}");
                                }
                            }
                            // Очистка неуправляемых ресурсов на каждой итерации
                            //if (urRows != null) Marshal.ReleaseComObject(urRows);
                            //if (urColums != null) Marshal.ReleaseComObject(urColums);
                            //if (UsedRange != null) Marshal.ReleaseComObject(UsedRange);
                            //if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                        }
                        helper.Set(column: 5, row: 1, data: "lksadklsajdkl");
                        //var val = helper.Get(column: "A", row: 6);
                        helper.Set(column: 5, row: 1, data: DateTime.Now);

                        helper.Save();
                    }
                }

                Console.Read();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
