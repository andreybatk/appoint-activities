using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelParser
{
    class ExcelHelper : IDisposable
    {
        public Excel.Application Excel;
        public Excel.Workbook Workbook;
        private string _filePath;

        public ExcelHelper()
        {
            Excel = new Excel.Application();
        }

        internal bool Open(string filePath)
        {
            Console.WriteLine("Открытие файла..");
            try
            {
                if (File.Exists(filePath))
                {
                    _filePath = filePath;
                    Workbook = Excel.Workbooks.Open(filePath);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Ошибка при открытии! Файла {_filePath} не существует!");
                    Console.ResetColor();
                    Console.ReadKey();
                    return false;
                }

                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }
        internal void Save()
        {
            if (string.IsNullOrEmpty(_filePath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Ошибка при сохранении! Файла {_filePath} не существует!");
                Console.ResetColor();
                Console.ReadKey();
                return;
            }
            else
            {
                Workbook.Save();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Файл {_filePath} успешно сохранен!");
                Console.ResetColor();
            }
        }
        internal bool Set(int column, int row, object data)
        {
            try
            {
                ((Excel.Worksheet)Excel.ActiveSheet).Cells[row, column] = data;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }
        public void Dispose()
        {
            try
            {
                Workbook.Close();
                Excel.Quit();
                Console.WriteLine($"Файл {_filePath} успешно закрыт!");
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}