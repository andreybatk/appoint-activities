using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExelParser
{
    class ExcelHelper : IDisposable
    {
        public Excel.Application _excel;
        public Excel.Workbook _workbook;
        private string _filePath;

        public ExcelHelper()
        {
            _excel = new Excel.Application();
        }

        internal bool Open(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    _filePath = filePath;
                    _workbook = _excel.Workbooks.Open(filePath);
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
        internal void Save()
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
                Console.WriteLine("Прогресс: завершено");
                Console.WriteLine($"Файл {_filePath} успешно сохранен!");
                Console.ResetColor();
            }
        }
        internal bool Set(int column, int row, object data)
        {
            try
            {
                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column] = data;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }
        public void Dispose()
        {
            try
            {
                _workbook.Close();
                _excel.Quit();
                Console.WriteLine($"Файл {_filePath} успешно закрыт!");
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}