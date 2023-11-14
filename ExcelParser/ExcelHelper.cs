using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelParser
{
    internal class ExcelHelper : IDisposable
    {
        private string _filePath;

        public ExcelHelper()
        {
            Excel = new Excel.Application();
        }

        public Excel.Application Excel;
        public Excel.Workbook Workbook;

        public bool Open(string filePath)
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
                    PrintMessage.PrintErrorMessage($"Ошибка при открытии! Файла { _filePath} не существует!");
                    Console.ReadKey();
                    return false;
                }

                return true;
            }
            catch (Exception ex) { PrintMessage.PrintErrorMessage(ex.Message); }
            return false;
        }
        public void Save()
        {
            if (string.IsNullOrEmpty(_filePath))
            {
                PrintMessage.PrintErrorMessage($"Ошибка при сохранении! Файла {_filePath} не существует!");
                Console.ReadKey();
                return;
            }
            else
            {
                Workbook.Save();
                PrintMessage.PrintSuccessMessage($"Файл {_filePath} успешно сохранен!");
            }
        }
        public bool Set(int column, int row, object data)
        {
            try
            {
                ((Excel.Worksheet)Excel.ActiveSheet).Cells[row, column] = data;
                return true;
            }
            catch (Exception ex) { PrintMessage.PrintErrorMessage(ex.Message); }
            return false;
        }
        public void Dispose()
        {
            try
            {
                Workbook.Close();
                Excel.Quit();
                PrintMessage.PrintSuccessMessage($"Файл {_filePath} успешно закрыт!");
            }
            catch (Exception ex) { PrintMessage.PrintErrorMessage(ex.Message); }
        }
    }
}