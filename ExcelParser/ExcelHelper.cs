using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelParser
{
    public class ExcelHelper : IDisposable
    {
        private string _filePath;
        private SuccessMessage _successMessage = PrintMessage.PrintSuccessMessage;
        private ErrorMessage _errorMessage = PrintMessage.PrintErrorMessage;

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
                    _errorMessage?.Invoke($"Ошибка при открытии! Файла { _filePath} не существует!");
                    Console.ReadKey();
                    return false;
                }

                return true;
            }
            catch (Exception ex) { _errorMessage?.Invoke(ex.Message); }
            return false;
        }
        public void Save()
        {
            if (string.IsNullOrEmpty(_filePath))
            {
                _errorMessage?.Invoke($"Ошибка при сохранении! Файла {_filePath} не существует!");
                Console.ReadKey();
                return;
            }
            else
            {
                Workbook.Save();
                _successMessage?.Invoke($"Файл {_filePath} успешно сохранен!");
            }
        }
        public bool Set(int column, int row, object data)
        {
            try
            {
                ((Excel.Worksheet)Excel.ActiveSheet).Cells[row, column] = data;
                return true;
            }
            catch (Exception ex) { _errorMessage?.Invoke(ex.Message); }
            return false;
        }
        public void Dispose()
        {
            try
            {
                Workbook.Close();
                Excel.Quit();
                _successMessage?.Invoke($"Файл {_filePath} успешно закрыт!");
            }
            catch (Exception ex) { _errorMessage?.Invoke(ex.Message); }
        }
    }
}