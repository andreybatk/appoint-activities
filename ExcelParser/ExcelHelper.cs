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
            SettingsCount = 0;
        }

        public int SettingsCount;
        public Excel.Application Excel;
        public Excel.Workbook Workbook;

        public bool Open(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    Console.WriteLine("Открытие файла..");
                    _filePath = filePath;
                    Workbook = Excel.Workbooks.Open(filePath);
                }
                else
                {
                    PrintMessage.PrintErrorMessage($"Ошибка при открытии! Файла {filePath} не существует!");
                    return false;
                }

                return true;
            }
            catch (Exception ex) { PrintMessage.PrintErrorMessage(ex.Message); }
            return false;
        }
        public void Save()
        {
            try
            {
                Workbook.Save();
                PrintMessage.PrintSuccessMessage($"Файл {_filePath} успешно сохранен!");
            }
            catch (Exception ex) { PrintMessage.PrintErrorMessage(ex.Message); }
        }
        public bool Set(int column, int row, object data)
        {
            try
            {
                ((Excel.Worksheet)Excel.ActiveSheet).Cells[row, column] = data;
                SettingsCount++;
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