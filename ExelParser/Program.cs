using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExelParser
{
    internal class Program
    {
        private static string file = "Test.xlsx";
        static void Main(string[] args)
        {
            Console.WriteLine($"Работа с файлом {file}");
            Console.WriteLine("Чтобы запустить работу нажмити \"Enter\"");
            Console.ReadLine();
            ExcelHelper excel = new ExcelHelper(file);
            excel.StartParser();
        }
    }
}
