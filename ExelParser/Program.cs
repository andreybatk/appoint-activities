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
            StartSetValue();
        }
        private static void StartSetValue()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Работа с файлом {file} началась.");
            Console.ResetColor();
            try
            {
                using (BO.ExcelHelper helper = new BO.ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, file)))
                    {
                        helper.Set(column: "A", row: 2, data: "Test");
                        //var val = helper.Get(column: "A", row: 6);
                        helper.Set(column: "B", row: 2, data: DateTime.Now);

                        helper.Save();
                    }
                }

                Console.Read();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
