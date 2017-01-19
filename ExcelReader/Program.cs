using System;
using Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter your file path:\n ");
            var filePath = Console.ReadLine();

            Console.WriteLine("Your file path is: '{0}' ", filePath);
            Console.Read();

            var trial = new ExcelReader();
            var result = trial.Read(filePath);
        }
    }
}
