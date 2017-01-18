using System;
using Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            //Excel.Application xlApp;
            //Excel.Workbook xlWorkBook;
            //Excel.Worksheet xlWorkSheet;
            //Excel.Range range;

            //string FilePath;
            //int rCount, cCount, row, col;

            Console.WriteLine("Enter your file path:\n ");
            var filePath = Console.ReadLine();

            Console.WriteLine("Your file path is: '{0}' ", filePath);
            Console.Read();

            //xlApp = new Excel.Application();
        }
    }
}
