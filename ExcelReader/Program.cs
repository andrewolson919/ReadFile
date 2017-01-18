using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string FilePath;
            int rCount, cCount, row, col;

            Console.WriteLine("Enter your file path:\n ");
            FilePath = Console.ReadLine();

            Console.WriteLine("Your file path is: '{0}' ", FilePath);
            Console.Read();

            xlApp = new Excel.Application();

        }
    }
}
