using System;
using System.Runtime.InteropServices;
using Excel=Microsoft.Office.Interop.Excel;

namespace ExcelMod
{
     public class ExcelMod
    {
         public string[,] Read(string filePath)
         {
             var xlApp = new Excel.Application();
             var xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
             var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item[1];

             var range = xlWorkSheet.UsedRange;
             var rowCount = range.Rows.Count;
             var columnCount = range.Columns.Count;

             var data = new string[rowCount, columnCount];
             
             for (var rowNumber = 1; rowNumber <= rowCount; rowNumber++)
             {
                 for (var columnNumber = 1; columnNumber <= columnCount; columnNumber++)
                 {
                     var cell = range.Cells[rowNumber, columnNumber] as Excel.Range;
                     if (cell != null)
                     {
                        var cellValue = cell.Value2 != null ? cell.Value2 : "";
                        data[rowNumber, columnNumber] = cellValue.ToString();
                     }
                 }
             }

             xlWorkBook.Close(true);
             xlApp.Quit();

             Marshal.ReleaseComObject(xlWorkSheet);
             Marshal.ReleaseComObject(xlWorkBook);
             Marshal.ReleaseComObject(xlApp);

             return data;
         }
          
          public void Display(var result)
          {
               var rows =result.GetLength(0);
               var columns=result.GetLength(1);
               
               for(var rowNumber=1; rowNumber <= rows; rowNumber++)
               {
                    for(var columnNumber=1; columnNumber <= columns; columnNumber++)
                    {
                         Console.Write("{0}\t", result[rowNumber, columnNumber]);
                    }
                    Console.WriteLine();
               }
          }
    }
}
