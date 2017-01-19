using System;
using System.Runtime.InteropServices;
using Excel=Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
     public class ExcelReader
    {
         public string Read(string FilePath)
         {
             Excel.Application xlApp;
             Excel.Workbook xlWorkBook;
             Excel.Worksheet xlWorkSheet;
             Excel.Range range;

             int rCnt, cCnt;
             int rw = 0;
             int cl = 0;

             xlApp = new Excel.Application();
             xlWorkBook = xlApp.Workbooks.Open(@FilePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
             xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

             range = xlWorkSheet.UsedRange;
             rw = range.Rows.Count;
             cl = range.Columns.Count;

             string[,] Data = new string[rw, cl];
             
             for(rCnt=1; rCnt <= rw; rCnt++)
             {
                 for(cCnt=1; cCnt <=cl; cCnt++)
                 {
                     Data[rCnt, cCnt] = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                 }
             }

             xlWorkBook.Close(true, null, null);
             xlApp.Quit();

             Marshal.ReleaseComObject(xlWorkSheet);
             Marshal.ReleaseComObject(xlWorkBook);
             Marshal.ReleaseComObject(xlApp);

             return Data[rw,cl];
         }
    }
}
