using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelToDb
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(Constants.Constants.Connection_string);
            Excel.Worksheet xlWorksheet = xlWorkBook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            XlsRead read = new XlsRead();
            read.xlsRead(xlRange, rowCount, colCount);


            //Validation email = new Validation();
            //bool isvalid = Validation.EmailIsValid("ngo@test.co.in.st");
            //Console.WriteLine(isvalid); 

            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkBook.Close();
            Marshal.ReleaseComObject(xlWorkBook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            Console.ReadKey();
         }
     }
}
