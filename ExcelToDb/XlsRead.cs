using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToDb
{
    class XlsRead
    {
        public void xlsRead(Excel.Range xlRange, int rowCount, int colCount)
        {

            for (int i = 4; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    if (j == 1)
                        Console.WriteLine("\r\n");

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                    }
                }
            }
        }
    }
}
