using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToDb
{
    class StudentData
    {
        List<object> columns = new List<object>();
        List<List<object>> rows = new List<List<object>>();

        public void HeaderRow(Excel.Range xlRange, int rowCount, int colCount)
        {

        }
    }
}
