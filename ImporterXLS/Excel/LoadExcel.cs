using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ImporterXLS
{
    public class LoadExcel
    {      

        public List<Dictionary<string,object>> OpenExcelFile(string fileName, int SheetPage = 1, int HeaderRow = 1)
        {
            var excelApp = new Excel.Application();
            var workbook = excelApp.Workbooks.Open(fileName);
            var sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[SheetPage];
            var totalRows = sheet.UsedRange.Rows.Count;
            var totalColumns = sheet.UsedRange.Columns.Count;
            List<Dictionary<string, object>> dictList = new List<Dictionary<string, object>>();
            
            Excel.Range xlRange = sheet.UsedRange;


            for (int row = 1 + HeaderRow; row <= totalRows; row++)
            {
                Dictionary<string, object> dict = new Dictionary<string, object>();

                for (int column = 1; column <= totalColumns; column++)
                {                      
                    if (xlRange.Cells[row, column] != null)
                    {
                        var key = xlRange.Cells[HeaderRow, column].Value2;
                        var value = xlRange.Cells[row, column].Value2;
                        dict.Add(key, value);
                    }
                   
                }                
               dictList.Add(dict);
            }


            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(sheet);

          
            workbook.Close();
            Marshal.ReleaseComObject(sheet);

           
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            return dictList;
       
        }
   

    }
}
