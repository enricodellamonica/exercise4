using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;


namespace Excel_Read {
    class ReadExcel {

        public ReadExcel() {
            
             Console.WriteLine("enter the path where you have to read a Excel with Excel Name like C:\\Test.xls");
            var Path = Console.ReadLine();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlPlatform.xlWindows, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    string temp = (string)(xlRange.Cells[i, j] as Excel.Range).Value;
                    Console.WriteLine(temp);
                }
            }
            xlWorkbook.Close();
            Console.ReadLine();
        }
        
        }
    }
