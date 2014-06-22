using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

 //Question 1: why when i create the excel file and i try to open, it always alerts me about possible file corruption?
//Question 2: how do you get ALL column values only for a specific format, for instance, only for integers
 
namespace Excel_Read
{
     class CreateExcel
    {

        public  CreateExcel() {

            /* */

        try {
            Console.WriteLine("enter the path where you have to create a Excel with Excel Name like C:\\Excel\\Test.xls");
            var path = Console.ReadLine();
            var app = new Excel.Application();
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range workSheetRange = null;
            Excel.Range workSheetDate = null;
            workbook = app.Workbooks.Add(1);
            worksheet = (Excel.Worksheet)workbook.Sheets[1];
            //worksheet.Cells[1, 1] = "dhaval";
            //worksheet.Cells[2, 1] = "Patel";

            worksheet.Cells[5, 5] = "Daily Expense Sheet";
            worksheet.Cells[7, 4] = "Date";
            worksheet.Cells[7, 5] = "Details";
            worksheet.Cells[7, 6] = "Amount";
            worksheet.Cells[8, 4] = "6/11/2014";
            worksheet.Cells[9, 4] = "6/12/2014";
            worksheet.Cells[10, 4] = "6/13/2014";
            worksheet.Cells[8, 5] = "Purchase from ABC Shop";
            worksheet.Cells[9, 5] = "petrol";
            worksheet.Cells[10, 5] = "Other";
            worksheet.Cells[8, 6] = 1250;
            worksheet.Cells[8,6].Interior.Color = Color.Red;
            worksheet.Cells[9, 6] = 500;
            worksheet.Cells[10, 6] = 800;
            worksheet.Cells[7, 9] = "TOTAL";
            worksheet.Cells[7, 10] = 2250;
            worksheet.Cells[5, 13].Interior.Color= Color.Red;
            worksheet.Cells[5, 14] = "Amount Increase 1000";


            workSheetRange = worksheet.Range["D5", "F5"];
            workSheetDate = worksheet.Range["D8", "D10"];
            workSheetDate.NumberFormat = "MM/DD/YYYY";
            //workSheetRange.MergeCells = true;
            workSheetRange.Interior.Color =Color.DodgerBlue ;
            // workSheetRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //workSheetRange.Interior.Color = System.Drawing.Color.Green;
            workbook.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbook.Close();
            }
        catch(Exception exp) {

            Console.WriteLine(exp.Message);
            Console.ReadLine();
            }
            }

    }
}
    
