using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automation_tool
{
    class ReadExcel
    {
        public string path = string.Empty;
        public string name = string.Empty;
        string value = "";
        Excel.Application xlApp = new Excel.Application();
        public ReadExcel(string path, string name)
        {
            this.path = path;
            this.name = name;
        }       
        public void openExcel()
        {
            if (path != "")
            {
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(path);
                xlWorkBook.Unprotect();
                Worksheet sheet = xlWorkBook.Worksheets["CustomFunctions"];
                string sheetName = sheet.Name;
                sheet.Activate();
                ExcelOperation excelOpt = new ExcelOperation(xlWorkBook, sheet);
                excelOpt.readCells();
            }            
        }
        public void closeWorkbook()
        {
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(path);
            xlWorkBook.Close();
        }
    }
}
