﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automation_tool
{
    class ExcelOperation
    {        
        Worksheet xlWorkSheet = null;
        Workbook xlWorkBook = null;
        string dataType = "DataPoint";
        public ExcelOperation(Workbook xlWorkBook, Worksheet xlWorkSheet)
        {
            this.xlWorkBook = xlWorkBook;
            this.xlWorkSheet = xlWorkSheet;
        }
        public void readCells()
        {
            if (xlWorkBook != null && xlWorkSheet!=null)
            {
                Excel.Range xlRange = xlWorkSheet.UsedRange;
                double rowsCount = xlRange.Rows.Count;
                double columnCount = xlRange.Columns.Count;
                bool found = false;
                for (int i = 2; i < rowsCount; i++)
                {
                    string value = xlWorkSheet.Cells[i, 2].Value2;
                    string[] codeSplit = value.Split(';');
                    for (int j = 0; j < codeSplit.Length; j++)
                    {
                        string line = codeSplit[j];
                        if (line.Contains(dataType) && !line.StartsWith("/") && !line.StartsWith("\n"))
                        {
                            char[] whitespace = new char[] { ' ', '\t' };
                            string[] fetchVariableName = line.Split(whitespace, StringSplitOptions.RemoveEmptyEntries);
                            if (fetchVariableName.Length > 0)
                            {
                                int index = Array.IndexOf(fetchVariableName, dataType);
                                if (index > -1)
                                {
                                    string variableTobeFound = fetchVariableName[index + 1];
                                    string toBeSearch = variableTobeFound+ "!=" + "null";
                                    string toBeSearch_1 = variableTobeFound + " " + "!=" + " " + "null";
                                    if (Array.Exists(codeSplit, element => element.Contains(toBeSearch)) || Array.Exists(codeSplit, element => element.Contains(toBeSearch_1)))
                                    {
                                        found = true;
                                    }
                                }
                            }
                        }
                    }
                }
            }            
        }
    }
}
