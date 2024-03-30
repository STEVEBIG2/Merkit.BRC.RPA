using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.IO;
using System.Drawing;
using System.Xml.Linq;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Merkit.RPA.PA.Framework
{
    public static class ExcelManager
    {

        public static Excel.Application ExcelApp = null;
        public static Excel.Workbook ExcelWorkbook = null;
        public static Excel.Worksheet ExcelSheet = null;

        public static bool OpenExcel(string excelFilename, bool visible=true)
        {
            bool retValue = false;

            try
            {
                if (File.Exists(excelFilename))
                {
                    ExcelApp = new Excel.Application();
                    ExcelApp.Visible = visible;
                    ExcelApp.DisplayAlerts = false;
                    ExcelWorkbook = ExcelApp.Workbooks.Open(excelFilename);

                    // activate first sheet
                    ExcelSheet = (Worksheet)(ExcelWorkbook.Worksheets[1]);
                    ExcelSheet.Activate();
                }

            }
            finally
            {
                retValue = (ExcelWorkbook != null);
            }

            return retValue;

        }

        public static void CloseExcel()
        {

            try
            {
                if (ExcelWorkbook != null)
                {
                    ExcelWorkbook.Save();
                    ExcelWorkbook.Close(true);
                }
            }
            finally
            {
                if (ExcelWorkbook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkbook);
                }


                if (ExcelApp != null)
                {
                    ExcelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
                }

            }

        }

        public static void SetCellValue(string cell, object value)
        {
            Range rng;
            rng = ExcelSheet.get_Range(cell);
            rng.Value = value;
            return;
        }

        public static void SetRangeValues(string cellStart, string cellEnd, object[] value)
        {
            Range rng;
            rng = ExcelSheet.get_Range(cellStart, cellEnd);
            rng.Value = value; // mindenhova az első értéket írja
            return;
        }

        public static object ReadCellValue(string cell)
        {
            Range rng;
            rng = ExcelSheet.get_Range(cell);

            return rng.Value;


            // if (cell.Value2 == null)
            // cell is blank
            //else if (cell.Value2 is string)
            // cell is text
            //else if (cell.Value is double)
            // cell is number;
            //else if (cell.Value2 is double)
            // cell is date
        }

        public static void SetCellColor(string cell, Color colorValue)
        {

            Range rng;
            rng = ExcelSheet.get_Range(cell);
            rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(colorValue);
            return;
        }

        public static void SetRangeColor(string cellStart, string cellEnd, Color colorValue)
        {

            Range rng;
            rng = ExcelSheet.get_Range(cellStart, cellEnd);
            rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(colorValue);
            return;
        }

        public static void InsertFirstColumn(string value)
        {
            Range rng;
            rng = ExcelSheet.get_Range("A1");
            rng.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            rng = ExcelSheet.get_Range("A1");
            rng.Value = value;
            return;
        }

        public static DataTable WorksheetToDataTable(Excel.Worksheet oSheet, bool onlyHeader = false)
        {
            int totalRows = onlyHeader ? 1 : oSheet.UsedRange.Rows.Count;
            int totalCols = oSheet.UsedRange.Columns.Count;
            DataTable dt = new DataTable(oSheet.Name);
            DataRow dr = null;
            for (int i = 1; i <= totalRows; i++)
            {
                if (i > 1) dr = dt.Rows.Add();
                for (int j = 1; j <= totalCols; j++)
                {
                    if (i == 1)
                        dt.Columns.Add(oSheet.Cells[i, j].Value.ToString());
                    else
                        dr[j - 1] = oSheet.Cells[i, j].Value?.ToString();
                }
            }
            return dt;
        }


        /// <summary>
        /// Get Excel Column Names By DataTable
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static Dictionary<string, string> GetExcelColumnNamesByDataTable(System.Data.DataTable dt)
        {
            Dictionary<string, string> dictExcelColumnNameToExcellCol = new Dictionary<string, string>();
            int colNum=0;

            foreach (DataColumn col in dt.Columns) 
            {
                colNum++;
                dictExcelColumnNameToExcellCol.Add(col.ColumnName, GetExcelColumnNameByColumnNumber(colNum));
            }

            return dictExcelColumnNameToExcellCol;
        }

        /// <summary>
        /// Get ExcelColumnName by columnNumber
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        public static string GetExcelColumnNameByColumnNumber(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        /// <summary>
        /// Get DataRow Value
        /// </summary>
        /// <param name="currentRow"></param>
        /// <param name="colName"></param>
        /// <returns></returns>
        public static string GetDataRowValue(DataRow currentRow, string colName)
        {
            string value = "";
            
            if(currentRow[colName] != null)
            {
                value = currentRow[colName].ToString(); 
            }

            return value;
        }

    }

}