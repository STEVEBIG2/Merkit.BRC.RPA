using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.IO;
using System.Drawing;
using System.Xml.Linq;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;

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

        public static DataTable WorksheetToDataTable(Excel.Worksheet oSheet)
        {
            int totalRows = oSheet.UsedRange.Rows.Count;
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

    }

}