﻿using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.IO;
using System.Drawing;
using System.Xml.Linq;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Merkit.RPA.PA.Framework
{
    public class ExcelManager
    {

        public Excel.Application ExcelApp = null;
        public Excel.Workbook ExcelWorkbook = null;
        public Excel.Worksheet ExcelSheet = null;

        /// <summary>
        /// Open Excel file
        /// </summary>
        /// <param name="excelFilename"></param>
        /// <param name="visible"></param>
        /// <returns></returns>
        public bool OpenExcel(string excelFilename, bool visible=true)
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

        /// <summary>
        /// Create Excel file
        /// </summary>
        /// <param name="excelFilename"></param>
        /// <param name="visible"></param>
        /// <returns></returns>
        public bool CreateExcel(string excelFilename, string sheetName, bool visible = true)
        {
            bool retValue = false;

            try
            {
                if (!File.Exists(excelFilename))
                {
                    ExcelApp = new Excel.Application();
                    ExcelApp.Visible = visible;
                    ExcelApp.DisplayAlerts = false;
                    ExcelWorkbook = ExcelApp.Workbooks.Add();
                    ExcelWorkbook.SaveAs(excelFilename);

                    ExcelSheet = (Worksheet)(ExcelWorkbook.Worksheets[1]);
                    ExcelSheet.Activate();
                    ExcelSheet.Name = sheetName;
                }

            }
            finally
            {
                retValue = (ExcelWorkbook != null);
            }

            return retValue;

        }

        /// <summary>
        /// Get WorksheetNames
        /// </summary>
        /// <returns></returns>
        public List<string> WorksheetNames()
        {
            List<string> sheetNames = new List<string>();

            foreach (Worksheet worksheet in ExcelWorkbook.Sheets)
            {
                sheetNames.Add(worksheet.Name);
            }

            return sheetNames;
        }

        /// <summary>
        /// Add New Sheet into first position of workbook and activate
        /// </summary>
        /// <param name="sheetName"></param>
        public void AddNewSheet(string worksheetName)
        {
            List<string> sheetNames = WorksheetNames();
            ExcelWorkbook.Worksheets.Add();

            foreach (Worksheet worksheet in ExcelWorkbook.Sheets)
            {
                if (!sheetNames.Contains(worksheet.Name))
                {
                    ExcelSheet = ExcelWorkbook.Sheets[worksheet.Name];
                    ExcelSheet.Activate();
                    ExcelSheet.Name = worksheetName;
                    break;
                }               
            }          
        }

        /// <summary>
        /// Add New Sheet if not exists into first position of workbook and activate
        /// </summary>
        /// <param name="sheetName"></param>
        public bool AddNewSheetIfNotExist(string worksheetName)
        {
            bool isNewRow = false;

            if(! WorksheetNames().Contains(worksheetName))
            {
                AddNewSheet(worksheetName);
                isNewRow = true;
            }

            SelectWorksheetByName(worksheetName);
            return isNewRow;
        }

        /// <summary>
        /// Delete sheet from workbook
        /// </summary>
        /// <param name="sheetName"></param>
        public void DeleteSheetIfExist(string worksheetName)
        {
            if(SelectWorksheetByName(worksheetName))
            {
                ExcelSheet.Delete();
            }
        }

        /// <summary>
        /// Select Worksheet By Name
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <returns></returns>
        public bool SelectWorksheetByName(string worksheetName)
        {
            bool retValue = true;

            try
            {
                 ExcelSheet = ExcelWorkbook.Sheets[worksheetName];
                 ExcelSheet.Activate();
            }
            catch
            {
                retValue = false;
            }

            return retValue;

        }

        /// <summary>
        /// Select Worksheet By Index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public bool SelectWorksheetByIndex(int index)
        {
            bool retValue = true;

            try
            {
                ExcelSheet = ExcelWorkbook.Sheets[index];
                ExcelSheet.Activate();
            }
            catch
            {
                retValue = false;
            }

            return retValue;

        }

        /// <summary>
        /// Select First Worksheet
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public bool SelectFirstWorksheetByIndex()
        {
            return SelectWorksheetByIndex(1);
        }

        /// <summary>
        /// Save Excel File
        /// </summary>
        public void SaveExcel()
        {
            if (ExcelWorkbook != null)
            {
                ExcelWorkbook.Save();
            }
        }

        /// <summary>
        /// Save Excel File
        /// </summary>
        public void SaveAndCloseExcel()
        {

            try
            {
                SaveExcel();
            }
            finally
            {
                if (ExcelWorkbook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkbook);
                    ExcelWorkbook = null;
                }

                if (ExcelApp != null)
                {
                    ExcelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
                    ExcelApp = null;
                }

            }

        }

        /// <summary>
        /// Close Excel file without save
        /// </summary>
        public void CloseExcelWithoutSave()
        {

            try
            {
                if (ExcelWorkbook != null)
                {
                    ExcelWorkbook.Save();
                    ExcelWorkbook.Close(false);
                }
            }
            finally
            {
                if (ExcelWorkbook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkbook);
                    ExcelWorkbook = null;
                }

                if (ExcelApp != null)
                {
                    ExcelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
                    ExcelApp = null;
                }

            }

        }

        /// <summary>
        /// Auto fit in current sheet
        /// </summary>
        public void AutoFit()
        {
            ExcelSheet.get_Range("A1").EntireRow.EntireColumn.AutoFit();
            ExcelSheet.get_Range("A1").EntireColumn.EntireRow.AutoFit();
        }

        /// <summary>
        /// Last Column index in current sheet
        /// </summary>
        /// <returns></returns>
        public int LastColumn()
        {
            return ExcelSheet.UsedRange.Columns.Count;
        }

        /// <summary>
        /// Last Row index in current sheet
        /// </summary>
        /// <returns></returns>
        public int LastRow()
        {
            return ExcelSheet.UsedRange.Rows.Count;
        }


        /// <summary>
        /// Set Cel lValue by cell name
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public void SetCellValue(string cell, object value)
        {
            ExcelSheet.get_Range(cell).Value = value;
            return;
        }

        /// <summary>
        /// Set Cell Value by row and col
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public void SetCellValue(int row, int col, object value)
        {
            ExcelSheet.Cells[row, col].Value = value;
            ExcelSheet.Cells[row, col].NumberFormat = "@";
            return;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellStart"></param>
        /// <param name="cellEnd"></param>
        /// <param name="value"></param>
        public void SetRangeValues(string cellStart, string cellEnd, object[] value)
        {
            ExcelSheet.get_Range(cellStart, cellEnd).Value = value; // mindenhova az első értéket írja
            return;
        }

        /// <summary>
        /// Read Cell Value by cell name
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public object ReadCellValue(string cell)
        {
            Range rng = ExcelSheet.get_Range(cell);
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

        /// <summary>
        /// Read Cell Value by row and col
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public string ReadCellValue(int row, int col)
        {
            string value = ExcelSheet.Cells[row, col].Value?.ToString();
            return value;
        }

        /// <summary>
        /// Get range for specified cell
        /// </summary>
        /// <param name="startCell"></param>
        /// <returns></returns>
        public Range GetCellRange(string startCell)
        {
            Range range = ExcelSheet.get_Range(startCell);
            return range;
        }

        /// <summary>
        /// Read entire row from cell
        /// </summary>
        /// <param name="startCell"></param>
        /// <returns></returns>
        public Range ReadEntireRow(string startCell)
        {
            Range range = GetCellRange(startCell).EntireRow;
            return range;
        }

        /// <summary>
        /// Read entire row from cell
        /// </summary>
        /// <param name="startCell"></param>
        /// <returns></returns>
        //public Range ReadEntireSheet(string startCell)
        //{
        //    Range range = ExcelSheet.get_Range(0, 0, 1, 1);
        //    return range;
        //}

        /// <summary>
        /// Set Cell Color by cell name
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="colorValue"></param>
        public void SetCellColor(string cell, Color colorValue)
        {
            ExcelSheet.get_Range(cell).Interior.Color = System.Drawing.ColorTranslator.ToOle(colorValue);
            return;
        }


        /// <summary>
        /// Set Cell Color by row and col
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="colorValue"></param>
        public void SetCellColor(int row, int col, Color colorValue)
        {
            ExcelSheet.Cells[row, col].Interior.Color = System.Drawing.ColorTranslator.ToOle(colorValue);
            return;
        }

        /// <summary>
        /// Set Range Color
        /// </summary>
        /// <param name="cellStart"></param>
        /// <param name="cellEnd"></param>
        /// <param name="colorValue"></param>
        public void SetRangeColor(string cellStart, string cellEnd, Color colorValue)
        {
            ExcelSheet.get_Range(cellStart, cellEnd).Interior.Color = System.Drawing.ColorTranslator.ToOle(colorValue);
            return;
        }

        /// <summary>
        /// Insert First Column
        /// </summary>
        /// <param name="value"></param>
        public void InsertFirstColumn(string value)
        {
            ExcelSheet.Cells[1, 1].EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            SetCellValue(1, 1, value);
            return;
        }

        /// <summary>
        /// Worksheet To DataTable
        /// </summary>
        /// <param name="oSheet"></param>
        /// <param name="onlyHeader"></param>
        /// <returns></returns>
        public DataTable WorksheetToDataTable(Excel.Worksheet oSheet, bool onlyHeader = false)
        {
            // only headers or all rows
            int totalRows = onlyHeader ? 1 : oSheet.UsedRange.Rows.Count;

            int totalCols = oSheet.UsedRange.Columns.Count;
            DataTable dt = new DataTable(oSheet.Name);
            DataRow dr = null;

            for (int i = 1; i <= totalRows; i++)
            {
                // no header row?
                if (i > 1)
                {
                    dr = dt.Rows.Add();
                }

                for (int j = 1; j <= totalCols; j++)
                {
                    // header row?
                    if (i == 1)
                    {
                        dt.Columns.Add(oSheet.Cells[i, j].Value.ToString());
                    }
                    else
                    {
                        dr[j - 1] = oSheet.Cells[i, j].Value?.ToString();
                    }
                }

            }

            return dt;
        }

        /// <summary>
        /// DataTable To Worksheet
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="oSheet"></param>
        /// <param name="onlyHeader"></param>
        public void DataTableToWorksheet(DataTable dt, Excel.Worksheet oSheet, bool onlyHeader = false)
        {
            int colRow = 1;
            int colNum = 0;

            // column headings
            foreach (DataColumn dc in dt.Columns)
            {
                colNum++;
                SetCellValue(colRow, colNum, dc.ColumnName);
                SetCellColor(colRow, colNum, Color.Gray);
            }

            // no only header?
            if(!onlyHeader)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    colRow++;

                    for (var i = 0; i < dt.Columns.Count; i++)
                    {
                        colNum++;
                        SetCellValue(colRow, i+1, dr[i]);
                    }
                }
            }

            AutoFit();

            return;
        }

        /// <summary>
        /// Get Excel Column Names By DataTable
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public Dictionary<string, string> GetExcelColumnNamesByDataTable(System.Data.DataTable dt)
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
        public string GetExcelColumnNameByColumnNumber(int columnNumber)
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
        public string GetDataRowValue(DataRow currentRow, string colName)
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