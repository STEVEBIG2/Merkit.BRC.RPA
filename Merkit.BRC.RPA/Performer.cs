using Merkit.RPA.PA.Framework;
using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace Merkit.BRC.RPA
{
    public static class Performer
    {

        /// <summary>
        /// Create Result Excels
        /// </summary>
        /// <param name="destRootFolder"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static bool CreateResultExcels(string destRootFolder)
        {
            bool isOk = true;
            bool isConnected = false;
            string sqlQuery = "";
            int excelFileId = 30;
            string exceldestFileName = Path.Combine(@"c:\Munka", $"Munka0530_{excelFileId}.xlsx");
            //string sysAdminName = Config.NotifyEmail;
            MSSQLManager sqlManager = new MSSQLManager();

            List<string> viewColums = new List<string>();
            viewColums.Add("[Státusz]");
            viewColums.Add("[Ügyszám]");

            foreach (ExcelCol excelCol in ExcelValidator.excelHeaders.Where(x => !String.IsNullOrEmpty(x.SQLColName)))
            {
                viewColums.Add($"[{excelCol.ExcelColName}]");
            }

            string sqlViewColumns = String.Join(",", viewColums);

            try
            {
                sqlManager.ConnectByConfig();
                isConnected = true;
                sqlQuery = $"SELECT {sqlViewColumns} FROM View_ExcelRowsByExcelColNames WHERE ExcelFileId={excelFileId}";
                sqlQuery += $" AND QStatusId IN ({(int)QStatusNum.RecordingOk},{(int)QStatusNum.RecordingFailed})";
                //sqlQuery += String.Format(" AND QStatusId IN ({0},{1})", (int)QStatusNum.RecordingOk, (int)QStatusNum.RecordingFailed);
                System.Data.DataTable dtExcelRows = sqlManager.ExecuteQuery(sqlQuery);

                ExcelManager excelManager = new ExcelManager();
                excelManager.CreateExcel(exceldestFileName, "Feldolgozási eredmények");
                excelManager.DataTableToWorksheet(dtExcelRows, excelManager.ExcelSheet);
                excelManager.SaveAndCloseExcel();

                // kész
                sqlManager.Disconnect();
                isConnected = false;
            }
            catch (Exception ex)
            {
                // Connected?
                if (isConnected)
                {
                    sqlManager.Disconnect();
                }

                throw new Exception(ex.Message);
            }

            return isOk;
        }

        public static bool CreateOneResultExcel()
        {
            bool IsOk = false;

            return IsOk;    
        }
    }
}
