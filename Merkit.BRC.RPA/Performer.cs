using Merkit.RPA.PA.Framework;
using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Data.SqlTypes;

namespace Merkit.BRC.RPA
{
    public static class Performer
    {
        /// <summary>
        /// Create Result Excels
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static bool CreateResultExcels()
        {
            bool isOk = true;
            bool isConnected = false;
            int excelFileId = 0;
            string excelDestFileName = "";
            MSSQLManager sqlManager = new MSSQLManager();
            string sqlViewSelect = CreateViewSelect();

            try
            {
                sqlManager.ConnectByConfig();
                isConnected = true;
                string sqlQuery = $"SELECT ExcelFileId, ExcelFileName FROM ExcelFiles WHERE QStatusId={(int)QStatusNum.RecordingOk}";
                System.Data.DataTable dtExcels = sqlManager.ExecuteQuery(sqlQuery);

                foreach (DataRow dr in dtExcels.Rows)
                {
                    excelFileId = Convert.ToInt32(dr["ExcelFileId"]);
                    excelDestFileName = Path.Combine(Config.EmailAttachmentsRootFolder, excelFileId.ToString(), String.Format("{0}_log.xlsx",Path.GetFileNameWithoutExtension(dr["ExcelFileName"].ToString())));
                    isOk = CreateOneResultExcel(sqlManager, sqlViewSelect, excelFileId, excelDestFileName);
                }

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

        /// <summary>
        /// Create One Result Excel
        /// </summary>
        /// <param name="sqlManager"></param>
        /// <param name="sqlViewSelect"></param>
        /// <param name="excelFileId"></param>
        /// <param name="excelDestFileName"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static bool CreateOneResultExcel(MSSQLManager sqlManager, string sqlViewSelect, int excelFileId, string excelDestFileName)
        {
            bool isOk = true;
            bool isExcelOpen = false;
            string sqlQuery = "";
            string sqlString = "";
            ExcelManager excelManager = new ExcelManager();
            SqlTransaction tr = sqlManager.BeginTransaction();

            try
            {
                // Read from SQL
                sqlQuery = String.Format(sqlViewSelect, excelFileId);
                System.Data.DataTable dtExcelRows = sqlManager.ExecuteQuery(sqlQuery, tr);

                // Create result excel, change status, create e-mail
                excelManager.CreateExcel(excelDestFileName, "Feldolgozási eredmények");
                isExcelOpen = true;
                excelManager.DataTableToWorksheet(dtExcelRows, excelManager.ExcelSheet);
                sqlString = String.Format(String.Format("UPDATE ExcelFiles SET QStatusId={0}, QStatusTime=getdate() WHERE ExcelFileId={1}", (int)QStatusNum.Exported, excelFileId));
                sqlManager.ExecuteNonQuery(sqlString, null, tr);
                Dispatcher.InsertEmailQueue(excelFileId, Config.NotifyEmail, Config.ResultExcelEmailSubject, Config.ResultExcelEmailBody, excelDestFileName, sqlManager, tr);

                // close excel, commit tran
                excelManager.SaveAndCloseExcel();
                isExcelOpen = false;
                tr.Commit();
            }
            catch (Exception ex)
            {
                if (isExcelOpen)
                {
                    excelManager.CloseExcelWithoutSave();
                }

                tr.Rollback();
                isOk = false;
                throw new Exception(ex.Message);
            }

            return isOk;
        }

        /// <summary>
        /// Create View_ExcelRowsByExcelColNames select command
        /// </summary>
        /// <returns></returns>
        private static string CreateViewSelect()
        {
            List<string> viewColums = new List<string>();
            viewColums.Add("[Státusz]");
            viewColums.Add("[Ügyszám]");

            foreach (ExcelCol excelCol in ExcelValidator.excelHeaders.Where(x => !String.IsNullOrEmpty(x.SQLColName)))
            {
                viewColums.Add($"[{excelCol.ExcelColName}]");
            }

            string sqlViewColumns = String.Join(",", viewColums);
            string sqlViewQuery = "SELECT " + sqlViewColumns + " FROM View_ExcelRowsByExcelColNames WHERE ExcelFileId={0}";
            sqlViewQuery += $" AND QStatusId IN ({(int)QStatusNum.RecordingOk},{(int)QStatusNum.RecordingFailed})";
            return sqlViewQuery;
        }
    }
}
