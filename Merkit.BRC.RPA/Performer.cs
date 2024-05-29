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
            int excelFileId = 0;
            string excelSourceFileName = "";
            string sysAdminName = Config.NotifyEmail;
            MSSQLManager sqlManager = new MSSQLManager();

            try
            {
                sqlManager.ConnectByConfig();
                isConnected = true;
                sqlQuery = "SELECT ExcelFileId, ExcelFileName FROM ExcelFiles ";
                sqlQuery += "WHERE QStatusId IN({0}, {1}) AND ErrorExcelsCreated=0 ";
                sqlQuery = String.Format(sqlQuery, (int)QStatusNum.CheckedOk, (int)QStatusNum.CheckedFailed);
                System.Data.DataTable dtExcelFiles = sqlManager.ExecuteQuery(sqlQuery);

                foreach (DataRow dr in dtExcelFiles.Rows)
                {
                    excelFileId = Convert.ToInt32(dr["ExcelFileId"]);
                    excelSourceFileName = dr["ExcelFileName"].ToString();
                    //isOk = CreateOneResultExcel(sqlManager, excelFileId, excelSourceFileName, destRootFolder, sysAdminName);
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

        public static bool CreateOneResultExcel()
        {
            bool IsOk = false;

            return IsOk;    
        }
    }
}
