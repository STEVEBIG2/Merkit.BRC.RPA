using Merkit.RPA.PA.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Linq;
using System.Text.RegularExpressions;

namespace Merkit.BRC.RPA
{
	public class DbManager
	{
		public DbManager()
		{
		}

        /// <summary>
        /// Call InsertExcelSheetProc stored procedure
        /// </summary>
        /// <param name="excelFileId"></param>
        /// <param name="excelSheetName"></param>
        /// <param name="qStatusId"></param>
        /// <param name="sqlManager"></param>
        /// <param name="tr"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public int InsertExcelSheetProc(int excelFileId, string excelSheetName, int qStatusId, MSSQLManager sqlManager, SqlTransaction tr = null)
        {
            int result = -1;

            try
            {
                result = sqlManager.ExecuteProcWithReturnValue(
                    "InsertExcelSheet",
                    new Dictionary<string, object>() {
                        { "@ExcelFileId", excelFileId },
                        { "@ExcelSheetName", excelSheetName },
                        { "@QStatusId", qStatusId },
                        { "@RobotName", Environment.UserName }
                    },
                    tr);

            }
            catch (SqlException ex)
            {
                sqlManager.Rollback(tr);
                throw new Exception("SqlException: " + ex.Message);
            }

            return result;
        }
    }
}
