using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Merkit.BRC.RPA
{
    public static class Dispatcher
    {

        /// <summary>
        /// Main Process
        /// </summary>
        /// <returns></returns>
        public static bool MainProcess(string excelFileName)
        {
            bool processOk = false;
            MSSQLManager sqlManager = new MSSQLManager();
            SqlTransaction tr = null;
            sqlManager.ConnectByConfig();

            try
            {
                // ügyintéző login adatok begyűjtése
                ExcelValidator.GetEnterHungaryLogins(sqlManager);

                // *** dropdown lista ellenőrzéshez előkészülés
                foreach (ExcelCol col in ExcelValidator.excelHeaders.Where(x => x.ExcelColType == ExcelColTypeNum.Dropdown && x.ExcelColRole != ExcelColRoleNum.ZipCode))
                {
                    ExcelValidator.dropDownValuesbyType.Add(col.ExcelColName, new List<string>());
                }

                // excel feldolgozás
                tr = sqlManager.BeginTransaction();
                int excelFileId = ExcelValidator.InsertExcelFileProc(excelFileName, sqlManager, tr);
                processOk = ExcelValidator.ExcelWorkbookValidator(excelFileName, excelFileId, sqlManager, tr);

                sqlManager.Commit(tr);
            }
            catch (Exception ex)
            {
                Framework.Logger(0, "MainProcess", "Err", "", "-", String.Format("MainProcess hiba: {0}", ex.Message));

                if (tr != null)
                {
                    sqlManager.Rollback(tr);
                }

                throw new Exception(ex.Message); ;
            }
            finally
            {
                sqlManager.Disconnect();
            }

            return processOk;
        }
    }
}
