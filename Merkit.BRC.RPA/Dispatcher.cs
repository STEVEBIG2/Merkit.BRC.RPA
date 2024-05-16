using Merkit.RPA.PA.Framework;
using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace Merkit.BRC.RPA
{
    public enum QStatusNum
    {
        Deleted = -2,
        Locked = -1,
        New = 0,
        CheckingInProgress = 1,
        CheckedOk = 2,
        CheckedFailed = 3,
        RecordingInProgress = 11,
        RecordingOk = 12,
        RecordingFailed = 13,
        Exported = 14
    };

    public class EnterHungaryLogin
    {
        public int EnterHungaryLoginId { get; set; }
        public string Email { get; set; }
        public string PasswordText { get; set; }

        public EnterHungaryLogin(int enterHungaryLoginId, string email, string passwordText)
        {
            this.EnterHungaryLoginId = enterHungaryLoginId;
            this.Email = email;
            this.PasswordText = passwordText;
        }
    }

    public static class Dispatcher
    {
        public static Dictionary<string, EnterHungaryLogin> enterHungaryLogins = new Dictionary<string, EnterHungaryLogin>(); // ügyintézők
        public static List<string> zipCodes = new List<string>();


        private static System.Data.DataTable GetNextExcelFile(MSSQLManager sqlManager, string inputDir, string workDir)
        {
            int excelFileId = 0;
            string excelFileName = "";
            string sqlQuery = "SELECT TOP 1 ExcelFileId, ExcelFileName, QStatusId FROM ExcelFiles ";
            sqlQuery += "WHERE QStatusId IN ({0},{1}) AND RobotName='{2}' ";
            sqlQuery += "ORDER BY ExcelFileId";
            sqlQuery = String.Format(sqlQuery, (int)QStatusNum.New, (int)QStatusNum.CheckingInProgress, Environment.UserName);
            System.Data.DataTable dt = sqlManager.ExecuteQuery(sqlQuery);
            
            // nincs félbemaradt ellenörzött excel?
            if(dt.Rows.Count == 0)
            {
                // Input mappából excel mozgatása a munka mappába?
                excelFileName = FileManager.GetFileFromQueue(inputDir, "*.xlsx", workDir);

                if(String.IsNullOrEmpty(excelFileName))
                {
                    Framework.Logger(0, "MainProcess", "Info", "", "GetNextExcelFile", "Nincs több feldolgozandó excel file!");
                }
                else
                {
                    //betenni az SQL-be
                    excelFileId = InsertExcelFileProc(excelFileName, sqlManager);
                    sqlQuery = "SELECT TOP 1 ExcelFileId, ExcelFileName, QStatusId FROM ExcelFiles ";
                    sqlQuery += "WHERE ExcelFileId={0}";
                    sqlQuery = String.Format(sqlQuery, excelFileId);
                    dt = sqlManager.ExecuteQuery(sqlQuery);
                }

            }
            else
            {
                excelFileName = dt.Rows[0]["ExcelFileName"].ToString();

                // nem létezik az excel file?
                if(! File.Exists(excelFileName))
                {
                    throw new Exception(String.Format("Nem létező feldolgozandó excel file: {0}", excelFileName));
                }
            }

            return dt;
        }

        /// <summary>
        /// Dispatcher Main Process
        /// </summary>
        /// <returns></returns>
        public static bool MainProcess(string inputDir, string workDir)
        {
            bool processOk = false;
            string excelFileName = "";
            int excelFileId = 0;
            MSSQLManager sqlManager = InitDispatcher();

            try
            {
                System.Data.DataTable dt = GetNextExcelFile(sqlManager, inputDir, workDir);

                // excel file-ok feldolgozása
                while(dt.Rows.Count > 0)
                {
                    // excel feldolgozás
                    excelFileName = dt.Rows[0]["ExcelFileName"].ToString();
                    excelFileId = Convert.ToInt32(dt.Rows[0]["ExcelFileId"]);
                    processOk = ExcelValidator.ExcelWorkbookValidator(excelFileName, excelFileId, sqlManager);

                    // következő excel
                    dt = GetNextExcelFile(sqlManager, inputDir, workDir);
                }

                processOk = true;
            }
            catch (Exception ex)
            {
                Framework.Logger(0, "MainProcess", "Err", "", "-", String.Format("MainProcess hiba: {0}", ex.Message));
                throw new Exception(ex.Message); ;
            }
            finally
            {
                sqlManager.Disconnect();
            }

            return processOk;
        }

        /// <summary>
        /// Init Dispatcher
        /// </summary>
        /// <returns></returns>
        private static MSSQLManager InitDispatcher()
        {
            MSSQLManager sqlManager = new MSSQLManager();

            try
            {
                sqlManager.ConnectByConfig();

                // ügyintéző login adatok begyűjtése
                GetEnterHungaryLogins(sqlManager);

                // *** dropdown lista ellenőrzéshez előkészülés
                foreach (ExcelCol col in ExcelValidator.excelHeaders.Where(x => x.ExcelColType == ExcelColTypeNum.Dropdown && x.ExcelColRole != ExcelColRoleNum.ZipCode))
                {
                    ExcelValidator.dropDownValuesbyType.Add(col.ExcelColName, new List<string>());
                }

            }
            catch (Exception ex)
            {
                sqlManager = null;
                throw new Exception(String.Format("InitDispatcher hiba: {0}", ex.Message));
            }

            return sqlManager;
        }

        /// <summary>
        /// GetEnterHungaryLogins
        /// </summary>
        /// <param name="sqlManager"></param>
        /// <returns></returns>
        public static void GetEnterHungaryLogins(MSSQLManager sqlManager)
        {
            int enterHungaryLoginId = 0;
            string email = "";
            string passwordText = "";
            System.Data.DataTable dt = sqlManager.ExecuteQuery("SELECT EnterHungaryLoginId, Email, PasswordText FROM EnterHungaryLogins WHERE Deleted=0");

            foreach (DataRow row in dt.Rows)
            {
                enterHungaryLoginId = Convert.ToInt32(row[0]);
                email = row[1].ToString().ToLower();
                passwordText = row[2].ToString();
                enterHungaryLogins.Add(email, new EnterHungaryLogin(enterHungaryLoginId, email, passwordText));
            }

            return;
        }

        /// <summary>
        /// A munkalapon lévő, a flowhoz szükséges dropdown elemek értékeit betölti SQL-ből
        /// </summary>
        /// <param name="sqlManager"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static bool LoadDropdownValuesFromSQL(MSSQLManager sqlManager, System.Data.DataTable dt, SqlTransaction tr = null)
        {
            string cellValue = "";
            string dropDownValue = "";

            // dropdown oszlopok típusa alapján kódlista készítése
            foreach (DataRow excelRow in dt.Rows)
            {
                foreach (string colName in ExcelValidator.dropDownValuesbyType.Keys)
                {
                    cellValue = excelRow[colName].ToString();

                    if (!ExcelValidator.dropDownValuesbyType[colName].Contains(cellValue))
                    {
                        ExcelValidator.dropDownValuesbyType[colName].Add(cellValue);
                    }
                }
            }

            // dropdown oszlopok típusa alapján kódok kigyűjtése
            string sqlParams = "";

            foreach (string colName in ExcelValidator.dropDownValuesbyType.Keys)
            {
                if (!ExcelValidator.dropDownIDsbyType.ContainsKey(colName))
                {
                    ExcelValidator.dropDownIDsbyType.Add(colName, new Dictionary<string, int>());
                }

                List<string> sqlColNames = ExcelValidator.dropDownValuesbyType[colName];
                string[] array = sqlColNames.ToArray();

                if (array != null && array.Length > 0)
                {
                    for (int i = 0; i < array.Length; i++)
                    {
                        array[i] = String.Format("'{0}'", array[i]);
                    }

                    sqlParams = String.Join(",", array);
                    string sqlCmd = String.Format("SELECT * FROM View_DropDowns WHERE ExcelColNames='{0}' AND DropDownValue IN ({1})", colName, sqlParams);
                    dt = sqlManager.ExecuteQuery(sqlCmd, tr);

                    foreach (DataRow dr in dt.Rows)
                    {
                        dropDownValue = dr["DropDownValue"].ToString().ToLower();

                        if (!ExcelValidator.dropDownIDsbyType[colName].ContainsKey(dropDownValue))
                        {
                            ExcelValidator.dropDownIDsbyType[colName].Add(dropDownValue, Int32.Parse(dr["DropDownsValueId"].ToString()));
                        }
                    }
                }

            }

            return true;
        }

        /// <summary>
        /// A munkalapon lévő, a flowhoz szükséges irányítószámok értékeit betölti SQL-ből
        /// </summary>
        /// <param name="sqlManager"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static bool LoadZipCodeValuesFromSQL(MSSQLManager sqlManager, System.Data.DataTable dt, SqlTransaction tr = null)
        {
            string sqlCmd = "";
            string cellValue = "";

            // minden sorból kigyűjtés
            foreach (DataRow excelRow in dt.Rows)
            {
                // aktuális sor irányítószám oszlopainak átnézése
                foreach (ExcelCol col in ExcelValidator.excelHeaders.Where(x => x.ExcelColRole == ExcelColRoleNum.ZipCode))
                {
                    cellValue = excelRow[col.ExcelColName].ToString();

                    // nem volt még ilyen irányítószám kigyűjtve?
                    if (!String.IsNullOrEmpty(cellValue) && ! zipCodes.Contains(cellValue))
                    {
                        sqlCmd = String.Format("SELECT COUNT(*) FROM ZipCodes WHERE ZipCode='{0}' AND DELETED=0", cellValue);

                        // létező irányítószám?
                        if (Convert.ToInt32(sqlManager.ExecuteScalar(sqlCmd, tr)) > 0)
                        {
                            zipCodes.Add(cellValue);
                        }
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Call InsertExcelFileProc stored procedure
        /// </summary>
        /// <param name="excelFileName"></param>
        /// <param name="sqlManager"></param>
        /// <param name="tr"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static int InsertExcelFileProc(string excelFileName, MSSQLManager sqlManager, SqlTransaction tr = null)
        {
            int result = -1;

            try
            {
                result = sqlManager.ExecuteProcWithReturnValue(
                    "InsertExcelFile",
                    new Dictionary<string, object>() {
                        { "@ExcelFileName", excelFileName },
                        { "@RobotName", Environment.UserName }
                    },
                    tr);

            }
            catch (SqlException ex)
            {
                throw new Exception("SqlException: " + ex.Message);
            }

            return result;
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
        public static int InsertExcelSheetProc(int excelFileId, string excelSheetName, int qStatusId, MSSQLManager sqlManager, SqlTransaction tr = null)
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

        /// <summary>
        /// Call InsertExcelSheetProc stored procedure
        /// </summary>
        /// <param name="excelFileId"></param>
        /// <param name="excelSheetName"></param>
        /// <param name="sqlManager"></param>
        /// <param name="tr"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static int InsertExcelRowProc(int excelFileId, int excelSheetId, int excelRownNum, DataRow dr, MSSQLManager sqlManager, SqlTransaction tr = null)
        {
            string[] yesValues = { "igen", "yes", "true" };
            int result = -1;
            string colStrValue = "";
            string dropdownValue = "";
            int colIntValue = -1;
            string ugyintezoValue = ExcelManager.GetDataRowValue(dr, "Ügyintéző").ToLower();

            Dictionary<string, object> paramsDict = new Dictionary<string, object>()
            {
                { "@ExcelFileId", excelFileId },
                { "@ExcelSheetId", excelSheetId },
                { "@ExcelRowNum", excelRownNum },
                { "@EnterHungaryLoginId", enterHungaryLogins[ugyintezoValue].EnterHungaryLoginId }
                //{ "@RobotName", Environment.UserName }
            };

            // paraméterek összeállítása
            foreach (ExcelCol excelCol in ExcelValidator.excelHeaders.Where(x => !String.IsNullOrEmpty(x.SQLColName)))
            {
                colStrValue = dr[excelCol.ExcelColName].ToString();

                // üres érték?
                if (String.IsNullOrEmpty(colStrValue))
                {
                    paramsDict.Add("@" + excelCol.SQLColName, null);
                }
                else
                {
                    // dropdown?
                    if (excelCol.ExcelColType == ExcelColTypeNum.Dropdown)
                    {
                        dropdownValue = dr[excelCol.ExcelColName].ToString().ToLower();

                        try
                        {
                            colIntValue = ExcelValidator.dropDownIDsbyType[excelCol.ExcelColName][dropdownValue];
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(ex.Message + ". " + excelCol.ExcelColName + " -> " + dropdownValue);
                        }

                        paramsDict.Add("@" + excelCol.SQLColName, colIntValue);
                    }
                    else
                    {
                        if (excelCol.ExcelColType == ExcelColTypeNum.Date)
                        {
                            colStrValue = colStrValue.Length > 10 ? colStrValue.Replace(" ", "").Substring(0, 10) : colStrValue;
                        }

                        if (excelCol.ExcelColType == ExcelColTypeNum.YesNo)
                        {
                            colStrValue = yesValues.Contains(colStrValue.ToLower()) ? "1" : "0";
                        }

                        paramsDict.Add("@" + excelCol.SQLColName, colStrValue);
                    }
                }

            }

            try
            {
                result = sqlManager.ExecuteProcWithReturnValue(
                    "InsertExcelRow",
                    paramsDict,
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
