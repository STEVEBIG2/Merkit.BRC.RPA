using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Text.RegularExpressions;
using Merkit.BRC.RPA;
using Merkit.RPA.PA.Framework;
using System.IO;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using static System.Net.Mime.MediaTypeNames;
using System.Reflection.Emit;
using System.Data.SqlClient;
using System.Net;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.ComponentModel;
using UnitTest;
using System.Data.SqlTypes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;

namespace UnitTestProject1
{

    [TestClass]
    public class UnitTest1
    {
        public static ExcelManager excelManager = new ExcelManager();
        public static ExcelManager excelManager2 = new ExcelManager();
        private const string PasswordName = "UiPath: Enter Hungary";
        private const string UserName = "istvan.nagy@merkit.hu";
        private const string Password = "Qw52267660";
        private const string ExcelFileName = @"c:\RPA\Munka\Teszt_adatok_hibaval.xlsx";
        // String.Format("Data Source={0};Initial Catalog={1};User Id={2};Password={3};Application Name={4};Connect Timeout={5};Encrypt=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False; MultipleActiveResultSets=True", in_Config.MsSqlHost, in_Config.MsSqlDatabase, userName, password, in_Config.AppCode, 30)

        public void InitConfig()
        {
            Config.AppName = "UnitTest";
            Config.LogLevel = 0;
            Config.LogFileName = @"c:\Munka\Work\log_{0}.txt";
            Config.NotifyEmail = "rendszergazda@merkit.hu";
            //
            Config.MsSqlHost = @"STEVE-LAPTOP\SQLEXPRESS";
            Config.MsSqlDatabase = "BRC_Hungary_Test";
            Config.MsSqlUserName = "BRCHungaryUserTest";
            Config.MsSqlPassword = "Qw52267660";  
        }


        [TestMethod]
        public void TestValidZip()
        {
            // (1011-9999)
            string zip = @"1011";
            //string pattern = @"^[0-9]{4}$";
            bool isOk = Regex.IsMatch(zip, @"^[0-9]{4}$");
            isOk = isOk && Convert.ToInt32(zip) >= 1011; 
            Assert.IsTrue(isOk);
        }

        [TestMethod]
        public void TestValidKSH()
        {
            // 12345678 1239 123 12
            string KSH = @"12345678 1234 123 12";
            bool isOk = Regex.IsMatch(KSH, @"^\d\d\d\d\d\d\d\d \d\d\d\d \d\d\d [012]\d$");
            Assert.IsTrue(isOk);
        }

        [TestMethod]
        public void TestValidURL()
        {
            string URL = @"https://merkithu0-my.sharepoint.com/:b:/g/personal/istvan_nagy_merkit_hu/ESot8UKUE2RCqnGzu9X3EksBmShPzLhfd33vLvd2mCThfA?e=lSS4RQ";
            string Pattern = @"^(?:http(s)?:\/\/)?[\w.-]+(?:\.[\w\.-]+)+[\w\-\._~:/?#[\]@!\$&'\(\)\*\+,;=.]+$";
            Regex Rgx = new Regex(Pattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);
            bool isOk = Rgx.IsMatch(URL);
            Assert.IsTrue(isOk);
        }

        [TestMethod]
        public void TestRegex()
        {
            string result = "";
            string pattern = "";

            string yearPattern = @"(19|20)\d{2}";
            string monthPattern = @"(0[1-9]|1[1,2])";
            string dayPattern = @"(0[1-9]|[12][0-9]|3[01])";
            string separator = @"(\/|-)";

            pattern = yearPattern + separator + monthPattern + separator + dayPattern;
            pattern = dayPattern + separator + monthPattern + separator + yearPattern;
            //pattern = yearPattern + separator + monthPattern + separator + dayPattern;

            string input = @"25-07-2023";
            RegexOptions options = RegexOptions.Multiline;

            foreach (Match m in Regex.Matches(input, pattern, options))
            {
                result = String.Format("'{0}' found at index {1}.", m.Value, m.Index);
            }

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void TestBankHolidays()
        {
            // set once at the beginning of the year
            string[] extraWorkDays = { "0427", "1201" }; // extra work days 
            string[] extraDaysOff = { "0101", "1227" }; // extra days off (for example holidays)

            // test values
            DateTime workday = new DateTime(2024, 04, 10);
            DateTime weekendDay = new DateTime(2024, 04, 14);
            DateTime workDaySaturday = new DateTime(2024, 04, 27);
            DateTime dayOffMonday = new DateTime(2024, 01, 01);

            // change this value for test
            DateTime dt = dayOffMonday;

            // *** put it into function
            string monthDay = dt.ToString("MMdd");
            bool isWorkDay = !(dt.DayOfWeek == DayOfWeek.Saturday || dt.DayOfWeek == DayOfWeek.Sunday);

            if (isWorkDay)
            {
                isWorkDay = isWorkDay && !extraDaysOff.Contains(monthDay);
            }
            else
            {
                isWorkDay = isWorkDay || extraWorkDays.Contains(monthDay);
            }

            // return isWorkDay;

            Assert.IsTrue(true);
        }


        [TestMethod]
        public void TestCreateExcel()
        {
            bool isOk = excelManager.CreateExcel(@"c:\Munka\Test_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx", "Ez@neve");

            if(isOk)
            {
                excelManager.CloseExcelWithoutSave();
            }

            Assert.IsTrue(isOk);
        }

        [TestMethod]
        public void TestDeleteSheetIfExist()
        {
            bool isOk = excelManager.OpenExcel(ExcelFileName);

            if (isOk)
            {
                excelManager.DeleteSheetIfExist("Sheet 2");
                excelManager.SaveAndCloseExcel();
            }

            Assert.IsTrue(isOk);
        }

        [TestMethod]
        public void TestCopySheetsToNewExcel()
        {
            // params
            int excelFileId = 1;
            string adminName = "Admin"; // Error - rendszergazda@merkit.hu, Reference
            string excelSourceFileName = @"c:\RPA\Munka\Teszt_adatok_hibaval.xlsx";
            string destRootFolder = @"c:\RPA\EmailAttachments";
            List<string> fixSheets = new List<string>() { "Error - rendszergazda@merkit.hu" };

            bool isOk = Dispatcher.CopySheetToNewExcel(excelManager, excelFileId, adminName, excelSourceFileName, destRootFolder, fixSheets);
            Assert.IsTrue(isOk);
        }

        [TestMethod]
        public void TestCreateErrorExcels()
        {
            int excelFileId = 22;
            string excelSourceFileName = @"c:\RPA\Munka\Teszt_adatok_hibaval.xlsx";
            string destRootFolder = @"c:\RPA\EmailAttachments";
            string sysAdminName = "rendszergazda@merkit.hu";

            InitConfig();
            MSSQLManager sqlManager = new MSSQLManager();
            sqlManager.ConnectByConfig();

            bool isOk = Dispatcher.CreateErrorExcels(sqlManager,excelFileId, excelSourceFileName, destRootFolder, sysAdminName);

            sqlManager.Disconnect();
            Assert.IsTrue(isOk);
        }

        [TestMethod]
        public void TestCopyExcelRows()
        {
            int lastRowMunka = 0;
            int lastRowUj = 0;
            Range dest;

            string excelFileName = @"C:\Munka\Teszt_adatok.xlsx";
            bool isOk = excelManager.OpenExcel(excelFileName);

            if(isOk)
            {
                excelManager.SelectWorksheetByName("Munka1");
                Range headerRange = excelManager.ReadEntireRow("A1");
                lastRowMunka = excelManager.LastRow();

                for(int i = 2;  i<= lastRowMunka; i++)
                {
                    excelManager.SelectWorksheetByName("Munka1");
                    Range copyRange = excelManager.ReadEntireRow("A"+ i.ToString());
                    
                    // létezik a munkalap?
                    if(excelManager.AddNewSheetIfNotExist("Új"))
                    {
                        dest = excelManager.GetCellRange("A1");
                        headerRange.Copy(dest);
                    }

                    lastRowUj = excelManager.LastRow()+1;
                    dest = excelManager.GetCellRange("A" + lastRowUj.ToString());
                    copyRange.Copy(dest);
                }

                excelManager.SelectWorksheetByName("Új");
                excelManager.AutoFit();

                excelManager.SaveAndCloseExcel();
            }
            
            Assert.IsTrue(isOk);
        }

        [TestMethod]
        public void TestDispatcherMainProcess()
        {
            string rootDir = @"C:\RPA";
            string inputDir = Path.Combine(rootDir, "Input");
            string workDir = Path.Combine(rootDir, "Munka");

            InitConfig();
            bool isOk = Dispatcher.MainProcess(inputDir, workDir);
            Assert.IsTrue(isOk);
        }

        [TestMethod]
        public void TestCreateExcelRowsSQLScripts()
        {
            string excelColName = "";
            string sqlcolName = "";
            string sqlType = "";
            string sqlNotNull = "";
            List<string> createColums = new List<string>();
            List<string> createIndexes = new List<string>();
            List<string> foreignKeys = new List<string>();
            List<string> view1Colums = new List<string>();
            List<string> view2Colums = new List<string>();
            //
            List<string> procParameters = new List<string>();
            List<string> insertPart1 = new List<string>();
            List<string> insertPart2 = new List<string>();

            foreach (ExcelCol excelCol in ExcelValidator.excelHeaders.Where(x => !String.IsNullOrEmpty(x.SQLColName)))
            {
                excelColName = excelCol.ExcelColName; 
                sqlcolName = excelCol.SQLColName;
                sqlType = "???";
                sqlNotNull = "";

                switch (excelCol.ExcelColType)
                {
                    case ExcelColTypeNum.Text:
                        sqlType = excelCol.ExcelColRole.Equals(ExcelColRoleNum.ZipCode) ? "VARCHAR(10)" : "VARCHAR(150)";

                        if(excelCol.ExcelColRole.Equals(ExcelColRoleNum.ZipCode))
                        {
                            createIndexes.Add(String.Format("CREATE INDEX IX_ExcelRows_{0} ON ExcelRows({0})", sqlcolName));
                            createIndexes.Add("GO");
                            createIndexes.Add("");
                        }

                        break;
                    case ExcelColTypeNum.Number:
                        sqlType = "INT";
                        break;
                    case ExcelColTypeNum.Date:
                        sqlType = "DATE";
                        break;
                    case ExcelColTypeNum.DateTime:
                        sqlType = "DATE";
                        break;
                    case ExcelColTypeNum.Dropdown:
                        sqlType = "INT";
                        createIndexes.Add(String.Format("CREATE INDEX IX_ExcelRows_{0} ON ExcelRows({0})", sqlcolName));
                        createIndexes.Add("GO");
                        createIndexes.Add("");
                        //
                        foreignKeys.Add(String.Format("ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_{0} FOREIGN KEY({0}) REFERENCES DropDownsValues(DropDownsValueId)", sqlcolName));
                        foreignKeys.Add("GO");
                        foreignKeys.Add("");
                        foreignKeys.Add(String.Format("ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_{0}", sqlcolName));
                        foreignKeys.Add("GO");
                        foreignKeys.Add("");
                        break;
                    case ExcelColTypeNum.YesNo:
                        sqlType = "BIT";
                        break;
                    case ExcelColTypeNum.Link:
                        sqlType = "VARCHAR(150)";
                        break;
                    default:
                        break;
                }

                if(excelCol.ExcelColRequired == ExcelColRequiredNum.Yes)
                {
                    sqlNotNull = "NOT NULL";
                    procParameters.Add(String.Format("@{0} {1}", sqlcolName, sqlType));
                }
                else
                {
                    procParameters.Add(String.Format("@{0} {1}=NULL", sqlcolName, sqlType));
                }

                createColums.Add(String.Format("{0} {1} {2}", sqlcolName, sqlType, sqlNotNull).Trim());
                insertPart1.Add(sqlcolName);
                insertPart2.Add("@" + sqlcolName);    

                // view sor
                if (excelCol.ExcelColType == ExcelColTypeNum.Dropdown)
                {
                    view1Colums.Add(String.Format("(SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.{0}) AS [{1}]", sqlcolName, excelColName));
                    view2Colums.Add(String.Format("(SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.{0}) AS {0}", sqlcolName));
                }
                else
                {
                    if (excelCol.ExcelColType == ExcelColTypeNum.YesNo)
                    {
                        view1Colums.Add(String.Format("CASE WHEN {0}=1 THEN 'Igen' ELSE 'Nem' END AS [{1}]", sqlcolName, excelColName));
                        view2Colums.Add("r." + sqlcolName);
                    }
                    else
                    {
                        view1Colums.Add(String.Format("r.{0} AS [{1}]", sqlcolName, excelColName));
                        view2Colums.Add("r." + sqlcolName);
                    }
                }

            }

            // scripts
            string sqlScriptColumns = String.Join("," + Environment.NewLine, createColums);  // "\r\n"
            string sqlScriptIndexes = String.Join(Environment.NewLine, createIndexes);
            string sqlforeignKeys = String.Join(Environment.NewLine, foreignKeys);

            string sqlView1Columns = String.Join("," + Environment.NewLine, view1Colums);
            string sqlView2Columns = String.Join("," + Environment.NewLine, view2Colums);

            string sqlProcParameters = String.Join("," + Environment.NewLine, procParameters);
            string sqlinsertPart1 = String.Join("," + Environment.NewLine, insertPart1);
            string sqlinsertPart2 = String.Join("," + Environment.NewLine, insertPart2);

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void TestSqlQuery_View_DropDowns()
        {
            InitConfig();
            MSSQLManager sqlManager = new MSSQLManager();
            sqlManager.ConnectByConfig();

            foreach (ExcelCol col in ExcelValidator.excelHeaders.Where(x => x.ExcelColType == ExcelColTypeNum.Dropdown))
            {
                ExcelValidator.dropDownValuesbyType.Add(col.ExcelColName, new List<string>());
            }

            // ** aktuális munkalap ellenőrzése
            bool isOk = excelManager.OpenExcel(ExcelFileName);

            System.Data.DataTable dt = excelManager.WorksheetToDataTable(excelManager.ExcelSheet);

            Dispatcher.LoadDropdownValuesFromSQL(sqlManager, dt);

            sqlManager.Disconnect();
            excelManager.CloseExcelWithoutSave();

            // Write value to Json file
            string path = @"C:\Munka\Log_UnitTest_{0}.txt";
            var jsonLogFile = new JsonRepo<Dictionary<string, Dictionary<string, int>>>(String.Format(path, DateTime.Now.ToString("yyyyMMdd")));
            jsonLogFile.Write(ExcelValidator.dropDownIDsbyType);

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void TestExcelManager()
        {
            List<string> fejlecek = new List<string>() {
                "Munkavállaló: Azonosító",
                "Személy: Születési vezetéknév",
                "Személy: Születési keresztnév",
                "Személy: Útlevél száma/Személy ig.",
                "Munkavállaló: Munkakör megnevezése",
                "Munkavállaló: FEOR",
                "Személy: Vezetéknév",
                "Személy: Keresztnév",
                "Személy: Születési ország",
                "Személy: Születési hely",
                "Személy: Születési dátum",
                "Személy: Anyja vezetékneve",
                "Személy: Anyja keresztneve",
                "Személy: Neme",
                "Személy: Igazolványkép",
                "Személy: Állampolgárság",
                "Személy: Családi állapot",
                "Személy: Útlevél",
                "Személy: Magyarországra érkezést megelőző foglalkozás",
                "Személy: Útlevél kiállításának helye",
                "Személy: Útlevél kiállításának dátuma",
                "Személy: Útlevél lejáratának dátuma",
                "Személy: Várható jövedelem",
                "Személy: Várható jövedelem pénznem",
                "Személy: Tartózkodási engedély érvényessége",
                "Díjmentes-e","Engedély hosszabbítás-e",
                "Útlevél típusa","Iskolai végzettsége",
                "Munkavállaló: Irányítószám",
                "Munkavállaló: Település",
                "Munkavállaló: Közterület neve",
                "Munkavállaló: Közterület jellege",
                "Munkavállaló: Házszám",
                "Munkavállaló: HRSZ",
                "Munkavállaló: Épület",
                "Munkavállaló: Lépcsőház",
                "Munkavállaló: Emelet",
                "Munkavállaló: Ajtó",
                "Tartózkodás jogcíme",
                "Egészségbiztosítás",
                "Visszautazási ország",
                "Visszautazáskor közlekedési eszköz",
                "Visszautazás - útlevél van-e",
                "Érkezést megelőző ország",
                "Érkezést megelőző település",
                "Schengeni tartkózkodási okmány van-e",
                "Elutasított tartózkodási kérelem",
                "Büntetett előélet",
                "Kiutasították-e korábban",
                "Szenved-e gyógykezelésre szoruló betegségekben",
                "Kiskorú gyermek vele utazik-e",
                "Okmány átvétele",
                "Postai kézbesítés címe:",
                "Email cím",
                "Telefonszám",
                "Benyújtó",
                "Okmány átvétel külképviseleten?",
                "Átvételi ország",
                "Átvételi település",
                "Munkáltató rövid cégnév",
                "Munkáltató irányítószám",
                "Munkáltató település",
                "Munkáltató közterület neve",
                "Munkáltató közterület jellege",
                "Munkáltató házszám/hrsz",
                "TEÁOR szám",
                "KSH-szám",
                "Munkáltató adószáma/adóazonosító jele",
                "A foglalkoztatás munkaerő-kölcsönzés keretében történik",
                "Munkakörhöz szükséges iskolai végzettség",
                "Szakképzettsége",
                "Munkavégzés helye",
                "Munkavégzési irányítószám",
                "Munkavégzési település",
                "Munkavégzési közterület neve",
                "Munkavégzési közterület jellege",
                "Munkavégzési házszám/hrsz",
                "Munkavégzési Épület",
                "Munkavégzési Lépcsőház",
                "Munkavégzési Emelet",
                "Munkavégzési ajtó",
                "Foglalkoztatóval kötött megállapodás kelte",
                "Anyanyelve",
                "Magyar nyelvismeret",
                "Dolgozott-e korábban Magarországon?",
                "Feldolgozottsági Állapot",
                "Ügyszám",
                "Ellenőrzés Státusz",
                "Fájl Feltöltés Státusz" };

            //ExcelManager excelManager = new ExcelManager();
            bool isOk = excelManager.OpenExcel(@"c:\Munka\x.xlsx");
            isOk = excelManager.SelectWorksheetByName("Munka1");


            //excelManager.SetRangeValues("C5", "C10", new object[] { 1, 2, 3, 4, 5, 6 });
            //var x = excelManager.ReadCellValue("C5");
            //excelManager.SetRangeColor("C5", "C10", Color.Red);
            //excelManager.InsertFirstColumn("Kukukcs");
            DataTable dt = excelManager.WorksheetToDataTable(excelManager.ExcelSheet, true);

            foreach (string fejlec in fejlecek)
            {
                if(! dt.Columns.Contains(fejlec))
                {
                    excelManager.InsertFirstColumn(fejlec);
                    excelManager.SetCellColor(1,1, System.Drawing.Color.Green); // LightCoral
                }
            }

            if (isOk)
            {
                excelManager.SaveAndCloseExcel();
            }

            Assert.IsTrue(isOk);
        }

        [TestMethod]
        public void TestReadTextFile()
        {
            string path = Path.Combine("C:\\Merkit\\BRC_EnterHungary\\Textfiles", "állampolgárság.txt");
            string content = FileManager.ReadTextFile(path);

            Assert.IsTrue(!String.IsNullOrEmpty(content));
        }

        [TestMethod]
        public void TestLogger()
        {
            string path = Path.Combine("C:\\Munka", "Log_{0}.txt");
            FileManager.Logger(0, 0, path, "Init", "Info", "TestLogger", "Teszt bejegyzés", "1234");
            FileManager.Logger(0, 0, path, "Init", "Info", "TestLogger", "Teszt bejegyzés 2", "1235");
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void TestSaveWindowsCredential()
        {
            PasswordRepository.SaveWindowsCredential(PasswordName, UserName, Password);
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void TestGetWindowsCredential()
        {
            string userName = null;
            string password = null;
            bool isPassword = PasswordRepository.GetWindowsCredential(PasswordName, ref userName, ref password);
            Assert.AreEqual(userName, UserName, password, Password);
        }

        [TestMethod]
        public void TestDeleteWindowsCredential()
        {
            PasswordRepository.DeleteWindowsCredential(PasswordName);
            Assert.IsTrue(true);
        }
    }


}
