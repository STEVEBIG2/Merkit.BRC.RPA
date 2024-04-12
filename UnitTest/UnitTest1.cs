using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Merkit.BRC.RPA;
using Merkit.RPA.PA.Framework;
using System.IO;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using static System.Net.Mime.MediaTypeNames;
using System.Text.RegularExpressions;
using System.Reflection.Emit;
using System.Data.SqlClient;
using System.Net;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.ComponentModel;
using UnitTest;

namespace UnitTestProject1
{

    [TestClass]
    public class UnitTest1
    {
        private const string PasswordName = "UiPath: Enter Hungary";
        private const string UserName = "istvan.nagy@merkit.hu";
        private const string Password = "Qw52267660";
        private const string ExcelFleName = @"c:\Munka\x-10.xlsx";
        // String.Format("Data Source={0};Initial Catalog={1};User Id={2};Password={3};Application Name={4};Connect Timeout={5};Encrypt=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False; MultipleActiveResultSets=True", in_Config.MsSqlHost, in_Config.MsSqlDatabase, userName, password, in_Config.AppCode, 30)

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
            bool isWorkDay = ! (dt.DayOfWeek == DayOfWeek.Saturday || dt.DayOfWeek == DayOfWeek.Sunday) ;

            if (isWorkDay)
            {
                isWorkDay = isWorkDay && ! extraDaysOff.Contains(monthDay);
            }
            else
            {
                isWorkDay = isWorkDay || extraWorkDays.Contains(monthDay);
            }

            // return isWorkDay;

            Assert.IsTrue(true);
        }


        [TestMethod]
        public void TestSqlQuery()
        {
            MSSQLManager sqlManager = new MSSQLManager(@"STEVE-LAPTOP\SQLEXPRESS", "BRC_Hungary_Test", "BRCHungaryUserTest", "Qw52267660", "BRC_EH_Test");
            SqlConnection connection = sqlManager.Connect();

            bool isOk = ExcelManager.OpenExcel(ExcelFleName);
            System.Data.DataTable dt = ExcelManager.WorksheetToDataTable(ExcelManager.ExcelSheet);

            //  ** dropdown oszlopok kigyűjtése kódlista készítéshez
            Dictionary<string, List<string>> dropDownValuesbyType = new Dictionary<string, List<string>>();
            Dictionary<string, Dictionary<string, int>> dropDownIDsbyType = new Dictionary<string, Dictionary<string, int>>();

            foreach (ExcelCol col in ExcelValidator.excelHeaders.Where(x => x.ExcelColType == ExcelColTypeNum.Dropdown))
            {
                dropDownValuesbyType.Add(col.ExcelColName, new List<string>());
            }

            string cellValue = "";

            // dropdown oszlopok típusa alapján kódlista készítése
            foreach(DataRow excelRow in dt.Rows)
            {
                foreach (string colName in dropDownValuesbyType.Keys)
                {
                    cellValue = excelRow[colName].ToString();

                    if (! dropDownValuesbyType[colName].Contains(cellValue))
                    {
                        dropDownValuesbyType[colName].Add(cellValue);
                    }
                }
            }

            // dropdown oszlopok típusa alapján kódok kigyűjtése
            string sqlParams = "";

            foreach (string colName in dropDownValuesbyType.Keys)
            {
                dropDownIDsbyType.Add(colName, new Dictionary<string, int>());

                List<string> sqlColNames = dropDownValuesbyType[colName];
                string[] array = sqlColNames.ToArray();

                if(array != null && array.Length > 0)
                {
                    for (int i = 0; i < array.Length; i++)
                    {
                        array[i] = String.Format("'{0}'", array[i]);
                    }

                    sqlParams = String.Join(",", array);
                    string sqlCmd = String.Format("SELECT * FROM View_DropDowns WHERE ExcelColNames='{0}' AND DropDownValue IN ({1})", colName, sqlParams);
                    dt = sqlManager.ExecuteQuery(sqlCmd);

                    foreach (DataRow dr in dt.Rows)
                    {
                        dropDownIDsbyType[colName].Add(dr["DropDownValue"].ToString(), Int32.Parse(dr["DropDownsValueId"].ToString()));
                    }
                }

            }

            sqlManager.Disconnect();
            ExcelManager.CloseExcel();

            // Write value to Json file
            string path = @"C:\Munka\Log_UnitTest_{0}.txt";
            var jsonLogFile = new JsonRepo<Dictionary<string, Dictionary<string, int>>>(String.Format(path, DateTime.Now.ToString("yyyyMMdd")));
            jsonLogFile.Write(dropDownIDsbyType);

            Assert.IsTrue(true);
        }


        [TestMethod]
        public void TestExcelHeaderValidator()
        {
            bool isOk = ExcelValidator.ExcelHeaderValidator(ExcelFleName);
            Assert.IsTrue(isOk);
        }

        [TestMethod]
        public void TestExcelRowsValidator()
        {
            bool isOk = ExcelValidator.ExcelRowsValidator(ExcelFleName);
            Assert.IsTrue(isOk);
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
            bool isOk = ExcelManager.OpenExcel(@"c:\Munka\x.xlsx");
            isOk = ExcelManager.SelectWorksheetByName("Munka1");


            //ExcelManager.SetRangeValues("C5", "C10", new object[] { 1, 2, 3, 4, 5, 6 });
            //var x = ExcelManager.ReadCellValue("C5");
            //ExcelManager.SetRangeColor("C5", "C10", Color.Red);
            //ExcelManager.InsertFirstColumn("Kukukcs");
            DataTable dt = ExcelManager.WorksheetToDataTable(ExcelManager.ExcelSheet, true);

            foreach (string fejlec in fejlecek)
            {
                if(! dt.Columns.Contains(fejlec))
                {
                    ExcelManager.InsertFirstColumn(fejlec);
                    ExcelManager.SetCellColor(1,1, System.Drawing.Color.Green); // LightCoral
                }
            }

            if (isOk)
            {
                ExcelManager.CloseExcel();
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
