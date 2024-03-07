using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Merkit.BRC.RPA;
using Merkit.RPA.PA.Framework;
using System.IO;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestExcelManager()
        {
            //ExcelManager excelManager = new ExcelManager();
            bool isOk = ExcelManager.OpenExcel(@"c:\Munka\x.xlsx");
            //ExcelManager.SetRangeValues("C5", "C10", new object[] { 1, 2, 3, 4, 5, 6 });
            var x = ExcelManager.ReadCellValue("C5");
            ExcelManager.SetRangeColor("C5", "C10", Color.Red);
            DataTable dt = ExcelManager.WorksheetToDataTable(ExcelManager.ExcelSheet);

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
    }


}
