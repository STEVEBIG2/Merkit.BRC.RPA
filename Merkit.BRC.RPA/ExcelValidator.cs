using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Merkit.RPA.PA.Framework;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using Microsoft.Office.Interop.Excel;

namespace Merkit.BRC.RPA
{

    /// <summary>
    /// BRC_Enterhungary input excel ellenőrzése
    /// </summary>
    public static class ExcelValidator
    {
        #region public változók

        public static string TextFilePath { get; set; }

        public static Dictionary<string, string> loadDropdownDict = new Dictionary<string, string>();

        public static List<string> fejlecek = new List<string>() {
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

        #endregion

        #region Public függvények

        /// <summary>
        /// Az oldalon lévő , a flowhoz szükséges dropdown elemek értékeit  betölti a  lementett txt fájlokból
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool LoadDropdownValues(string path)
        {
            loadDropdownDict.Clear();

            string[] dropdownType = {
                "állampolgárság", "átvételi ország", "benyújtó", "családi állapot", "egészségbiztosítás",
                "előző ország", "FEOR", "iskolai végzettség", "munkakör iskolai végzettség", "munkáltató közterület jellege",
                "nem", "nemzetiség", "nyelv", "pénznem", "szállás emelet",
                "szállás közterület jellege", "szállás tartózkodási jogcíme", "szül_ország", "TEÁOR", "továbbut ország",
                "útlevél tipus", "zipcode"
            };

            foreach (string type in dropdownType)
            {
                loadDropdownDict.Add(
                    String.Format("{0}_dropdown", type.Replace(" ", "_")),
                    FileManager.ReadTextFile(Path.Combine(path, type + ".txt"))
                    );
            }

            return true;
        }

        /// <summary>
        /// Excel Header Validator
        /// </summary>
        /// <param name="excelFileName"></param>
        /// <returns></returns>
        public static bool ExcelHeaderValidator(string excelFileName)
        {
            bool isOk = ExcelManager.OpenExcel(excelFileName);
            bool isHeaderOk = true;
            System.Data.DataTable dt = ExcelManager.WorksheetToDataTable(ExcelManager.ExcelSheet, true);

            foreach (string fejlec in fejlecek)
            {
                if (!dt.Columns.Contains(fejlec))
                {
                    ExcelManager.InsertFirstColumn(fejlec);
                    ExcelManager.SetCellColor("A1", System.Drawing.Color.LightCoral);
                    isHeaderOk = false;
                }
            }

            if (isOk)
            {
                ExcelManager.CloseExcel();
            }

            return isOk && isHeaderOk;
        }

        /// <summary>
        /// Excel Rows Validator
        /// </summary>
        /// <param name="excelFileName"></param>
        /// <returns></returns>
        public static bool ExcelRowsValidator(string excelFileName)
        {
            Dictionary<string, bool> oszlopok_dictionary = new Dictionary<string, bool>() {
                {"Foglalkoztatóval kötött megállapodás kelte",true},
                {"Személy: Tartózkodási engedély érvényessége",true},
                {"Személy: Útlevél lejáratának dátuma",true},
                {"Személy: Útlevél kiállításának dátuma",true},
                {"Személy: Születési dátum",true},
                {"Munkavállaló: Emelet",true},
                {"Munkavégzési Emelet",true},
                {"Munkavégzési közterület jellege",true},
                {"Személy: Születési ország",true},
                {"Személy: Neme",true},
                {"Személy: Állampolgárság",true},
                {"Személy: Családi állapot",true},
                {"Iskolai végzettsége",true},
                {"Útlevél típusa",true},
                {"Munkavállaló: Irányítószám",true},
                {"Munkavállaló: Közterület neve",true},
                {"Tartózkodás jogcíme",true},
                {"Egészségbiztosítás",true},
                {"Visszautazási ország",true},
                {"Érkezést megelőző ország",true},
                {"Email cím",true},{"Benyújtó",true},
                {"Átvételi ország",true},
                {"Személy: Várható jövedelem",true},
                {"Személy: Várható jövedelem pénznem",true},
                {"Munkáltató irányítószám",true},
                {"Munkáltató közterület jellege",true},
                {"TEÁOR szám",true},
                {"KSH-szám",true},
                {"Munkáltató adószáma/adóazonosító jele",true},
                {"Munkakörhöz szükséges iskolai végzettség",true},
                {"Munkavégzési irányítószám",true},
                {"Munkavállaló: FEOR",true},
                {"Anyanyelve",true},
                {"Munkavállaló: Házszám", true},
                {"Munkavállaló: HRSZ",true}
            };

            bool isOk = ExcelManager.OpenExcel(excelFileName);
            System.Data.DataTable dt = ExcelManager.WorksheetToDataTable(ExcelManager.ExcelSheet);
            Dictionary<string, string> dictExcelColumnNameToExcellCol = ExcelManager.GetExcelColumnNamesByDataTable(dt);

            int rowNum = 1;

            // összes sor 
            foreach (DataRow currentRow in dt.Rows)
            {
                rowNum++;                

                // nem kihagyandó tétel?
                if(! ExcelManager.GetDataRowValue(currentRow, "Ellenőrzés Státusz").ToLower().Contains("pass"))
                {
                    CsaladiAllapot(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

                    // *** követ
                }

            }

            return true;
        }

        #endregion

        #region Private függvények (oszloponkénti ellenőrzések)

        private static void CsaladiAllapot(DataRow currentRow, int rowNum, ref Dictionary<string, string> fieldList)
        {
            // *** "Nőtlen/hajadon"
            string colName = "Személy: Családi állapot";
            string cellValue = ExcelManager.GetDataRowValue(currentRow, colName).ToLower();
            string cellName = fieldList[colName] + rowNum.ToString();

            if (cellValue.Equals("nőtlen") || cellValue.Equals("hajadon"))
            {
                ExcelManager.SetCellValue(cellName, "Nőtlen/hajadon");
            }
        }

        #endregion
    }
}
