using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Merkit.RPA.PA.Framework;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using Microsoft.Office.Interop.Excel;
using Merkit.BRC.RPA;

namespace Merkit.BRC.RPA
{
    public enum ExcelColTypeNum
    {
        None = 0,
        Text = 1,
        Number = 2,
        Date = 3,
        DateTime = 4,
        Dropdown = 5,
        Link = 6
    };

    public enum ExcelColRequiredNum
    {
        No = 0,
        Yes = 1
    };

    public class ExcelCol
    {
        public string ColName { get; set; }
        public ExcelColTypeNum ColType { get; set; }
        public ExcelColRequiredNum ExcelColRequired { get; set; }

        public ExcelCol(string colName, ExcelColTypeNum colType, ExcelColRequiredNum excelColRequired)
        {
            this.ColName = colName;
            this.ColType = colType;
            this.ExcelColRequired = excelColRequired;
        }  
        
    }

    /// <summary>
    /// BRC_Enterhungary input excel ellenőrzése
    /// </summary>
    public static class ExcelValidator
    {
        #region public változók

        public static string TextFilePath { get; set; }

        public static Dictionary<string, string> loadDropdownDict = new Dictionary<string, string>();
        public static Dictionary<string, List<string>> loadDropdownList = new Dictionary<string, List<string>>();

        public static List<ExcelCol> excelHeaders = new List<ExcelCol>() {
                new ExcelCol("Munkavállaló: Azonosító", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Születési vezetéknév", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Születési keresztnév", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Útlevél száma/Személy ig.", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkavállaló: Munkakör megnevezése", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Munkavállaló: FEOR", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Vezetéknév", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Keresztnév", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Születési ország", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Születési hely", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Születési dátum", ExcelColTypeNum.Date, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Anyja vezetékneve", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Anyja keresztneve", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Neme", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Igazolványkép", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Állampolgárság", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Családi állapot", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Útlevél", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Magyarországra érkezést megelőző foglalkozás", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Személy: Útlevél kiállításának helye", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Útlevél kiállításának dátuma", ExcelColTypeNum.Date, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Útlevél lejáratának dátuma", ExcelColTypeNum.Date, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Várható jövedelem", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Várható jövedelem pénznem", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Személy: Tartózkodási engedély érvényessége", ExcelColTypeNum.Date, ExcelColRequiredNum.Yes),
                new ExcelCol("Díjmentes-e", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Engedély hosszabbítás-e", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Útlevél típusa",ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Iskolai végzettsége", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkavállaló: Irányítószám", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkavállaló: Település", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkavállaló: Közterület neve", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkavállaló: Közterület jellege", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkavállaló: Házszám", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Munkavállaló: HRSZ", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Munkavállaló: Épület", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Munkavállaló: Lépcsőház", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Munkavállaló: Emelet", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.No),
                new ExcelCol("Munkavállaló: Ajtó", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Tartózkodás jogcíme", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Egészségbiztosítás", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Visszautazási ország", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Visszautazáskor közlekedési eszköz", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Visszautazás - útlevél van-e", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Érkezést megelőző ország", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Érkezést megelőző település", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Schengeni tartkózkodási okmány van-e", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Elutasított tartózkodási kérelem", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Büntetett előélet", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Kiutasították-e korábban", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Szenved-e gyógykezelésre szoruló betegségekben", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Kiskorú gyermek vele utazik-e", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Okmány átvétele", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Postai kézbesítés címe:", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Email cím", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Telefonszám", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Benyújtó", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.No),
                new ExcelCol("Okmány átvétel külképviseleten?", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Átvételi ország", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.No),
                new ExcelCol("Átvételi település", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Munkáltató rövid cégnév", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkáltató irányítószám", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkáltató település", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkáltató közterület neve", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkáltató közterület jellege", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkáltató házszám/hrsz", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("TEÁOR szám", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("KSH-szám", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkáltató adószáma/adóazonosító jele", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("A foglalkoztatás munkaerő-kölcsönzés keretében történik", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkakörhöz szükséges iskolai végzettség", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Szakképzettsége", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Munkavégzés helye", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Munkavégzési irányítószám", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.No),
                new ExcelCol("Munkavégzési település", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Munkavégzési közterület neve", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Munkavégzési közterület jellege", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.No),
                new ExcelCol("Munkavégzési házszám/hrsz", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Munkavégzési Épület", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Munkavégzési Lépcsőház", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Munkavégzési Emelet", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Munkavégzési ajtó", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Foglalkoztatóval kötött megállapodás kelte", ExcelColTypeNum.Date, ExcelColRequiredNum.Yes),
                new ExcelCol("Anyanyelve", ExcelColTypeNum.Dropdown, ExcelColRequiredNum.Yes),
                new ExcelCol("Magyar nyelvismeret", ExcelColTypeNum.Text, ExcelColRequiredNum.Yes),
                new ExcelCol("Dolgozott-e korábban Magarországon?", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Feldolgozottsági Állapot", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Ügyszám", ExcelColTypeNum.Text, ExcelColRequiredNum.No),
                new ExcelCol("Ellenőrzés Státusz", ExcelColTypeNum.None, ExcelColRequiredNum.No),
                new ExcelCol("Fájl Feltöltés Státusz", ExcelColTypeNum.None, ExcelColRequiredNum.No)
            };

        #endregion

        #region Public függvények

        /// <summary>
        /// Az oldalon lévő, a flowhoz szükséges dropdown elemek értékeit  betölti a  lementett txt fájlokból
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool LoadDropdownValuesFromTextFiles(string path)
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
        /// Az oldalon lévő, a flowhoz szükséges dropdown elemek értékeit  betölti a  lementett txt fájlokból
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool LoadDropdownValuesFromSQL(string path)
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

            foreach (ExcelCol fejlec in excelHeaders)
            {
                if (!dt.Columns.Contains(fejlec.ColName))
                {
                    ExcelManager.InsertFirstColumn(fejlec.ColName);
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
            string checkStatuscellName = dictExcelColumnNameToExcellCol["Ellenőrzés Státusz"];
            bool isRowOk = true;
            bool isGoodRow = false;
            int rowNum = 1;

            // összes sor 
            foreach (DataRow currentRow in dt.Rows)
            {
                isRowOk = true;
                rowNum++;

                // feldolgozatlan sor?
                if(String.IsNullOrEmpty(ExcelManager.GetDataRowValue(currentRow, "Feldolgozottsági Állapot")))
                {
                    // nem kihagyandó tétel?
                    if (!ExcelManager.GetDataRowValue(currentRow, "Ellenőrzés Státusz").ToLower().Contains("pass"))
                    {
                        // Nőtlen v. hajadon -> Nőtlen/hajadon
                        isRowOk = isRowOk && CsaladiAllapot(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

                        // kötelező szöveges oszlopok ellenőrzése
                        isRowOk = isRowOk && AllRequiredFieldChecker(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

                        // *** Dátum átalakítás és ellenőrzés
                        isRowOk = isRowOk && AllDateCheckAndConvert(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);


                        // *** követ

                        // Ellenőrzés státusz állítása
                        ExcelManager.SetCellValue(checkStatuscellName + rowNum.ToString(), isRowOk ? "PASS" : "FAIL");
                        isGoodRow = isGoodRow || isRowOk; // van legalább egy jó sor
                    }
                    else
                    {
                        isGoodRow = true;
                    }
                }
            }
            
            ExcelManager.CloseExcel();
            return isGoodRow;
        }

        /// <summary>
        /// Excel Column Number By Name from DataTable
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static int ColumnNumberByName(string columnName, System.Data.DataTable dt)
        {
            int colNum = 0;

            foreach (DataColumn col in dt.Columns)
            {
                colNum++;

                if (col.ColumnName.ToLower() == columnName.ToLower())
                {
                    break;
                }               
            }

            return colNum;
        }

        #endregion

        #region Private függvények (oszloponkénti ellenőrzések)

        /// <summary>
        /// CsaladiAllapot
        /// </summary>
        /// <param name="currentRow"></param>
        /// <param name="rowNum"></param>
        /// <param name="fieldList"></param>
        private static bool CsaladiAllapot(DataRow currentRow, int rowNum, ref Dictionary<string, string> fieldList)
        {
            // *** "Nőtlen/hajadon"
            string colName = "Személy: Családi állapot";
            string cellValue = ExcelManager.GetDataRowValue(currentRow, colName).ToLower();
            string cellName = fieldList[colName] + rowNum.ToString();

            if (cellValue.Equals("nőtlen") || cellValue.Equals("hajadon"))
            {
                ExcelManager.SetCellValue(cellName, "Nőtlen/hajadon");
            }

            return true;
        }

        /// <summary>
        /// Check All Required Text Fields
        /// </summary>
        /// <param name="currentRow"></param>
        /// <param name="rowNum"></param>
        /// <param name="fieldList"></param>
        /// <param name="datumHeaderek"></param>
        private static bool AllRequiredFieldChecker(DataRow currentRow, int rowNum, ref Dictionary<string, string> fieldList)
        {
            bool isCellValueOk = true;

            // dátum oszlopokon végigmenni
            foreach (ExcelCol col in excelHeaders.Where(x => x.ExcelColRequired == ExcelColRequiredNum.Yes))
            {
                string cellValue = ExcelManager.GetDataRowValue(currentRow, col.ColName).ToLower();
                string cellName = fieldList[col.ColName] + rowNum.ToString();

                if (cellValue.Length == 0)
                {
                    isCellValueOk = false;
                    ExcelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
                }

            }

            return isCellValueOk;
        }

        /// <summary>
        /// Check And Convert All Date Fields
        /// </summary>
        /// <param name="currentRow"></param>
        /// <param name="rowNum"></param>
        /// <param name="fieldList"></param>
        /// <param name="datumHeaderek"></param>
        private static bool AllDateCheckAndConvert(DataRow currentRow, int rowNum, ref Dictionary<string, string> fieldList)
        {
            bool isCellValueOk = true;

            // dátum oszlopokon végigmenni
            foreach (ExcelCol col in excelHeaders.Where(x => x.ColType == ExcelColTypeNum.Date))
            {
                string cellValue = ExcelManager.GetDataRowValue(currentRow, col.ColName).ToLower();
                string cellName = fieldList[col.ColName] + rowNum.ToString();

                if (cellValue.Length > 0)
                {
                    DateTime dateTime = DateTime.MinValue;
                    bool isGoodDate = DateTime.TryParse(cellValue, out dateTime);
                }
                else
                {
                    isCellValueOk = isCellValueOk && col.ExcelColRequired == ExcelColRequiredNum.No;
                }

            }

            return isCellValueOk;
        }

        #endregion
    }
}
