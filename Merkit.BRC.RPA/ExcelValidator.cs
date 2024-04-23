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
        YesNo = 6,
        Link = 7
    };

    public enum ExcelColRequiredNum
    {
        No = 0,
        Yes = 1
    };

    public enum ExcelColRoleNum
    {
        None = 0,
        PastDate = 1,
        FutureDate = 2,
        ZipCode = 3,
        Regex = 4,
        CreateIfNoExists = 99
    };

    public class ExcelCol
    {
        public int ExcelColNum { get; set; }
        public string ExcelColName { get; set; }
        public ExcelColTypeNum ExcelColType { get; set; }
        public ExcelColRoleNum ExcelColRole { get; set; }
        public string ExcelColRoleExpression  { get; set; }
        public ExcelColRequiredNum ExcelColRequired { get; set; }
        public string SQLColName { get; set; }

        public ExcelCol(int excelColNum, string excelColName, ExcelColTypeNum excelColType, ExcelColRoleNum excelColRole, string excelColRoleExpression, ExcelColRequiredNum excelColRequired, string sqlColName)
        {
            this.ExcelColNum = excelColNum;   
            this.ExcelColName = excelColName;
            this.ExcelColType = excelColType;
            this.ExcelColRole = excelColRole;
            this.ExcelColRoleExpression = excelColRoleExpression;
            this.ExcelColRequired = excelColRequired;
            this.SQLColName = sqlColName;
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

        public static int excelColNum = 0;

        public static List<ExcelCol> excelHeaders = new List<ExcelCol>() {
                // new ExcelCol(++excelColNum, "Ügyszám", ExcelColTypeNum.Text, ExcelColRoleNum.CreateIfNoExists, null, ExcelColRequiredNum.No, "Ugyszam"),
                new ExcelCol(++excelColNum, "Ellenőrzés Státusz", ExcelColTypeNum.None, ExcelColRoleNum.CreateIfNoExists, null, ExcelColRequiredNum.No, ""),
                // new ExcelCol(++excelColNum, "Munkavállaló: Azonosító", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Mv_Azonosito"),
                new ExcelCol(++excelColNum, "Személy: Születési vezetéknév", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Szul_Vezeteknev"),
                new ExcelCol(++excelColNum, "Személy: Születési keresztnév", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Szul_Keresztnev"),
                new ExcelCol(++excelColNum, "Személy: Útlevél száma/Személy ig.", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Utlevel_Szig"),
                new ExcelCol(++excelColNum, "Munkavállaló: Munkakör megnevezése", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mv_Munkakor"),
                new ExcelCol(++excelColNum, "Munkavállaló: FEOR", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Mv_FEOR"),
                new ExcelCol(++excelColNum, "Személy: Vezetéknév", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Vezeteknev"),
                new ExcelCol(++excelColNum, "Személy: Keresztnév", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Keresztnev"),
                new ExcelCol(++excelColNum, "Személy: Születési ország", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Szul_Orszag"),
                new ExcelCol(++excelColNum, "Személy: Születési hely", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Szul_Hely"),
                new ExcelCol(++excelColNum, "Személy: Születési dátum", ExcelColTypeNum.Date, ExcelColRoleNum.PastDate, null, ExcelColRequiredNum.Yes, "Sz_Szul_Datum"),
                new ExcelCol(++excelColNum, "Személy: Anyja vezetékneve", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Anyja_Vezeteknev"),
                new ExcelCol(++excelColNum, "Személy: Anyja keresztneve", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Anyja_Keresztnev"),
                new ExcelCol(++excelColNum, "Személy: Neme", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Neme"),
                new ExcelCol(++excelColNum, "Személy: Igazolványkép", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Igazolvanykep"),
                new ExcelCol(++excelColNum, "Személy: Állampolgárság", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Allampolgarsag"),
                new ExcelCol(++excelColNum, "Személy: Családi állapot", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Csaladi_allapot"),
                new ExcelCol(++excelColNum, "Személy: Útlevél", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Utlevel"),
                new ExcelCol(++excelColNum, "Személy: Magyarországra érkezést megelőző foglalkozás", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Sz_Magy_erk_meg_fogl"),
                new ExcelCol(++excelColNum, "Személy: Útlevél kiállításának helye", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Utlevel_kiall_helye"),
                new ExcelCol(++excelColNum, "Személy: Útlevél kiállításának dátuma", ExcelColTypeNum.Date, ExcelColRoleNum.PastDate, null, ExcelColRequiredNum.Yes, "Sz_Utlevel_kiall_datuma"),
                new ExcelCol(++excelColNum, "Személy: Útlevél lejáratának dátuma", ExcelColTypeNum.Date, ExcelColRoleNum.FutureDate, null, ExcelColRequiredNum.Yes, "Sz_Utlevel_lejarat_datuma"),
                new ExcelCol(++excelColNum, "Személy: Várható jövedelem", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Varhato_jovedelem"),
                new ExcelCol(++excelColNum, "Személy: Várható jövedelem pénznem", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Varhato_jov_penznem"),
                new ExcelCol(++excelColNum, "Személy: Tartózkodási engedély érvényessége", ExcelColTypeNum.Date, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Tart_eng_erv"),
                new ExcelCol(++excelColNum, "Díjmentes-e", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Dijmentes"),
                new ExcelCol(++excelColNum, "Engedély hosszabbítás-e", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Engedely_hosszabbitas"),
                new ExcelCol(++excelColNum, "Útlevél típusa",ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Utlevel_tipusa"),
                new ExcelCol(++excelColNum, "Iskolai végzettsége", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Iskolai_vegzettsege"),
                new ExcelCol(++excelColNum, "Munkavállaló: Irányítószám", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Mv_Iranyitoszam"),
                new ExcelCol(++excelColNum, "Munkavállaló: Település", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Mv_Telepules"),
                new ExcelCol(++excelColNum, "Munkavállaló: Közterület neve", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Mv_Kozterulet_neve"),
                new ExcelCol(++excelColNum, "Munkavállaló: Közterület jellege", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Mv_Kozterulet_jellege"),
                new ExcelCol(++excelColNum, "Munkavállaló: Házszám", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mv_Hazszam"),
                new ExcelCol(++excelColNum, "Munkavállaló: HRSZ", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mv_HRSZ"),
                new ExcelCol(++excelColNum, "Munkavállaló: Épület", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mv_Epulet"),
                new ExcelCol(++excelColNum, "Munkavállaló: Lépcsőház", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mv_Lepcsohaz"),
                new ExcelCol(++excelColNum, "Munkavállaló: Emelet", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mv_Emelet"),
                new ExcelCol(++excelColNum, "Munkavállaló: Ajtó", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mv_Ajto"),
                new ExcelCol(++excelColNum, "Tartózkodás jogcíme", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Tartozkodas_jogcime"),
                new ExcelCol(++excelColNum, "Egészségbiztosítás", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Egeszsegbiztositas"),
                new ExcelCol(++excelColNum, "Visszautazási ország", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Visszautazasi_orszag"),
                new ExcelCol(++excelColNum, "Visszautazáskor közlekedési eszköz", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Visszaut_kozl_eszk"),
                new ExcelCol(++excelColNum, "Visszautazás - útlevél van-e", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Visszautazas_utlevel"),
                new ExcelCol(++excelColNum, "Érkezést megelőző ország", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Erkezest_meg_orszag"),
                new ExcelCol(++excelColNum, "Érkezést megelőző település", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Erkezest_meg_telepules"),
                new ExcelCol(++excelColNum, "Schengeni tartkózkodási okmány van-e", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Schengeni_tart_eng"),
                new ExcelCol(++excelColNum, "Elutasított tartózkodási kérelem", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Elut_tart_kerelem"),
                new ExcelCol(++excelColNum, "Büntetett előélet", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Buntetett_eloelet"),
                new ExcelCol(++excelColNum, "Kiutasították-e korábban", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Kiutasitottak_e"),
                new ExcelCol(++excelColNum, "Szenved-e gyógykezelésre szoruló betegségekben", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Szenv_gyogyk_sz_betegseg"),
                new ExcelCol(++excelColNum, "Kiskorú gyermek vele utazik-e", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Kiskoru_gyermek"),
                new ExcelCol(++excelColNum, "Okmány átvétele", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Okmany_atvetele"),
                new ExcelCol(++excelColNum, "Postai kézbesítés címe:", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Postai_kezb_cime"),
                new ExcelCol(++excelColNum, "Email cím", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Email"),
                new ExcelCol(++excelColNum, "Telefonszám", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Telefonszam"),
                new ExcelCol(++excelColNum, "Benyújtó", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Benyujto"),
                new ExcelCol(++excelColNum, "Okmány átvétel külképviseleten?", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Okmany_atv_kulkepviselet"),
                new ExcelCol(++excelColNum, "Átvételi ország", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Atveteli_orszag"),
                new ExcelCol(++excelColNum, "Átvételi település", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Atveteli_telepules"),
                new ExcelCol(++excelColNum, "Munkáltató rövid cégnév", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Munk_rovid_cegnev"),
                new ExcelCol(++excelColNum, "Munkáltató irányítószám", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Munk_Iranyitoszam"),
                new ExcelCol(++excelColNum, "Munkáltató település", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Munk_Telepules"),
                new ExcelCol(++excelColNum, "Munkáltató közterület neve", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Munk_kozt_neve"),
                new ExcelCol(++excelColNum, "Munkáltató közterület jellege", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Munk_kozt_jellege"),
                new ExcelCol(++excelColNum, "Munkáltató házszám/hrsz", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Munk_hazszam"),
                new ExcelCol(++excelColNum, "TEÁOR szám", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "TEAOR_szam"),
                new ExcelCol(++excelColNum, "KSH-szám", ExcelColTypeNum.Text, ExcelColRoleNum.Regex, @"\d\d\d\d\d\d\d\d \d\d\d\d \d\d\d [012]\d", ExcelColRequiredNum.Yes, "KSH_szam"),
                new ExcelCol(++excelColNum, "Munkáltató adószáma/adóazonosító jele", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Munk_Adoszam"),
                new ExcelCol(++excelColNum, "A foglalkoztatás munkaerő-kölcsönzés keretében történik", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Munkaero_kolcsonzes"),
                new ExcelCol(++excelColNum, "Munkakörhöz szükséges iskolai végzettség", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Munkakor_szuks_isk_vegz"),
                new ExcelCol(++excelColNum, "Szakképzettsége", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Szakkepzettsege"),
                new ExcelCol(++excelColNum, "Munkavégzés helye", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Mvegz_helye"),
                new ExcelCol(++excelColNum, "Munkavégzési irányítószám", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mvegz_iranyitoszam"),
                new ExcelCol(++excelColNum, "Munkavégzési település", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mvegz_telepules"),
                new ExcelCol(++excelColNum, "Munkavégzési közterület neve", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mvegz_kozt_neve"),
                new ExcelCol(++excelColNum, "Munkavégzési közterület jellege", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mvegz_kozt_jellege"),
                new ExcelCol(++excelColNum, "Munkavégzési házszám/hrsz", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mvegz_hazszam"),
                new ExcelCol(++excelColNum, "Munkavégzési Épület", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mvegz_epulet"),
                new ExcelCol(++excelColNum, "Munkavégzési Lépcsőház", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mvegz_lepcsohaz"),
                new ExcelCol(++excelColNum, "Munkavégzési Emelet", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mvegz_emelet"),
                new ExcelCol(++excelColNum, "Munkavégzési ajtó", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Mvegz_ajto"),
                new ExcelCol(++excelColNum, "Foglalkoztatóval kötött megállapodás kelte", ExcelColTypeNum.Date, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Fogl_megall_kelte"),
                new ExcelCol(++excelColNum, "Anyanyelve", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Anyanyelve"),
                new ExcelCol(++excelColNum, "Magyar nyelvismeret", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Magyar_nyelvismeret"),
                new ExcelCol(++excelColNum, "Dolgozott-e korábban Magyarországon?", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.No, "Dolgozott_Magyarorszagon"),
                // nem kellenek a forrás excelben, csak a resultban
                // new ExcelCol(++excelColNum, "Feldolgozottsági Állapot", ExcelColTypeNum.Text, ExcelColRoleNum.CreateIfNoExists, null, ExcelColRequiredNum.No, ""),
                //new ExcelCol(++excelColNum, "Fájl Feltöltés Státusz", ExcelColTypeNum.None, ExcelColRoleNum.CreateIfNoExists, null, ExcelColRequiredNum.No, "")
            };

        #endregion

        #region Public függvények

        /// <summary>
        /// Az oldalon lévő, a flowhoz szükséges dropdown elemek értékeit  betölti a  lementett txt fájlokból
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool DEADCODE__LoadDropdownValuesFromTextFiles(string path)
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
        /// Az oldalon lévő, a flowhoz szükséges dropdown elemek értékeit betölti SQL-bők
        /// </summary>
        /// <param name="sqlManager"></param>
        /// <returns></returns>
        public static bool LoadDropdownValuesFromSQL(MSSQLManager sqlManager)
        {


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

            // Excel megnyitása sikeres?
            if (isOk)
            {
                List<string> sheetNames = ExcelManager.WorksheetNames();

                // munkalapok feldolgozása
                foreach (string sheetName in sheetNames)
                {
                    ExcelSheetHeaderValidator(sheetName);
                }

                ExcelManager.CloseExcel();
            }

            return isOk;
        }

        /// <summary>
        /// Excel Sheet Header Validator
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static bool ExcelSheetHeaderValidator(string sheetName)
        {
            // megadott munkalap beolvasása
            ExcelManager.SelectWorksheetByName(sheetName);
            bool isHeaderOk = true;
            System.Data.DataTable dt = ExcelManager.WorksheetToDataTable(ExcelManager.ExcelSheet, true);

            // oszlopok meglétének ellenőrzése
            foreach (ExcelCol fejlec in excelHeaders.OrderByDescending(x => x.ExcelColNum))
            {
                // nem létezik
                if (!dt.Columns.Contains(fejlec.ExcelColName))
                {
                    ExcelManager.InsertFirstColumn(fejlec.ExcelColName);

                    if (!fejlec.ExcelColRole.Equals(ExcelColRoleNum.CreateIfNoExists))
                    {
                        ExcelManager.SetCellColor("A1", System.Drawing.Color.LightCoral);
                        isHeaderOk = false;
                    }
                    else
                    {
                        if (isHeaderOk)
                        {
                            ExcelManager.SetCellColor("A1", System.Drawing.Color.Khaki);
                        }
                    }
                }
            }

            if (!isHeaderOk)
            {
                ExcelManager.ExcelSheet.Rows[1].Insert();
                ExcelManager.SetCellValue("A1", "Hibás excel: hiányzó oszlopok. A hiányzó oszlopok világos korall színű fejléccel be lettek szúrva.");
                ExcelManager.SetRangeColor("A1", "E1", System.Drawing.Color.Red);
            }

            return isHeaderOk;
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
            string checkStatus = "";
            int rowNum = 1;

            // összes sor 
            foreach (DataRow currentRow in dt.Rows)
            {
                isRowOk = true;
                rowNum++;
                checkStatus = ExcelManager.GetDataRowValue(currentRow, "Ellenőrzés Státusz");

                // nem ellenőrzött sor?
                if (String.IsNullOrEmpty(checkStatus))
                {
                    // Nőtlen v. hajadon -> Nőtlen/hajadon
                    isRowOk = isRowOk && CsaladiAllapot(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

                    // kötelező szöveges oszlopok ellenőrzése
                    isRowOk = isRowOk && AllRequiredFieldChecker(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

                    // Dátum átalakítás és ellenőrzés
                    isRowOk = isRowOk && AllDateCheckAndConvert(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

                    // legördülő értékek ellenőrzése

                    // Ellenőrzés státusz állítása
                    checkStatus = isRowOk ? "OK" : "Hibás";
                    ExcelManager.SetCellValue(checkStatuscellName + rowNum.ToString(), checkStatus);
                }
                else
                {
                    isRowOk = checkStatus.ToLower().Equals("ok");
                }

                isGoodRow = isGoodRow || isRowOk; // van legalább egy jó sor
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
                string cellValue = ExcelManager.GetDataRowValue(currentRow, col.ExcelColName).ToLower();
                string cellName = fieldList[col.ExcelColName] + rowNum.ToString();

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
            foreach (ExcelCol col in excelHeaders.Where(x => x.ExcelColType == ExcelColTypeNum.Date))
            {
                DateTime dateTime = DateTime.MinValue;
                string cellValue = ExcelManager.GetDataRowValue(currentRow, col.ExcelColName).ToLower();
                string cellName = fieldList[col.ExcelColName] + rowNum.ToString();

                // van érték?
                if (cellValue.Length > 0)
                {
                    isCellValueOk = DateTime.TryParse(cellValue, out dateTime);

                    // dátum érték?
                    if (isCellValueOk)
                    {
                        // múltbélinek kell lennie?
                        if (col.ExcelColRole.Equals(ExcelColRoleNum.PastDate))
                        {
                            isCellValueOk = dateTime < DateTime.Today;
                        }

                        // jövőbélinek kell lennie?
                        if (col.ExcelColRole.Equals(ExcelColRoleNum.FutureDate))
                        {
                            isCellValueOk = dateTime > DateTime.Today;
                        }
                    }
                }
                else
                {
                    // lehet üres?
                    isCellValueOk = col.ExcelColRequired == ExcelColRequiredNum.No;
                }
            }

            return isCellValueOk;
        }

        #endregion
    }
}
