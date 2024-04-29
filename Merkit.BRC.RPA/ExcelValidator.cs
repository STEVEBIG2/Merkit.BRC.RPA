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
using System.Data.SqlClient;
using System.Security.Policy;
using System.Text.RegularExpressions;

namespace Merkit.BRC.RPA
{
    public enum QStatusNum
    {
        Locked = -1,
        New = 0,
        InProgress = 1,
        Failed = 2,
        SuccessFullExcel = 3,
        SuccessFullRow = 4,
        Exported = 5,
        Deleted = 6
    };

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
        Link = 5,
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

        public static Dictionary<string, string> enterHungaryLogins = new Dictionary<string, string>(); // ügyintézők

        //  ** dropdown oszlopok kigyűjtése kódlista készítéshez
        public static Dictionary<string, List<string>> dropDownValuesbyType = new Dictionary<string, List<string>>();
        public static Dictionary<string, Dictionary<string, int>> dropDownIDsbyType = new Dictionary<string, Dictionary<string, int>>();

        public static int excelColNum = 0;

        public static List<ExcelCol> excelHeaders = new List<ExcelCol>() {
                // new ExcelCol(++excelColNum, "Ügyszám", ExcelColTypeNum.Text, ExcelColRoleNum.CreateIfNoExists, null, ExcelColRequiredNum.No, "Ugyszam"),
                new ExcelCol(++excelColNum, "Ellenőrzés Státusz", ExcelColTypeNum.None, ExcelColRoleNum.CreateIfNoExists, null, ExcelColRequiredNum.No, ""),
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
                new ExcelCol(++excelColNum, "Személy: Állampolgárság", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Allampolgarsag"),
                new ExcelCol(++excelColNum, "Személy: Családi állapot", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Sz_Csaladi_allapot"),
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
                new ExcelCol(++excelColNum, "Munkavállaló: Irányítószám", ExcelColTypeNum.Text, ExcelColRoleNum.ZipCode, null, ExcelColRequiredNum.Yes, "Mv_Iranyitoszam"),
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
                new ExcelCol(++excelColNum, "Munkáltató irányítószám", ExcelColTypeNum.Text, ExcelColRoleNum.ZipCode, null, ExcelColRequiredNum.Yes, "Munk_Iranyitoszam"),
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
                new ExcelCol(++excelColNum, "Munkavégzési irányítószám", ExcelColTypeNum.Dropdown, ExcelColRoleNum.ZipCode, null, ExcelColRequiredNum.No, "Mvegz_iranyitoszam"),
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
                //
                new ExcelCol(++excelColNum, "Érvényes útlevél teljes másolata", ExcelColTypeNum.Text, ExcelColRoleNum.Link, null, ExcelColRequiredNum.Yes, "Utlevel_link"),
                new ExcelCol(++excelColNum, "Arckép", ExcelColTypeNum.Text, ExcelColRoleNum.Link, null, ExcelColRequiredNum.Yes, "Arckep_link"),
                new ExcelCol(++excelColNum, "Lakásbérleti jogviszonyt igazoló lakásbérleti szerződés", ExcelColTypeNum.Text, ExcelColRoleNum.Link, null, ExcelColRequiredNum.Yes, "Lakasberlet_link"),
                new ExcelCol(++excelColNum, "Lakás tulajdonjogát igazoló okirat", ExcelColTypeNum.Text, ExcelColRoleNum.Link, null, ExcelColRequiredNum.Yes, "Lakas_tulajdonjog_link"),
                new ExcelCol(++excelColNum, "A foglalkoztatási jogviszony létesítésére irányuló előzetes megállapodás", ExcelColTypeNum.Text, ExcelColRoleNum.Link, null, ExcelColRequiredNum.Yes, "Elozetes_megallapodas_link"),
                new ExcelCol(++excelColNum, "Céges meghatalmazás", ExcelColTypeNum.Text, ExcelColRoleNum.Link, null, ExcelColRequiredNum.Yes, "Ceges_megh_link"),
                new ExcelCol(++excelColNum, "Szálláshely bejelentő lap", ExcelColTypeNum.Text, ExcelColRoleNum.Link, null, ExcelColRequiredNum.Yes, "Szallashely_bej_link"),
                new ExcelCol(++excelColNum, "Postázási kérelem", ExcelColTypeNum.Text, ExcelColRoleNum.Link, null, ExcelColRequiredNum.Yes, "Postazasi_kerelem_link"),
                new ExcelCol(++excelColNum, "Vízumfelvételi nyilatkozat", ExcelColTypeNum.Text, ExcelColRoleNum.Link, null, ExcelColRequiredNum.Yes, "Vizumfelv_ny_link"),
                new ExcelCol(++excelColNum, "Kölcsönzési szerződés", ExcelColTypeNum.Text, ExcelColRoleNum.Link, null, ExcelColRequiredNum.Yes, "Kolcs_szerz_link")
                // new ExcelCol(++excelColNum, "Munkavállaló: Azonosító", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Mv_Azonosito"),
                // nem kellenek a forrás excelben, csak a resultban
                // new ExcelCol(++excelColNum, "Feldolgozottsági Állapot", ExcelColTypeNum.Text, ExcelColRoleNum.CreateIfNoExists, null, ExcelColRequiredNum.No, ""),
                //new ExcelCol(++excelColNum, "Fájl Feltöltés Státusz", ExcelColTypeNum.None, ExcelColRoleNum.CreateIfNoExists, null, ExcelColRequiredNum.No, "")
            };

        #endregion

        #region Public függvények

        /// <summary>
        /// Main Process
        /// </summary>
        /// <returns></returns>
        public static bool MainProcess(string excelFileName)
        {
            bool processOk = false;            
            MSSQLManager sqlManager = new MSSQLManager();
            sqlManager.ConnectByConfig();

            try
            {
                // ügyintéző login adatok begyűjtése
                GetEnterHungaryLogins(sqlManager);

                // *** dropdown lista ellenőrzéshez előkészülés
                foreach (ExcelCol col in ExcelValidator.excelHeaders.Where(x => x.ExcelColType == ExcelColTypeNum.Dropdown))
                {
                    dropDownValuesbyType.Add(col.ExcelColName, new List<string>());
                }

                // kész
                processOk = ExcelWorkbookValidator(excelFileName, sqlManager);
            }
            catch (Exception ex)
            {
                Framework.Logger(0, "MainProcess", "Err", "", "-", String.Format("MainProcess hiba: {0}", ex.Message));
                throw;
            }
            finally
            {
                sqlManager.Disconnect();
            }

            return processOk;
        }

        /// <summary>
        /// Excel Workbook Validator
        /// </summary>
        /// <param name="excelFileName"></param>
        /// <returns></returns>
        public static bool ExcelWorkbookValidator(string excelFileName, MSSQLManager sqlManager)
        {
            Framework.Logger(0, "ExcelHeaderValidator", "Info", "", "-", String.Format("{0} file ellenőrzése elkezdődött.", excelFileName));
            Dictionary<string, bool> excelSheetHeaderChecking = new Dictionary<string, bool>();
            bool isOk = ExcelManager.OpenExcel(excelFileName);

            // Excel megnyitása sikeres?
            if (isOk)
            {
                List<string> sheetNames = ExcelManager.WorksheetNames();

                // munkalapok fejléceinek ellenőrzése
                foreach (string sheetName in sheetNames)
                {
                    // EH adatokat tartalmazhat?
                    if (!sheetName.Contains("Referen"))
                    {
                        excelSheetHeaderChecking.Add(sheetName, ExcelSheetHeaderValidator(sheetName, sqlManager));
                    }
                }

                // munkalapok sorainak ellenőrzése
                foreach(KeyValuePair<string, bool> goodSheetNameItem in excelSheetHeaderChecking.Where(x => x.Value))
                {
                    ExcelSheetRowsValidator(goodSheetNameItem.Key, sqlManager);
                }

                ExcelManager.CloseExcel();
                Framework.Logger(0, "ExcelHeaderValidator", "Info", "", "-", String.Format("{0} file ellenőrzése sikeresen befejeződött.", excelFileName));
            }
            else
            {
                Framework.Logger(0, "ExcelHeaderValidator", "Err", "", "-", String.Format("{0} file ellenőrzése sikertelen volt.", excelFileName));
            }

            return isOk;
        }

        /// <summary>
        /// Excel Sheet Header Validator
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static bool ExcelSheetHeaderValidator(string sheetName, MSSQLManager sqlManager)
        {
            Framework.Logger(0, "ExcelSheetHeaderValidator", "Info", "", "-", String.Format("A(z) {0} munkalap fejléc ellenőrzése elkezdődött.", sheetName));
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

                    if (!fejlec.ExcelColRole.Equals(ExcelColRoleNum.CreateIfNoExists))
                    {
                        ExcelManager.InsertFirstColumn(fejlec.ExcelColName);
                        Framework.Logger(0, "ExcelSheetHeaderValidator", "Err", "", "-", String.Format("Hiányzó oszlop a(z) {0} munkalapon : {1}", sheetName, fejlec.ExcelColName));
                        ExcelManager.SetCellColor("A1", System.Drawing.Color.LightCoral);
                        isHeaderOk = false;
                    }
                    else
                    {
                        if(isHeaderOk)
                        {
                            ExcelManager.InsertFirstColumn(fejlec.ExcelColName);
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

            Framework.Logger(0, "A(z) ExcelSheetHeaderValidator", "Info", "", "-", String.Format("{0} munkalap fejléc ellenőrzése befejeződött.", sheetName));
            return isHeaderOk;
        }

        /// <summary>
        /// Excel Rows Validator
        /// </summary>
        /// <param name="excelFileName"></param>
        /// <returns></returns>
        public static bool ExcelSheetRowsValidator(string sheetName, MSSQLManager sqlManager)
        {
            Framework.Logger(0, "ExcelSheetRowsValidator", "Info", "", "-", String.Format("A(z) {0} munkalap sorainak ellenőrzése elkezdődött.", sheetName));

            // megadott munkalap beolvasása
            ExcelManager.SelectWorksheetByName(sheetName);
            System.Data.DataTable dt = ExcelManager.WorksheetToDataTable(ExcelManager.ExcelSheet);
            Dictionary<string, string> dictExcelColumnNameToExcellCol = ExcelManager.GetExcelColumnNamesByDataTable(dt);
            string checkStatuscellName = dictExcelColumnNameToExcellCol["Ellenőrzés Státusz"];

            LoadDropdownValuesFromSQL(sqlManager, dt);

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
                    isRowOk = isRowOk & CsaladiAllapot(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

                    // ügyintéző ellenőrzése
                    isRowOk = isRowOk & AdministratorChecker(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

                    // kötelező szöveges oszlopok ellenőrzése
                    isRowOk = isRowOk & AllRequiredFieldChecker(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

                    // Dátum átalakítás és ellenőrzés
                    isRowOk = isRowOk & AllDateCheckAndConvert(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

                    // legördülő értékek ellenőrzése
                    isRowOk = isRowOk & AllDropdownCheck(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

                    // link értékek ellenőrzése
                    isRowOk = isRowOk & AllLinkCheck(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

                    // egyéb üzleti szabályok ellenőrzése
                    isRowOk = isRowOk & AllExtraBusinessRuleCheck(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

                    // Ellenőrzés státusz állítása
                    checkStatus = isRowOk ? "OK" : "Hibás";
                    ExcelManager.SetCellValue(checkStatuscellName + rowNum.ToString(), checkStatus);
                    //  var x = ExcelSheet.get_Range("C2").Style;
                }
                else
                {
                    isRowOk = checkStatus.ToLower().Equals("ok");
                }

                isGoodRow = isGoodRow || isRowOk; // van legalább egy jó sor
            }

            Framework.Logger(0, "ExcelSheetRowsValidator", "Info", "", "-", String.Format("A(z) {0} munkalap sorainak ellenőrzése befejeződött .", sheetName));
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

        /// <summary>
        /// Az oldalon lévő, a flowhoz szükséges dropdown elemek értékeit betölti SQL-ből
        /// </summary>
        /// <param name="sqlManager"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static bool LoadDropdownValuesFromSQL(MSSQLManager sqlManager, System.Data.DataTable dt)
        {
            string cellValue = "";

            // dropdown oszlopok típusa alapján kódlista készítése
            foreach (DataRow excelRow in dt.Rows)
            {
                foreach (string colName in dropDownValuesbyType.Keys)
                {
                    cellValue = excelRow[colName].ToString();

                    if (!dropDownValuesbyType[colName].Contains(cellValue))
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

                if (array != null && array.Length > 0)
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
                        dropDownIDsbyType[colName].Add(dr["DropDownValue"].ToString().ToLower(), Int32.Parse(dr["DropDownsValueId"].ToString()));
                    }
                }

            }

            return true;
        }

        #endregion

        #region Private függvények (oszloponkénti ellenőrzések)

        /// <summary>
        /// GetEnterHungaryLogins
        /// </summary>
        /// <param name="sqlManager"></param>
        /// <returns></returns>
        private static void GetEnterHungaryLogins(MSSQLManager sqlManager)
        {
            System.Data.DataTable dt = sqlManager.ExecuteQuery("SELECT Email,PasswordText FROM EnterHungaryLogins WHERE Deleted=0");

            foreach (DataRow row in dt.Rows)
            {
                enterHungaryLogins.Add(row[0].ToString().ToLower(), row[1].ToString());
            }

            return;
        }

        /// <summary>
        /// All Extra Business Rule Check
        /// </summary>
        /// <param name="currentRow"></param>
        /// <param name="rowNum"></param>
        /// <param name="fieldList"></param>
        /// <returns></returns>
        private static bool AllExtraBusinessRuleCheck(DataRow currentRow, int rowNum, ref Dictionary<string, string> fieldList)
        {
            string cellValue = "";
            string cellName = "";
            bool isCellaValueOk = true;
            bool isRowOk = true;

            // ** irányítószám (1011-9999)
            //cellValue = ExcelManager.GetDataRowValue(currentRow, "col.ExcelColName").ToLower();
            //isCellaValueOk = Regex.IsMatch(cellValue, @"^[0-9]{4}$");
            //isCellaValueOk = isCellaValueOk && Convert.ToInt32(isCellaValueOk) >= 1011;
            //
            //if (!isCellaValueOk)
            //{
            //    isRowOk = false;
            //    cellName = fieldList["col.ExcelColName"] + rowNum.ToString();
            //    ExcelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
            //}

            // ** KSH szám ("\d\d\d\d\d\d\d\d \d\d\d\d \d\d\d [012]\d"  /12345678 1234 123 12/
            cellValue = ExcelManager.GetDataRowValue(currentRow, "KSH-szám").ToLower();
            isCellaValueOk = Regex.IsMatch(cellValue, @"^\d\d\d\d\d\d\d\d \d\d\d\d \d\d\d [012]\d$");

            if (!isCellaValueOk)
            {
                isRowOk = false;
                cellName = fieldList["KSH-szám"] + rowNum.ToString();
                ExcelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
            }

            // ** címek összefüggései (A "Munkavállaló: Házszám" és "Munkavállaló: HRSZ" közül az egyik kötelező adat, és mindkettő egyszerre nem lehet kitöltve)
            cellValue = ExcelManager.GetDataRowValue(currentRow, "Munkavállaló: Házszám").ToLower();
            string cellValue2 = ExcelManager.GetDataRowValue(currentRow, "Munkavállaló: HRSZ").ToLower();
            isCellaValueOk = String.IsNullOrEmpty(cellValue) != String.IsNullOrEmpty(cellValue2); 

            if (!isCellaValueOk)
            {
                isRowOk = false;
                cellName = fieldList["Munkavállaló: Házszám"] + rowNum.ToString();
                ExcelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
                cellName = fieldList["Munkavállaló: HRSZ"] + rowNum.ToString();
                ExcelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
            }

            return isRowOk;
        }

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
        private static bool AdministratorChecker(DataRow currentRow, int rowNum, ref Dictionary<string, string> fieldList)
        {
            string colName = "Ügyintéző";
            string cellValue = ExcelManager.GetDataRowValue(currentRow, colName).ToLower();
            string cellName = fieldList[colName] + rowNum.ToString();
            bool isCellValueOk = enterHungaryLogins.ContainsKey(cellValue);
            
            // ügyintéző létezik?
            if (!isCellValueOk)
            {
                ExcelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
            }

            return isCellValueOk;
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
            string cellValue = "";
            string cellName = "";
            DateTime dateTime = DateTime.MinValue;
            bool isCellValueOk = true;

            // dátum oszlopokon végigmenni
            foreach (ExcelCol col in excelHeaders.Where(x => x.ExcelColType == ExcelColTypeNum.Date))
            {
                dateTime = DateTime.MinValue;
                cellValue = ExcelManager.GetDataRowValue(currentRow, col.ExcelColName).ToLower();
                cellValue = cellValue.Length > 10 ? cellValue.Replace(" ", "").Substring(0, 10) : cellValue;
                cellName = fieldList[col.ExcelColName] + rowNum.ToString();

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

                // hibás dátum?
                if (! isCellValueOk)
                {
                    ExcelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
                }
            }

            return isCellValueOk;
        }

        /// <summary>
        /// Check All Dropdown Fields
        /// </summary>
        /// <param name="currentRow"></param>
        /// <param name="rowNum"></param>
        /// <param name="fieldList"></param>
        /// <param name="datumHeaderek"></param>
        private static bool AllDropdownCheck(DataRow currentRow, int rowNum, ref Dictionary<string, string> fieldList)
        {
            string cellValue = "";
            string cellName = "";
            DateTime dateTime = DateTime.MinValue;
            bool isCellValuesOk = true;

            // dátum oszlopokon végigmenni
            foreach (ExcelCol col in excelHeaders.Where(x => x.ExcelColType == ExcelColTypeNum.Dropdown))
            {
                cellValue = ExcelManager.GetDataRowValue(currentRow, col.ExcelColName).Trim().ToLower();

                // nem lehet üres vagy van érték?
                if (! String.IsNullOrEmpty(cellValue) || col.ExcelColRequired.Equals(ExcelColRequiredNum.Yes))
                {
                    cellName = fieldList[col.ExcelColName] + rowNum.ToString();
                    //var temp = dropDownIDsbyType[col.ExcelColName];

                    // létező érték?
                    if (! dropDownIDsbyType[col.ExcelColName].ContainsKey(cellValue))
                    {
                        ExcelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
                        isCellValuesOk = false;
                    }
                }

            }

            return isCellValuesOk;
        }

        /// <summary>
        /// Check All Link Fields
        /// </summary>
        /// <param name="currentRow"></param>
        /// <param name="rowNum"></param>
        /// <param name="fieldList"></param>
        /// <returns></returns>
        private static bool AllLinkCheck(DataRow currentRow, int rowNum, ref Dictionary<string, string> fieldList)
        {
            string cellValue = "";
            string cellName = "";
            bool isCellValuesOk = true;
            bool isGoodLink = true;

            // dátum oszlopokon végigmenni
            foreach (ExcelCol col in excelHeaders.Where(x => x.ExcelColType == ExcelColTypeNum.Link))
            {
                cellValue = ExcelManager.GetDataRowValue(currentRow, col.ExcelColName).Trim().ToLower();

                // üres érték?
                if (String.IsNullOrEmpty(cellValue))
                {
                    isGoodLink = col.ExcelColRequired.Equals(ExcelColRequiredNum.Yes);
                }
                else
                {
                    // link?
                    if (Framework.IsValidURL(cellValue))
                    {
                        cellName = fieldList[col.ExcelColName] + rowNum.ToString();
                        ExcelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
                        isCellValuesOk = false;
                    }
                }
            }

            return isCellValuesOk;
        }

        #endregion
    }
}
