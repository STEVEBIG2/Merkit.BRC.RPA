﻿using Merkit.RPA.PA.Framework;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

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

        //  ** dropdown oszlopok kigyűjtése kódlista készítéshez
        public static Dictionary<string, List<string>> dropDownValuesbyType = new Dictionary<string, List<string>>();
        public static Dictionary<string, Dictionary<string, int>> dropDownIDsbyType = new Dictionary<string, Dictionary<string, int>>();

        public static ExcelManager excelManager = new ExcelManager();
        public static int excelColNum = 0;

        public static List<ExcelCol> excelHeaders = new List<ExcelCol>() {
                // new ExcelCol(++excelColNum, "Ügyszám", ExcelColTypeNum.Text, ExcelColRoleNum.CreateIfNoExists, null, ExcelColRequiredNum.No, "Ugyszam"),
                new ExcelCol(++excelColNum, "Beadható-e", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Beadhato"),
                new ExcelCol(++excelColNum, "Ellenőrzés Státusz", ExcelColTypeNum.None, ExcelColRoleNum.CreateIfNoExists, null, ExcelColRequiredNum.No, ""),
                new ExcelCol(++excelColNum, "Ügyintéző", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, ""), // EnterHungaryLoginId
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
                new ExcelCol(++excelColNum, "Visszautazás - útlevél van-e", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Visszautazas_utlevel"),
                new ExcelCol(++excelColNum, "Érkezést megelőző ország", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Erkezest_meg_orszag"),
                new ExcelCol(++excelColNum, "Érkezést megelőző település", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Erkezest_meg_telepules"),
                new ExcelCol(++excelColNum, "Schengeni tartkózkodási okmány van-e", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Schengeni_tart_eng"),
                new ExcelCol(++excelColNum, "Elutasított tartózkodási kérelem", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Elut_tart_kerelem"),
                new ExcelCol(++excelColNum, "Büntetett előélet", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Buntetett_eloelet"),
                new ExcelCol(++excelColNum, "Kiutasították-e korábban", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Kiutasitottak_e"),
                new ExcelCol(++excelColNum, "Szenved-e gyógykezelésre szoruló betegségekben", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Szenv_gyogyk_sz_betegseg"),
                new ExcelCol(++excelColNum, "Kiskorú gyermek vele utazik-e", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Kiskoru_gyermek"),
                new ExcelCol(++excelColNum, "Okmány átvétele", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Okmany_atvetele"),
                new ExcelCol(++excelColNum, "Postai kézbesítés címe:", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Postai_kezb_cime"),
                new ExcelCol(++excelColNum, "Email cím", ExcelColTypeNum.Text, ExcelColRoleNum.Regex, @"^[\w-\.]+@([\w-]+\.)+[\w-]+$", ExcelColRequiredNum.Yes, "Email"),
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
                new ExcelCol(++excelColNum, "KSH-szám", ExcelColTypeNum.Text, ExcelColRoleNum.Regex, @"^\d{14}\d[012]\d$", ExcelColRequiredNum.Yes, "KSH_szam"),
                new ExcelCol(++excelColNum, "Munkáltató adószáma/adóazonosító jele", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Munk_Adoszam"),
                new ExcelCol(++excelColNum, "A foglalkoztatás munkaerő-kölcsönzés keretében történik", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Munkaero_kolcsonzes"),
                new ExcelCol(++excelColNum, "Munkakörhöz szükséges iskolai végzettség", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Munkakor_szuks_isk_vegz"),
                new ExcelCol(++excelColNum, "Szakképzettsége", ExcelColTypeNum.Text, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Szakkepzettsege"),
                new ExcelCol(++excelColNum, "Munkavégzés helye", ExcelColTypeNum.Dropdown, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Mvegz_helye"),
                new ExcelCol(++excelColNum, "Munkavégzési irányítószám", ExcelColTypeNum.Text, ExcelColRoleNum.ZipCode, null, ExcelColRequiredNum.No, "Mvegz_iranyitoszam"),
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
                new ExcelCol(++excelColNum, "Dolgozott-e korábban Magyarországon?", ExcelColTypeNum.YesNo, ExcelColRoleNum.None, null, ExcelColRequiredNum.Yes, "Dolgozott_Magyarorszagon"),
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
        /// Excel Workbook Validator
        /// </summary>
        /// <param name="excelFileName"></param>
        /// <param name="excelFileId"></param>
        /// <param name="sqlManager"></param>
        /// <param name="tr"></param>
        /// <returns></returns>
        public static bool ExcelWorkbookValidator(string excelFileName, int excelFileId, MSSQLManager sqlManager)
        {
            Framework.Logger(0, "ExcelHeaderValidator", "Info", "", "-", String.Format("{0} file ellenőrzése elkezdődött.", excelFileName));
            Dictionary<string, int> excelSheets = new Dictionary<string, int>();
            string sqlString = "";
            bool headerOk = false;
            object retValue = 0;
            int excelSheetId = 0;
            bool isOk = excelManager.OpenExcel(excelFileName);

            // Excel megnyitása sikeres?
            if (isOk)
            {
                sqlString = String.Format(String.Format("UPDATE ExcelFiles SET QStatusId={0}, QStatusTime=getdate() WHERE ExcelFileId={1}", (int)QStatusNum.CheckingInProgress, excelFileId));
                sqlManager.ExecuteNonQuery(sqlString);

                List<string> sheetNames = excelManager.WorksheetNames();

                // munkalapok fejléceinek ellenőrzése
                foreach (string sheetName in sheetNames.Where(x => !x.Contains("Referen") && !x.Contains("Error")))
                {
                    // le lett már ellenőrizve a fejléc?
                    sqlString = String.Format("SELECT COUNT(*) FROM ExcelSheets WHERE ExcelFileId={0} AND ExcelSheetName='{1}'", excelFileId, sheetName);
                    retValue = sqlManager.ExecuteScalar(sqlString);

                    // ha nem volt még ellenőrizve, aktuális munkalap fejléceinek ellenőrzése
                    if (Convert.ToInt32(retValue) == 0)
                    {
                        headerOk = ExcelSheetHeaderValidator(sheetName);
                        excelSheetId = Dispatcher.InsertExcelSheetProc(excelFileId, sheetName, headerOk ? (int)QStatusNum.New : (int)QStatusNum.CheckedFailed, sqlManager);
                    }                          
                }

                // összes munkalap sorainak ellenőrzése
                isOk = ExcelAllSheetRowsValidator(excelFileId, sqlManager); 
                sqlString = String.Format("UPDATE ExcelFiles SET QStatusId={0}, QStatusTime=getdate() WHERE ExcelFileId={1}", (int)QStatusNum.CheckedOk, excelFileId);
                sqlManager.ExecuteNonQuery(sqlString);
                excelManager.CloseExcelWithoutSave();
                Framework.Logger(0, "ExcelHeaderValidator", "Info", "", "-", String.Format("{0} file ellenőrzése sikeresen befejeződött.", excelFileName));
            }
            else
            {
                sqlString = String.Format("UPDATE ExcelFiles SET QStatusId={0}, QStatusTime=getdate() WHERE ExcelFileId={1}", (int)QStatusNum.CheckedFailed, excelFileId);
                sqlManager.ExecuteNonQuery(sqlString);
                Framework.Logger(0, "ExcelHeaderValidator", "Err", "", "-", String.Format("{0} file ellenőrzése sikertelen volt.", excelFileName));
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
            Framework.Logger(0, "ExcelSheetHeaderValidator", "Info", "", "-", String.Format("A(z) {0} munkalap fejléc ellenőrzése elkezdődött.", sheetName));
            // megadott munkalap beolvasása
            excelManager.SelectWorksheetByName(sheetName);
            bool isHeaderOk = true;
            System.Data.DataTable dt = excelManager.WorksheetToDataTable(excelManager.ExcelSheet, true);

            // oszlopok meglétének ellenőrzése
            foreach (ExcelCol fejlec in excelHeaders.OrderByDescending(x => x.ExcelColNum))
            {
                // nem létezik
                if (!dt.Columns.Contains(fejlec.ExcelColName))
                {

                    if (!fejlec.ExcelColRole.Equals(ExcelColRoleNum.CreateIfNoExists))
                    {
                        excelManager.InsertFirstColumn(fejlec.ExcelColName);
                        Framework.Logger(0, "ExcelSheetHeaderValidator", "Err", "", "-", String.Format("Hiányzó oszlop a(z) {0} munkalapon : {1}", sheetName, fejlec.ExcelColName));
                        excelManager.SetCellColor("A1", System.Drawing.Color.LightCoral);
                        isHeaderOk = false;
                    }
                    else
                    {
                        if(isHeaderOk)
                        {
                            excelManager.InsertFirstColumn(fejlec.ExcelColName);
                            excelManager.SetCellColor("A1", System.Drawing.Color.Khaki);
                            excelManager.AutoFit();
                        }
                    }
                }
            }

            if (!isHeaderOk)
            {
                excelManager.ExcelSheet.Rows[1].Insert();
                excelManager.SetCellValue("A1", "Hibás excel: hiányzó oszlopok. A hiányzó oszlopok világos korall színű fejléccel be lettek szúrva.");
                excelManager.SetRangeColor("A1", "E1", System.Drawing.Color.Red);
                excelManager.SaveExcel();
            }

            Framework.Logger(0, "A(z) ExcelSheetHeaderValidator", "Info", "", "-", String.Format("{0} munkalap fejléc ellenőrzése befejeződött.", sheetName));
            return isHeaderOk;
        }

        /// <summary>
        /// Excel All Sheet Rows Validator
        /// </summary>
        /// <param name="excelFileId"></param>
        /// <param name="sqlManager"></param>
        /// <returns></returns>
        public static bool ExcelAllSheetRowsValidator(int excelFileId, MSSQLManager sqlManager)
        {
            bool isOk = true;
            int excelSheetId = 0;   

            // munkalapok sorainak ellenőrzése
            string excelSheetName = "";
            string sqlString = String.Format("SELECT * FROM ExcelSheets WHERE ExcelFileId={0} AND QStatusId IN ({1}, {2}) ORDER BY ExcelSheetId", excelFileId, (int)QStatusNum.New, (int)QStatusNum.CheckingInProgress);
            System.Data.DataTable dtExcelSheets = sqlManager.ExecuteQuery(sqlString);

            // részben vagy teljesen ellenőrizetlen munkalapok
            foreach (DataRow sheetRow in dtExcelSheets.Rows)
            {
                excelSheetId = Convert.ToInt32(sheetRow["ExcelSheetId"]);
                excelSheetName = sheetRow["ExcelSheetName"].ToString();
                ExcelSheetRowsValidator(excelFileId, excelSheetId, excelSheetName, sqlManager);
            }

            return isOk;
        }

        /// <summary>
        /// Excel Rows Validator
        /// </summary>
        /// <param name="excelFileId"></param>
        /// <param name="excelSheetId"></param>
        /// <param name="sheetName"></param>
        /// <param name="sqlManager"></param>
        /// <param name="tr"></param>
        /// <returns></returns>
        public static bool ExcelSheetRowsValidator(int excelFileId, int excelSheetId, string sheetName, MSSQLManager sqlManager)
        {
            bool isRowOk = true;
            bool isGoodRow = false;
            string checkStatus = "";
            int rowNum = 1;
            Framework.Logger(0, "ExcelSheetRowsValidator", "Info", "", "-", String.Format("A(z) {0} munkalap sorainak ellenőrzése elkezdődött.", sheetName));

            // megadott munkalap beolvasása
            excelManager.SelectWorksheetByName(sheetName);
            Range headerRange = excelManager.ReadEntireRow("A1");

            System.Data.DataTable dt = excelManager.WorksheetToDataTable(excelManager.ExcelSheet);
            Dictionary<string, string> dictExcelColumnNameToExcellCol = excelManager.GetExcelColumnNamesByDataTable(dt);
            string checkStatuscellName = dictExcelColumnNameToExcellCol["Ellenőrzés Státusz"];
            Dispatcher.LoadDropdownValuesFromSQL(sqlManager, dt);
            Dispatcher.LoadZipCodeValuesFromSQL(sqlManager, dt);

            SqlTransaction tr = sqlManager.BeginTransaction();

            try
            {
                // összes sor ellenörző ciklus
                foreach (DataRow currentRow in dt.Rows)
                {
                    isRowOk = true;
                    rowNum++;
                    checkStatus = excelManager.GetDataRowValue(currentRow, "Ellenőrzés Státusz");

                    // nem ellenőrzött sor?
                    if (String.IsNullOrEmpty(checkStatus))
                    {
                        isRowOk = ExcelSheetCurrentRowValidator(currentRow, sheetName, headerRange, rowNum, ref dictExcelColumnNameToExcellCol);
                        // Ellenőrzés státusz állítása
                        checkStatus = isRowOk ? "OK" : "Hibás";
                        excelManager.SetCellValue(checkStatuscellName + rowNum.ToString(), checkStatus);
                        
                        if (isRowOk) // SQL-be írás kell?
                        {
                            string ugyintezoValue = excelManager.GetDataRowValue(currentRow, "Ügyintéző").ToLower();
                            int excelRowId = Dispatcher.InsertExcelRowProc(excelFileId, excelSheetId, rowNum, ugyintezoValue, currentRow, sqlManager, tr);
                        }
                    }
                    else
                    {
                        isRowOk = checkStatus.ToLower().Equals("ok");
                    }

                    isGoodRow = isGoodRow || isRowOk; // van legalább egy jó sor
                }

                Framework.Logger(0, "ExcelSheetRowsValidator", "Info", "", "-", String.Format("A(z) {0} munkalap sorainak ellenőrzése befejeződött.", sheetName));
                tr.Commit();
                excelManager.SelectFirstWorksheetByIndex();
                excelManager.SaveExcel();
            }
            catch (Exception ex )
            {
                tr.Rollback();
                throw new Exception(ex.Message);
            }
            return isGoodRow;
        }

        /// <summary>
        /// Excel Sheet Current Row Validator
        /// </summary>
        /// <param name="currentRow"></param>
        /// <param name="sheetName"></param>
        /// <param name="rowNum"></param>
        /// <param name="dictExcelColumnNameToExcellCol"></param>
        /// <returns></returns>
        public static bool ExcelSheetCurrentRowValidator(DataRow currentRow, string sheetName, Range headerRange, int rowNum, ref Dictionary<string, string> dictExcelColumnNameToExcellCol)
        {
            bool isRowOk = true;
            bool isAdminOk = true;
            Range dest;
            string errorSheet = "";

            // kötelező szöveges oszlopok ellenőrzése
            isRowOk = isRowOk & AllRequiredFieldChecker(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

            // kötelező igen/nem oszlopok ellenőrzése
            isRowOk = isRowOk & AllBoolFieldChecker(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

            // regex szöveges oszlopok ellenőrzése
            isRowOk = isRowOk & AllRegexFieldChecker(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

            // Dátum átalakítás és ellenőrzés
            isRowOk = isRowOk & AllDateCheckAndConvert(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

            // link értékek ellenőrzése
            isRowOk = isRowOk & AllLinkCheck(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

            // ügyintéző ellenőrzése
            isAdminOk = AdministratorChecker(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);
            isRowOk = isRowOk & isAdminOk;

            // irányítószámok ellenőrzése
            isRowOk = isRowOk & AllZipCodeCheck(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

            // Nőtlen v. hajadon -> Nőtlen/hajadon
            isRowOk = isRowOk & CsaladiAllapot(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

            // legördülő értékek ellenőrzése
            isRowOk = isRowOk & AllDropdownCheck(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

            // egyéb üzleti szabályok ellenőrzése
            isRowOk = isRowOk & AllExtraBusinessRuleCheck(currentRow, rowNum, ref dictExcelColumnNameToExcellCol);

            // hibás sor?
            if (!isRowOk)
            {
                errorSheet = String.Format("Error - {0}", !isAdminOk ? Config.NotifyEmail.ToLower() : excelManager.GetDataRowValue(currentRow, "Ügyintéző").ToLower());
                Range copyRange = excelManager.ReadEntireRow("A" + rowNum.ToString());

                // létezik a munkalap?
                if (excelManager.AddNewSheetIfNotExist(errorSheet))
                {
                    dest = excelManager.GetCellRange("A1");
                    headerRange.Copy(dest);
                }

                int lastRow = excelManager.LastRow() + 1;
                dest = excelManager.GetCellRange("A" + lastRow.ToString());
                copyRange.Copy(dest);
                excelManager.AutoFit();
                excelManager.SelectWorksheetByName(sheetName);
            }

            return isRowOk;
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

            // ** címek összefüggései (A "Munkavállaló: Házszám" és "Munkavállaló: HRSZ" közül az egyik kötelező adat, és mindkettő egyszerre nem lehet kitöltve)
            cellValue = excelManager.GetDataRowValue(currentRow, "Munkavállaló: Házszám").ToLower();
            string cellValue2 = excelManager.GetDataRowValue(currentRow, "Munkavállaló: HRSZ").ToLower();
            isCellaValueOk = String.IsNullOrEmpty(cellValue) != String.IsNullOrEmpty(cellValue2); 

            if (!isCellaValueOk)
            {
                isRowOk = false;
                cellName = fieldList["Munkavállaló: Házszám"] + rowNum.ToString();
                excelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
                cellName = fieldList["Munkavállaló: HRSZ"] + rowNum.ToString();
                excelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
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
            string cellValue = excelManager.GetDataRowValue(currentRow, colName).ToLower();
            string cellName = fieldList[colName] + rowNum.ToString();

            if (cellValue.Equals("nőtlen") || cellValue.Equals("hajadon"))
            {
                excelManager.SetCellValue(cellName, "Nőtlen/hajadon");
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
            string cellValue = excelManager.GetDataRowValue(currentRow, colName).ToLower();
            string cellName = fieldList[colName] + rowNum.ToString();
            bool isCellValueOk = cellValue.Length>0 && Dispatcher.enterHungaryLogins.ContainsKey(cellValue);
            
            // ügyintéző létezik?
            if (!isCellValueOk)
            {
                excelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
            }

            return isCellValueOk;
        }

        /// <summary>
        /// Check All Required Text Fields
        /// </summary>
        /// <param name="currentRow"></param>
        /// <param name="rowNum"></param>
        /// <param name="fieldList"></param>
        /// <returns></returns>
        private static bool AllRequiredFieldChecker(DataRow currentRow, int rowNum, ref Dictionary<string, string> fieldList)
        {
            bool isRowValueOk = true;

            // dátum oszlopokon végigmenni
            foreach (ExcelCol col in excelHeaders.Where(x => x.ExcelColRequired == ExcelColRequiredNum.Yes && x.ExcelColRole != ExcelColRoleNum.Regex && x.ExcelColType != ExcelColTypeNum.YesNo))
            {
                string cellValue = excelManager.GetDataRowValue(currentRow, col.ExcelColName).ToLower();
                string cellName = fieldList[col.ExcelColName] + rowNum.ToString();

                if (cellValue.Length == 0)
                {
                    isRowValueOk = false;
                    excelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
                }

            }

            return isRowValueOk;
        }

        /// <summary>
        /// Check All Boolean Fields
        /// </summary>
        /// <param name="currentRow"></param>
        /// <param name="rowNum"></param>
        /// <param name="fieldList"></param>
        /// <returns></returns>
        private static bool AllBoolFieldChecker(DataRow currentRow, int rowNum, ref Dictionary<string, string> fieldList)
        {
            bool isRowValueOk = true;
            string[] yesNoValues = { "igen", "nem", "yes", "no", "true", "false" };

            // dátum oszlopokon végigmenni
            foreach (ExcelCol col in excelHeaders.Where(x => x.ExcelColRequired == ExcelColRequiredNum.Yes && x.ExcelColType == ExcelColTypeNum.YesNo))
            {
                string cellValue = excelManager.GetDataRowValue(currentRow, col.ExcelColName).ToLower();
                string cellName = fieldList[col.ExcelColName] + rowNum.ToString();

                if (cellValue.Length == 0 || ! yesNoValues.Contains(cellValue))
                {
                    isRowValueOk = false;
                    excelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
                }

            }

            return isRowValueOk;
        }

        /// <summary>
        /// Check All Regex Text Fields
        /// </summary>
        /// <param name="currentRow"></param>
        /// <param name="rowNum"></param>
        /// <param name="fieldList"></param>
        /// <param name="datumHeaderek"></param>
        private static bool AllRegexFieldChecker(DataRow currentRow, int rowNum, ref Dictionary<string, string> fieldList)
        {
            bool isRowValueOk = true;
            bool isCellValueOk = true;

            // regex oszlopokon végigmenni
            foreach (ExcelCol col in excelHeaders.Where(x => x.ExcelColRole == ExcelColRoleNum.Regex))
            {
                string cellValue = excelManager.GetDataRowValue(currentRow, col.ExcelColName).ToLower();
                string cellName = fieldList[col.ExcelColName] + rowNum.ToString();

                if (cellValue.Length > 0 || col.ExcelColRequired == ExcelColRequiredNum.Yes)
                {
                    isCellValueOk = Regex.IsMatch(cellValue.Replace(" ", ""), col.ExcelColRoleExpression); // ^\d{14}\d[012]\d$

                    if (!isCellValueOk)
                    {
                        isRowValueOk = false;
                        excelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
                    }
                }

            }

            return isRowValueOk;
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
                cellValue = excelManager.GetDataRowValue(currentRow, col.ExcelColName).ToLower();
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
                    excelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
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
        /// <returns></returns>
        private static bool AllDropdownCheck(DataRow currentRow, int rowNum, ref Dictionary<string, string> fieldList)
        {
            string cellValue = "";
            string cellName = "";
            bool isCellValuesOk = true;

            // dropdown oszlopokon végigmenni
            foreach (ExcelCol col in excelHeaders.Where(x => x.ExcelColType == ExcelColTypeNum.Dropdown))
            {
                cellValue = excelManager.GetDataRowValue(currentRow, col.ExcelColName).Trim().ToLower();

                // nem lehet üres vagy van érték?
                if (! String.IsNullOrEmpty(cellValue) || col.ExcelColRequired.Equals(ExcelColRequiredNum.Yes))
                {
                    cellName = fieldList[col.ExcelColName] + rowNum.ToString();
                    //var temp = dropDownIDsbyType[col.ExcelColName];

                    // létező érték?
                    if (! dropDownIDsbyType[col.ExcelColName].ContainsKey(cellValue))
                    {
                        excelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
                        isCellValuesOk = false;
                    }
                }

            }

            return isCellValuesOk;
        }

        /// <summary>
        /// Check All Dropdown Fields
        /// </summary>
        /// <param name="currentRow"></param>
        /// <param name="rowNum"></param>
        /// <param name="fieldList"></param>
        /// <returns></returns>
        private static bool AllZipCodeCheck(DataRow currentRow, int rowNum, ref Dictionary<string, string> fieldList)
        {
            string cellValue = "";
            string cellName = "";
            DateTime dateTime = DateTime.MinValue;
            bool isCellValuesOk = true;

            // irányítószám oszlopokon végigmenni
            foreach (ExcelCol col in excelHeaders.Where(x => x.ExcelColRole == ExcelColRoleNum.ZipCode))
            {
                cellValue = excelManager.GetDataRowValue(currentRow, col.ExcelColName).Trim().ToLower();

                // nem lehet üres vagy van érték?
                if (!String.IsNullOrEmpty(cellValue) || col.ExcelColRequired.Equals(ExcelColRequiredNum.Yes))
                {
                    cellName = fieldList[col.ExcelColName] + rowNum.ToString();
                    //var temp = dropDownIDsbyType[col.ExcelColName];

                    // nem létező érték?
                    if (!Dispatcher.zipCodes.Contains(cellValue))
                    {
                        excelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
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
                cellValue = excelManager.GetDataRowValue(currentRow, col.ExcelColName).Trim().ToLower();

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
                        excelManager.SetCellColor(cellName, System.Drawing.Color.LightCoral);
                        isCellValuesOk = false;
                    }
                }
            }

            return isCellValuesOk;
        }

        #endregion
    }
}
