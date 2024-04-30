USE BRC_Hungary_Test
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================  
-- Author: Steve
-- Create date: 2024.05.01
-- Description: Insert Excel row
-- ============================================= 
CREATE PROCEDURE [dbo].[InsertExcelRow]
  @ExcelFileId int,
  @ExcelSheetId int,
  @ExcelRowNum int,
  @EnterHungaryLoginId int,
  @Sz_Szul_Vezeteknev VARCHAR(150),
  @Sz_Szul_Keresztnev VARCHAR(150),
  @Sz_Utlevel_Szig VARCHAR(150),
  @Mv_Munkakor VARCHAR(150)=NULL,
  @Mv_FEOR INT,
  @Sz_Vezeteknev VARCHAR(150),
  @Sz_Keresztnev VARCHAR(150),
  @Sz_Szul_Orszag INT,
  @Sz_Szul_Hely VARCHAR(150),
  @Sz_Szul_Datum DATE,
  @Sz_Anyja_Vezeteknev VARCHAR(150),
  @Sz_Anyja_Keresztnev VARCHAR(150),
  @Sz_Neme INT,
  @Sz_Allampolgarsag INT,
  @Sz_Csaladi_allapot INT,
  @Sz_Magy_erk_meg_fogl VARCHAR(150)=NULL,
  @Sz_Utlevel_kiall_helye VARCHAR(150),
  @Sz_Utlevel_kiall_datuma DATE,
  @Sz_Utlevel_lejarat_datuma DATE,
  @Sz_Varhato_jovedelem VARCHAR(150),
  @Sz_Varhato_jov_penznem INT,
  @Sz_Tart_eng_erv DATE,
  @Dijmentes BIT,
  @Engedely_hosszabbitas BIT,
  @Utlevel_tipusa INT,
  @Iskolai_vegzettsege INT,
  @Mv_Iranyitoszam VARCHAR(10),
  @Mv_Telepules VARCHAR(150),
  @Mv_Kozterulet_neve VARCHAR(150),
  @Mv_Kozterulet_jellege INT,
  @Mv_Hazszam VARCHAR(150)=NULL,
  @Mv_HRSZ VARCHAR(150)=NULL,
  @Mv_Epulet VARCHAR(150)=NULL,
  @Mv_Lepcsohaz VARCHAR(150)=NULL,
  @Mv_Emelet INT=NULL,
  @Mv_Ajto VARCHAR(150)=NULL,
  @Tartozkodas_jogcime INT,
  @Egeszsegbiztositas INT,
  @Visszautazasi_orszag INT,
  @Visszaut_kozl_eszk VARCHAR(150)=NULL,
  @Visszautazas_utlevel INT,
  @Erkezest_meg_orszag INT,
  @Erkezest_meg_telepules VARCHAR(150),
  @Schengeni_tart_eng BIT=NULL,
  @Elut_tart_kerelem BIT=NULL,
  @Buntetett_eloelet BIT=NULL,
  @Kiutasitottak_e BIT=NULL,
  @Szenv_gyogyk_sz_betegseg BIT=NULL,
  @Kiskoru_gyermek BIT=NULL,
  @Okmany_atvetele INT,
  @Postai_kezb_cime INT,
  @Email VARCHAR(150),
  @Telefonszam VARCHAR(150)=NULL,
  @Benyujto INT=NULL,
  @Okmany_atv_kulkepviselet VARCHAR(150)=NULL,
  @Atveteli_orszag INT=NULL,
  @Atveteli_telepules VARCHAR(150)=NULL,
  @Munk_rovid_cegnev VARCHAR(150),
  @Munk_Iranyitoszam VARCHAR(10),
  @Munk_Telepules VARCHAR(150),
  @Munk_kozt_neve VARCHAR(150),
  @Munk_kozt_jellege INT,
  @Munk_hazszam VARCHAR(150),
  @TEAOR_szam INT,
  @KSH_szam VARCHAR(150),
  @Munk_Adoszam VARCHAR(150),
  @Munkaero_kolcsonzes VARCHAR(150),
  @Munkakor_szuks_isk_vegz INT,
  @Szakkepzettsege VARCHAR(150),
  @Mvegz_helye INT,
  @Mvegz_iranyitoszam VARCHAR(10)=NULL,
  @Mvegz_telepules VARCHAR(150)=NULL,
  @Mvegz_kozt_neve VARCHAR(150)=NULL,
  @Mvegz_kozt_jellege INT=NULL,
  @Mvegz_hazszam VARCHAR(150)=NULL,
  @Mvegz_epulet VARCHAR(150)=NULL,
  @Mvegz_lepcsohaz VARCHAR(150)=NULL,
  @Mvegz_emelet VARCHAR(150)=NULL,
  @Mvegz_ajto VARCHAR(150)=NULL,
  @Fogl_megall_kelte DATE,
  @Anyanyelve INT,
  @Magyar_nyelvismeret BIT,
  @Dolgozott_Magyarorszagon BIT=NULL,
  @Utlevel_link VARCHAR(150),
  @Arckep_link VARCHAR(150),
  @Lakasberlet_link VARCHAR(150),
  @Lakas_tulajdonjog_link VARCHAR(150),
  @Elozetes_megallapodas_link VARCHAR(150),
  @Ceges_megh_link VARCHAR(150),
  @Szallashely_bej_link VARCHAR(150),
  @Postazasi_kerelem_link VARCHAR(150),
  @Vizumfelv_ny_link VARCHAR(150),
  @Kolcs_szerz_link VARCHAR(150)
AS
BEGIN  
 SET NOCOUNT ON; 

 INSERT INTO ExcelRows (
      ExcelFileId,
      ExcelSheetId,
      ExcelRowNum, 
      EnterHungaryLoginId,

      QStatusId,
      QStatusTime
    ) VALUES (
      @ExcelFileId,
      @ExcelSheetId,
      @ExcelRowNum,
      @EnterHungaryLoginId,

      0,
      getdate()
   )

  --SELECT @@IDENTITY AS NewExcelRowId
  RETURN @@IDENTITY
END
GO