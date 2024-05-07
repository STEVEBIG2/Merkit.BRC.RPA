-- Use BRC_Hungary_Test
-- GO

-- DROP TABLE ExcelRows
-- GO
-- DROP TABLE ExcelFiles
-- GO
-- DROP TABLE QStatuses
-- GO
-- Drop TABLE EnterHungaryLogins
-- GO

select @@VERSION
GO

CREATE TABLE EnterHungaryLogins
(
	EnterHungaryLoginId INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
	Email VARCHAR(100) NOT NULL,
	PasswordText VARCHAR(100) NOT NULL,
	Deleted INT NOT NULL DEFAULT 0
)
GO

CREATE UNIQUE INDEX IX1_EnterHungaryLogins On EnterHungaryLogins(Email)
GO

CREATE INDEX IX2_EnterHungaryLogins On EnterHungaryLogins(Deleted)
GO

-- ZipCodes
-- DROP TABLE ZipCodes
-- GO

CREATE TABLE ZipCodes
(
	ZipCodeId INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
	ZipCode VARCHAR(10) NOT NULL,
	Deleted INT NOT NULL DEFAULT 0
)
GO

CREATE UNIQUE INDEX IX1_ZipCodes On ZipCodes(ZipCode)
GO

CREATE UNIQUE INDEX IX2_ZipCodes On ZipCodes(ZipCode, Deleted)
GO

-- DropDownTypes, DropDownsValues
-- DROP TABLE ExcelRows
-- GO
-- DROP VIEW View_DropDowns
-- go
-- Drop TABLE DropDownsValues
-- go
-- Drop TABLE DropDownTypes
-- go

CREATE TABLE DropDownTypes
(
	DropDownTypeId INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
	DropDownName VARCHAR(50) NOT NULL,
	ExcelColNames VARCHAR(150) NOT NULL,
	Deleted INT NOT NULL DEFAULT 0
)
GO

CREATE UNIQUE INDEX IX1_DropDownTypes On DropDownTypes(DropDownName)
GO

CREATE INDEX IX2_DropDownTypes On DropDownTypes(ExcelColNames)
GO

CREATE INDEX IX3_DropDownTypes On DropDownTypes(Deleted)
GO

CREATE TABLE DropDownsValues
(
	DropDownsValueId INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
	DropDownTypeId INT NOT NULL,
	DropDownValue VARCHAR(100) NOT NULL,
	Deleted INT NOT NULL DEFAULT 0
)
GO

CREATE UNIQUE INDEX IX1_DropDownsValues On DropDownsValues(DropDownTypeId, DropDownValue)
GO

CREATE INDEX IX2_DropDownsValues On DropDownsValues(DropDownTypeId, DropDownValue, Deleted)
GO

CREATE INDEX IX3_DropDownsValues On DropDownsValues(DropDownTypeId)
GO

ALTER TABLE DropDownsValues  WITH CHECK ADD  CONSTRAINT FK_DropDownsValues_DropDownTypes FOREIGN KEY(DropDownTypeId) REFERENCES DropDownTypes (DropDownTypeId)
GO

ALTER TABLE DropDownsValues CHECK CONSTRAINT FK_DropDownsValues_DropDownTypes
GO

CREATE VIEW View_DropDowns AS
SELECT dt.DropDownTypeId, dt.DropDownName, dt.ExcelColNames, dv.DropDownsValueId, dv.DropDownValue
  FROM DropDownsValues dv INNER JOIN DropDownTypes dt ON (dv.DropDownTypeId=dt.DropDownTypeId AND dt.Deleted=0)
  WHERE dv.Deleted=0
GO

--

CREATE TABLE QStatuses (
	QStatusId int NOT NULL,
	QStatusName nvarchar(50) NOT NULL
) 
GO

ALTER TABLE QStatuses ADD PRIMARY KEY (QStatusId)
GO

INSERT INTO QStatuses (QStatusId, QStatusName) VALUES (-1, 'Locked')
GO
INSERT INTO QStatuses (QStatusId, QStatusName) VALUES (0, 'New')
GO
INSERT INTO QStatuses (QStatusId, QStatusName) VALUES (1, 'In Progress')
GO
INSERT INTO QStatuses (QStatusId, QStatusName) VALUES (2, 'Failed')
GO
INSERT INTO QStatuses (QStatusId, QStatusName) VALUES (3, 'SuccessFullExcel')
GO
INSERT INTO QStatuses (QStatusId, QStatusName) VALUES (4, 'SuccessFullRow')
GO
INSERT INTO QStatuses (QStatusId, QStatusName) VALUES (5, 'Exported')
GO
INSERT INTO QStatuses (QStatusId, QStatusName) VALUES (6, 'Deleted')
GO

-- ** ExcelFiles, ExcelSheets, ExcelRows
-- DROP VIEW View_ExcelRowsByExcelColNames
-- go
-- DROP VIEW View_ExcelRows
-- GO
-- DROP TABLE ExcelRows
-- GO
-- DROP TABLE ExcelSheets
-- GO
-- DROP TABLE ExcelFiles 
-- GO

CREATE TABLE ExcelFiles (
	ExcelFileId int IDENTITY(1,1) NOT NULL ,
	ExcelFileName varchar(50) NOT NULL,
	QStatusId int NULL,
	QStatusTime datetime NULL,
	RobotName varchar(50) NULL,
 CONSTRAINT PK_ExcelFiles PRIMARY KEY NONCLUSTERED 
(
	ExcelFileId ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE UNIQUE INDEX IX1_ExcelFiles On ExcelFiles(ExcelFileName)
GO

CREATE INDEX IX2_ExcelFiles On ExcelFiles(QStatusId)
GO

ALTER TABLE ExcelFiles  WITH CHECK ADD  CONSTRAINT FK_ExcelFiles_QStatuses FOREIGN KEY(QStatusId) REFERENCES QStatuses (QStatusId)
GO

ALTER TABLE ExcelFiles CHECK CONSTRAINT FK_ExcelFiles_QStatuses
GO

---

CREATE TABLE ExcelSheets(
	ExcelSheetId int IDENTITY(1,1) NOT NULL,
	ExcelSheetName varchar(50) NOT NULL,
	ExcelFileId int NOT NULL,
	QStatusId int NULL,
	QStatusTime datetime NULL,
 CONSTRAINT PK_ExcelSheets PRIMARY KEY NONCLUSTERED 
(
	ExcelSheetId ASC
))
GO

CREATE INDEX IX1_ExcelSheets On ExcelSheets(ExcelFileId)
GO

ALTER TABLE ExcelSheets  WITH CHECK ADD  CONSTRAINT FK_ExcelSheets_ExcelFiles FOREIGN KEY(ExcelFileId) REFERENCES ExcelFiles (ExcelFileId)
GO

ALTER TABLE ExcelSheets CHECK CONSTRAINT FK_ExcelSheets_ExcelFiles
GO

CREATE INDEX IX2_ExcelSheets On ExcelSheets(QStatusId)
GO

ALTER TABLE ExcelSheets  WITH CHECK ADD  CONSTRAINT FK_ExcelSheets_QStatuses FOREIGN KEY(QStatusId) REFERENCES QStatuses (QStatusId)
GO

ALTER TABLE ExcelSheets CHECK CONSTRAINT FK_ExcelSheets_QStatuses
GO

---

CREATE TABLE ExcelRows(
	ExcelRowId int IDENTITY(1,1) NOT NULL,
	ExcelFileId int NOT NULL,
	ExcelSheetId int NOT NULL,
	ExcelRowNum int NOT NULL,
    Ugyszam VARCHAR(150),
	EnterHungaryLoginId INT NOT NULL,
	Beadhato BIT NOT NULL,
    Sz_Szul_Vezeteknev VARCHAR(150) NOT NULL,
    Sz_Szul_Keresztnev VARCHAR(150) NOT NULL,
    Sz_Utlevel_Szig VARCHAR(150) NOT NULL,
    Mv_Munkakor VARCHAR(150),
    Mv_FEOR INT NOT NULL,
    Sz_Vezeteknev VARCHAR(150) NOT NULL,
    Sz_Keresztnev VARCHAR(150) NOT NULL,
    Sz_Szul_Orszag INT NOT NULL,
    Sz_Szul_Hely VARCHAR(150) NOT NULL,
    Sz_Szul_Datum DATE NOT NULL,
    Sz_Anyja_Vezeteknev VARCHAR(150) NOT NULL,
    Sz_Anyja_Keresztnev VARCHAR(150) NOT NULL,
    Sz_Neme INT NOT NULL,
    Sz_Allampolgarsag INT NOT NULL,
    Sz_Csaladi_allapot INT NOT NULL,
    Sz_Magy_erk_meg_fogl VARCHAR(150),
    Sz_Utlevel_kiall_helye VARCHAR(150) NOT NULL,
    Sz_Utlevel_kiall_datuma DATE NOT NULL,
    Sz_Utlevel_lejarat_datuma DATE NOT NULL,
    Sz_Varhato_jovedelem VARCHAR(150) NOT NULL,
    Sz_Varhato_jov_penznem INT NOT NULL,
    Sz_Tart_eng_erv DATE NOT NULL,
    Dijmentes BIT NOT NULL,
    Engedely_hosszabbitas BIT NOT NULL,
    Utlevel_tipusa INT NOT NULL,
    Iskolai_vegzettsege INT NOT NULL,
    Mv_Iranyitoszam VARCHAR(10) NOT NULL,
    Mv_Telepules VARCHAR(150) NOT NULL,
    Mv_Kozterulet_neve VARCHAR(150) NOT NULL,
    Mv_Kozterulet_jellege INT NOT NULL,
    Mv_Hazszam VARCHAR(150),
    Mv_HRSZ VARCHAR(150),
    Mv_Epulet VARCHAR(150),
    Mv_Lepcsohaz VARCHAR(150),
    Mv_Emelet INT,
    Mv_Ajto VARCHAR(150),
    Tartozkodas_jogcime INT NOT NULL,
    Egeszsegbiztositas INT NOT NULL,
    Visszautazasi_orszag INT NOT NULL,
    Visszaut_kozl_eszk VARCHAR(150),
    Visszautazas_utlevel BIT NOT NULL,
    Erkezest_meg_orszag INT NOT NULL,
    Erkezest_meg_telepules VARCHAR(150) NOT NULL,
    Schengeni_tart_eng BIT NOT NULL,
    Elut_tart_kerelem BIT NOT NULL,
    Buntetett_eloelet BIT NOT NULL,
    Kiutasitottak_e BIT NOT NULL,
    Szenv_gyogyk_sz_betegseg BIT NOT NULL,
    Kiskoru_gyermek BIT NOT NULL,
    Okmany_atvetele INT NOT NULL,
    Postai_kezb_cime INT NOT NULL,
    Email VARCHAR(150) NOT NULL,
    Telefonszam VARCHAR(150),
    Benyujto INT,
    Okmany_atv_kulkepviselet VARCHAR(150),
    Atveteli_orszag INT,
    Atveteli_telepules VARCHAR(150),
    Munk_rovid_cegnev VARCHAR(150) NOT NULL,
    Munk_Iranyitoszam VARCHAR(10) NOT NULL,
    Munk_Telepules VARCHAR(150) NOT NULL,
    Munk_kozt_neve VARCHAR(150) NOT NULL,
    Munk_kozt_jellege INT NOT NULL,
    Munk_hazszam VARCHAR(150) NOT NULL,
    TEAOR_szam INT NOT NULL,
    KSH_szam VARCHAR(150) NOT NULL,
    Munk_Adoszam VARCHAR(150) NOT NULL,
    Munkaero_kolcsonzes BIT NOT NULL,
    Munkakor_szuks_isk_vegz INT NOT NULL,
    Szakkepzettsege VARCHAR(150) NOT NULL,
    Mvegz_helye INT NOT NULL,
    Mvegz_iranyitoszam VARCHAR(10),
    Mvegz_telepules VARCHAR(150),
    Mvegz_kozt_neve VARCHAR(150),
    Mvegz_kozt_jellege INT,
    Mvegz_hazszam VARCHAR(150),
    Mvegz_epulet VARCHAR(150),
    Mvegz_lepcsohaz VARCHAR(150),
    Mvegz_emelet VARCHAR(150),
    Mvegz_ajto VARCHAR(150),
    Fogl_megall_kelte DATE NOT NULL,
    Anyanyelve INT NOT NULL,
    Magyar_nyelvismeret BIT NOT NULL,
    Dolgozott_Magyarorszagon BIT,
    Utlevel_link VARCHAR(150) NOT NULL,
    Arckep_link VARCHAR(150) NOT NULL,
    Lakasberlet_link VARCHAR(150) NOT NULL,
    Lakas_tulajdonjog_link VARCHAR(150) NOT NULL,
    Elozetes_megallapodas_link VARCHAR(150) NOT NULL,
    Ceges_megh_link VARCHAR(150) NOT NULL,
    Szallashely_bej_link VARCHAR(150) NOT NULL,
    Postazasi_kerelem_link VARCHAR(150) NOT NULL,
    Vizumfelv_ny_link VARCHAR(150) NOT NULL,
    Kolcs_szerz_link VARCHAR(150) NOT NULL,
	QStatusId int NULL,
	QStatusTime datetime NULL,
 CONSTRAINT PK_ExcelRows PRIMARY KEY NONCLUSTERED 
(
	ExcelRowId ASC
))
GO

CREATE INDEX IX1_ExcelRows On ExcelRows(ExcelFileId)
GO

CREATE INDEX IX2_ExcelRows On ExcelRows(ExcelSheetId)
GO

ALTER TABLE ExcelRows  WITH CHECK ADD  CONSTRAINT FK_ExcelRows_ExcelFiles FOREIGN KEY(ExcelFileId) REFERENCES ExcelFiles (ExcelFileId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_ExcelFiles
GO

ALTER TABLE ExcelRows  WITH CHECK ADD  CONSTRAINT FK_ExcelRows_ExcelSheets FOREIGN KEY(ExcelSheetId) REFERENCES ExcelSheets (ExcelSheetId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_ExcelSheets
GO 

CREATE INDEX IX3_ExcelRows On ExcelRows(QStatusId)
GO

ALTER TABLE ExcelRows  WITH CHECK ADD  CONSTRAINT FK_ExcelRows_QStatuses FOREIGN KEY(QStatusId) REFERENCES QStatuses (QStatusId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_QStatuses
GO

CREATE INDEX IX4_ExcelRows On ExcelRows(EnterHungaryLoginId)
GO

ALTER TABLE ExcelRows  WITH CHECK ADD  CONSTRAINT FK_ExcelRows_EnterHungaryLogins FOREIGN KEY(EnterHungaryLoginId) REFERENCES EnterHungaryLogins (EnterHungaryLoginId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_EnterHungaryLogins
GO

--
CREATE INDEX IX_ExcelRows_Mv_FEOR ON ExcelRows(Mv_FEOR)
GO

CREATE INDEX IX_ExcelRows_Sz_Szul_Orszag ON ExcelRows(Sz_Szul_Orszag)
GO

CREATE INDEX IX_ExcelRows_Sz_Neme ON ExcelRows(Sz_Neme)
GO

CREATE INDEX IX_ExcelRows_Sz_Allampolgarsag ON ExcelRows(Sz_Allampolgarsag)
GO

CREATE INDEX IX_ExcelRows_Sz_Csaladi_allapot ON ExcelRows(Sz_Csaladi_allapot)
GO

CREATE INDEX IX_ExcelRows_Sz_Varhato_jov_penznem ON ExcelRows(Sz_Varhato_jov_penznem)
GO

CREATE INDEX IX_ExcelRows_Utlevel_tipusa ON ExcelRows(Utlevel_tipusa)
GO

CREATE INDEX IX_ExcelRows_Iskolai_vegzettsege ON ExcelRows(Iskolai_vegzettsege)
GO

CREATE INDEX IX_ExcelRows_Mv_Kozterulet_jellege ON ExcelRows(Mv_Kozterulet_jellege)
GO

CREATE INDEX IX_ExcelRows_Mv_Emelet ON ExcelRows(Mv_Emelet)
GO

CREATE INDEX IX_ExcelRows_Tartozkodas_jogcime ON ExcelRows(Tartozkodas_jogcime)
GO

CREATE INDEX IX_ExcelRows_Egeszsegbiztositas ON ExcelRows(Egeszsegbiztositas)
GO

CREATE INDEX IX_ExcelRows_Visszautazasi_orszag ON ExcelRows(Visszautazasi_orszag)
GO

CREATE INDEX IX_ExcelRows_Erkezest_meg_orszag ON ExcelRows(Erkezest_meg_orszag)
GO

CREATE INDEX IX_ExcelRows_Okmany_atvetele ON ExcelRows(Okmany_atvetele)
GO

CREATE INDEX IX_ExcelRows_Postai_kezb_cime ON ExcelRows(Postai_kezb_cime)
GO

CREATE INDEX IX_ExcelRows_Benyujto ON ExcelRows(Benyujto)
GO

CREATE INDEX IX_ExcelRows_Atveteli_orszag ON ExcelRows(Atveteli_orszag)
GO

CREATE INDEX IX_ExcelRows_Munk_kozt_jellege ON ExcelRows(Munk_kozt_jellege)
GO

CREATE INDEX IX_ExcelRows_TEAOR_szam ON ExcelRows(TEAOR_szam)
GO

CREATE INDEX IX_ExcelRows_Munkakor_szuks_isk_vegz ON ExcelRows(Munkakor_szuks_isk_vegz)
GO

CREATE INDEX IX_ExcelRows_Mvegz_helye ON ExcelRows(Mvegz_helye)
GO

CREATE INDEX IX_ExcelRows_Mvegz_iranyitoszam ON ExcelRows(Mvegz_iranyitoszam)
GO

CREATE INDEX IX_ExcelRows_Mvegz_kozt_jellege ON ExcelRows(Mvegz_kozt_jellege)
GO

CREATE INDEX IX_ExcelRows_Anyanyelve ON ExcelRows(Anyanyelve)
GO

--
ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Mv_FEOR FOREIGN KEY(Mv_FEOR) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Mv_FEOR
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Sz_Szul_Orszag FOREIGN KEY(Sz_Szul_Orszag) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Sz_Szul_Orszag
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Sz_Neme FOREIGN KEY(Sz_Neme) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Sz_Neme
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Sz_Allampolgarsag FOREIGN KEY(Sz_Allampolgarsag) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Sz_Allampolgarsag
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Sz_Csaladi_allapot FOREIGN KEY(Sz_Csaladi_allapot) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Sz_Csaladi_allapot
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Sz_Varhato_jov_penznem FOREIGN KEY(Sz_Varhato_jov_penznem) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Sz_Varhato_jov_penznem
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Utlevel_tipusa FOREIGN KEY(Utlevel_tipusa) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Utlevel_tipusa
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Iskolai_vegzettsege FOREIGN KEY(Iskolai_vegzettsege) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Iskolai_vegzettsege
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Mv_Kozterulet_jellege FOREIGN KEY(Mv_Kozterulet_jellege) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Mv_Kozterulet_jellege
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Mv_Emelet FOREIGN KEY(Mv_Emelet) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Mv_Emelet
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Tartozkodas_jogcime FOREIGN KEY(Tartozkodas_jogcime) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Tartozkodas_jogcime
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Egeszsegbiztositas FOREIGN KEY(Egeszsegbiztositas) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Egeszsegbiztositas
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Visszautazasi_orszag FOREIGN KEY(Visszautazasi_orszag) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Visszautazasi_orszag
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Erkezest_meg_orszag FOREIGN KEY(Erkezest_meg_orszag) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Erkezest_meg_orszag
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Okmany_atvetele FOREIGN KEY(Okmany_atvetele) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Okmany_atvetele
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Postai_kezb_cime FOREIGN KEY(Postai_kezb_cime) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Postai_kezb_cime
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Benyujto FOREIGN KEY(Benyujto) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Benyujto
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Atveteli_orszag FOREIGN KEY(Atveteli_orszag) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Atveteli_orszag
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Munk_kozt_jellege FOREIGN KEY(Munk_kozt_jellege) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Munk_kozt_jellege
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_TEAOR_szam FOREIGN KEY(TEAOR_szam) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_TEAOR_szam
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Munkakor_szuks_isk_vegz FOREIGN KEY(Munkakor_szuks_isk_vegz) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Munkakor_szuks_isk_vegz
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Mvegz_helye FOREIGN KEY(Mvegz_helye) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Mvegz_helye
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Mvegz_iranyitoszam FOREIGN KEY(Mvegz_iranyitoszam) REFERENCES ZipCodes(ZipCode)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Mvegz_iranyitoszam
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Munk_Iranyitoszam FOREIGN KEY(Munk_Iranyitoszam) REFERENCES ZipCodes(ZipCode)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Munk_Iranyitoszam
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Mv_Iranyitoszam FOREIGN KEY(Mv_Iranyitoszam) REFERENCES ZipCodes(ZipCode)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Mv_Iranyitoszam
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Mvegz_kozt_jellege FOREIGN KEY(Mvegz_kozt_jellege) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Mvegz_kozt_jellege
GO

ALTER TABLE ExcelRows WITH CHECK ADD CONSTRAINT FK_ExcelRows_Anyanyelve FOREIGN KEY(Anyanyelve) REFERENCES DropDownsValues(DropDownsValueId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_Anyanyelve
GO

--
-- DROP VIEW View_ExcelRowsByExcelColNames
-- GO
-- DROP VIEW View_ExcelRows
-- GO

CREATE VIEW View_ExcelRowsByExcelColNames AS
SELECT
  r.ExcelRowId,
  r.ExcelFileId,
  r.ExcelSheetId,
  r.ExcelRowNum,
  r.Ugyszam AS [Ügyszám],
  (SELECT Email FROM EnterHungaryLogins Where EnterHungaryLoginId = r.EnterHungaryLoginId) AS [Ügyintéző],
  CASE WHEN Beadhato=1 THEN 'Igen' ELSE 'Nem' END AS [Beadható-e],
  r.Sz_Szul_Vezeteknev AS [Személy: Születési vezetéknév],
  r.Sz_Szul_Keresztnev AS [Személy: Születési keresztnév],
  r.Sz_Utlevel_Szig AS [Személy: Útlevél száma/Személy ig.],
  r.Mv_Munkakor AS [Munkavállaló: Munkakör megnevezése],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Mv_FEOR) AS [Munkavállaló: FEOR],
  r.Sz_Vezeteknev AS [Személy: Vezetéknév],
  r.Sz_Keresztnev AS [Személy: Keresztnév],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Sz_Szul_Orszag) AS [Személy: Születési ország],
  r.Sz_Szul_Hely AS [Személy: Születési hely],
  r.Sz_Szul_Datum AS [Személy: Születési dátum],
  r.Sz_Anyja_Vezeteknev AS [Személy: Anyja vezetékneve],
  r.Sz_Anyja_Keresztnev AS [Személy: Anyja keresztneve],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Sz_Neme) AS [Személy: Neme],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Sz_Allampolgarsag) AS [Személy: Állampolgárság],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Sz_Csaladi_allapot) AS [Személy: Családi állapot],
  r.Sz_Magy_erk_meg_fogl AS [Személy: Magyarországra érkezést megelőző foglalkozás],
  r.Sz_Utlevel_kiall_helye AS [Személy: Útlevél kiállításának helye],
  r.Sz_Utlevel_kiall_datuma AS [Személy: Útlevél kiállításának dátuma],
  r.Sz_Utlevel_lejarat_datuma AS [Személy: Útlevél lejáratának dátuma],
  r.Sz_Varhato_jovedelem AS [Személy: Várható jövedelem],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Sz_Varhato_jov_penznem) AS [Személy: Várható jövedelem pénznem],
  r.Sz_Tart_eng_erv AS [Személy: Tartózkodási engedély érvényessége],
  CASE WHEN Dijmentes=1 THEN 'Igen' ELSE 'Nem' END AS [Díjmentes-e],
  CASE WHEN Engedely_hosszabbitas=1 THEN 'Igen' ELSE 'Nem' END AS [Engedély hosszabbítás-e],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Utlevel_tipusa) AS [Útlevél típusa],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Iskolai_vegzettsege) AS [Iskolai végzettsége],
  r.Mv_Iranyitoszam AS [Munkavállaló: Irányítószám],
  r.Mv_Telepules AS [Munkavállaló: Település],
  r.Mv_Kozterulet_neve AS [Munkavállaló: Közterület neve],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Mv_Kozterulet_jellege) AS [Munkavállaló: Közterület jellege],
  r.Mv_Hazszam AS [Munkavállaló: Házszám],
  r.Mv_HRSZ AS [Munkavállaló: HRSZ],
  r.Mv_Epulet AS [Munkavállaló: Épület],
  r.Mv_Lepcsohaz AS [Munkavállaló: Lépcsőház],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Mv_Emelet) AS [Munkavállaló: Emelet],
  r.Mv_Ajto AS [Munkavállaló: Ajtó],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Tartozkodas_jogcime) AS [Tartózkodás jogcíme],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Egeszsegbiztositas) AS [Egészségbiztosítás],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Visszautazasi_orszag) AS [Visszautazási ország],
  r.Visszaut_kozl_eszk AS [Visszautazáskor közlekedési eszköz],
  CASE WHEN Visszautazas_utlevel=1 THEN 'Igen' ELSE 'Nem' END AS [Visszautazás - útlevél van-e],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Erkezest_meg_orszag) AS [Érkezést megelőző ország],
  r.Erkezest_meg_telepules AS [Érkezést megelőző település],
  CASE WHEN Schengeni_tart_eng=1 THEN 'Igen' ELSE 'Nem' END AS [Schengeni tartkózkodási okmány van-e],
  CASE WHEN Elut_tart_kerelem=1 THEN 'Igen' ELSE 'Nem' END AS [Elutasított tartózkodási kérelem],
  CASE WHEN Buntetett_eloelet=1 THEN 'Igen' ELSE 'Nem' END AS [Büntetett előélet],
  CASE WHEN Kiutasitottak_e=1 THEN 'Igen' ELSE 'Nem' END AS [Kiutasították-e korábban],
  CASE WHEN Szenv_gyogyk_sz_betegseg=1 THEN 'Igen' ELSE 'Nem' END AS [Szenved-e gyógykezelésre szoruló betegségekben],
  CASE WHEN Kiskoru_gyermek=1 THEN 'Igen' ELSE 'Nem' END AS [Kiskorú gyermek vele utazik-e],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Okmany_atvetele) AS [Okmány átvétele],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Postai_kezb_cime) AS [Postai kézbesítés címe:],
  r.Email AS [Email cím],
  r.Telefonszam AS [Telefonszám],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Benyujto) AS [Benyújtó],
  r.Okmany_atv_kulkepviselet AS [Okmány átvétel külképviseleten?],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Atveteli_orszag) AS [Átvételi ország],
  r.Atveteli_telepules AS [Átvételi település],
  r.Munk_rovid_cegnev AS [Munkáltató rövid cégnév],
  r.Munk_Iranyitoszam AS [Munkáltató irányítószám],
  r.Munk_Telepules AS [Munkáltató település],
  r.Munk_kozt_neve AS [Munkáltató közterület neve],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Munk_kozt_jellege) AS [Munkáltató közterület jellege],
  r.Munk_hazszam AS [Munkáltató házszám/hrsz],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.TEAOR_szam) AS [TEÁOR szám],
  r.KSH_szam AS [KSH-szám],
  r.Munk_Adoszam AS [Munkáltató adószáma/adóazonosító jele],
  CASE WHEN Munkaero_kolcsonzes=1 THEN 'Igen' ELSE 'Nem' END AS [A foglalkoztatás munkaerő-kölcsönzés keretében történik],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Munkakor_szuks_isk_vegz) AS [Munkakörhöz szükséges iskolai végzettség],
  r.Szakkepzettsege AS [Szakképzettsége],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Mvegz_helye) AS [Munkavégzés helye],
  r.Mvegz_iranyitoszam AS [Munkavégzési irányítószám],
  r.Mvegz_telepules AS [Munkavégzési település],
  r.Mvegz_kozt_neve AS [Munkavégzési közterület neve],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Mvegz_kozt_jellege) AS [Munkavégzési közterület jellege],
  r.Mvegz_hazszam AS [Munkavégzési házszám/hrsz],
  r.Mvegz_epulet AS [Munkavégzési Épület],
  r.Mvegz_lepcsohaz AS [Munkavégzési Lépcsőház],
  r.Mvegz_emelet AS [Munkavégzési Emelet],
  r.Mvegz_ajto AS [Munkavégzési ajtó],
  r.Fogl_megall_kelte AS [Foglalkoztatóval kötött megállapodás kelte],
  (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Anyanyelve) AS [Anyanyelve],
  CASE WHEN Magyar_nyelvismeret=1 THEN 'Igen' ELSE 'Nem' END AS [Magyar nyelvismeret],
  CASE WHEN Dolgozott_Magyarorszagon=1 THEN 'Igen' ELSE 'Nem' END AS [Dolgozott-e korábban Magyarországon?],
  r.Utlevel_link AS [Érvényes útlevél teljes másolata],
  r.Arckep_link AS [Arckép],
  r.Lakasberlet_link AS [Lakásbérleti jogviszonyt igazoló lakásbérleti szerződés],
  r.Lakas_tulajdonjog_link AS [Lakás tulajdonjogát igazoló okirat],
  r.Elozetes_megallapodas_link AS [A foglalkoztatási jogviszony létesítésére irányuló előzetes megállapodás],
  r.Ceges_megh_link AS [Céges meghatalmazás],
  r.Szallashely_bej_link AS [Szálláshely bejelentő lap],
  r.Postazasi_kerelem_link AS [Postázási kérelem],
  r.Vizumfelv_ny_link AS [Vízumfelvételi nyilatkozat],
  r.Kolcs_szerz_link AS [Kölcsönzési szerződés]
From ExcelRows r
GO

CREATE VIEW View_ExcelRows AS
 SELECT
   r.ExcelRowId,
   r.ExcelFileId,
   r.ExcelSheetId,
   r.ExcelRowNum,
   r.Ugyszam,
   (SELECT Email FROM EnterHungaryLogins Where EnterHungaryLoginId = r.EnterHungaryLoginId) AS [Ugyintezo_Email],
   r.Beadhato,
   r.Sz_Szul_Vezeteknev,
   r.Sz_Szul_Keresztnev,
   r.Sz_Utlevel_Szig,
   r.Mv_Munkakor,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Mv_FEOR) AS Mv_FEOR,
   r.Sz_Vezeteknev,
   r.Sz_Keresztnev,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Sz_Szul_Orszag) AS Sz_Szul_Orszag,
   r.Sz_Szul_Hely,
   r.Sz_Szul_Datum,
   r.Sz_Anyja_Vezeteknev,
   r.Sz_Anyja_Keresztnev,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Sz_Neme) AS Sz_Neme,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Sz_Allampolgarsag) AS Sz_Allampolgarsag,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Sz_Csaladi_allapot) AS Sz_Csaladi_allapot,
   r.Sz_Magy_erk_meg_fogl,
   r.Sz_Utlevel_kiall_helye,
   r.Sz_Utlevel_kiall_datuma,
   r.Sz_Utlevel_lejarat_datuma,
   r.Sz_Varhato_jovedelem,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Sz_Varhato_jov_penznem) AS Sz_Varhato_jov_penznem,
   r.Sz_Tart_eng_erv,
   r.Dijmentes,
   r.Engedely_hosszabbitas,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Utlevel_tipusa) AS Utlevel_tipusa,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Iskolai_vegzettsege) AS Iskolai_vegzettsege,
   r.Mv_Iranyitoszam,
   r.Mv_Telepules,
   r.Mv_Kozterulet_neve,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Mv_Kozterulet_jellege) AS Mv_Kozterulet_jellege,
   r.Mv_Hazszam,
   r.Mv_HRSZ,
   r.Mv_Epulet,
   r.Mv_Lepcsohaz,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Mv_Emelet) AS Mv_Emelet,
   r.Mv_Ajto,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Tartozkodas_jogcime) AS Tartozkodas_jogcime,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Egeszsegbiztositas) AS Egeszsegbiztositas,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Visszautazasi_orszag) AS Visszautazasi_orszag,
   r.Visszaut_kozl_eszk,
   r.Visszautazas_utlevel,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Erkezest_meg_orszag) AS Erkezest_meg_orszag,
   r.Erkezest_meg_telepules,
   r.Schengeni_tart_eng,
   r.Elut_tart_kerelem,
   r.Buntetett_eloelet,
   r.Kiutasitottak_e,
   r.Szenv_gyogyk_sz_betegseg,
   r.Kiskoru_gyermek,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Okmany_atvetele) AS Okmany_atvetele,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Postai_kezb_cime) AS Postai_kezb_cime,
   r.Email,
   r.Telefonszam,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Benyujto) AS Benyujto,
   r.Okmany_atv_kulkepviselet,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Atveteli_orszag) AS Atveteli_orszag,
   r.Atveteli_telepules,
   r.Munk_rovid_cegnev,
   r.Munk_Iranyitoszam,
   r.Munk_Telepules,
   r.Munk_kozt_neve,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Munk_kozt_jellege) AS Munk_kozt_jellege,
   r.Munk_hazszam,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.TEAOR_szam) AS TEAOR_szam,
   r.KSH_szam,
   r.Munk_Adoszam,
   r.Munkaero_kolcsonzes,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Munkakor_szuks_isk_vegz) AS Munkakor_szuks_isk_vegz,
   r.Szakkepzettsege,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Mvegz_helye) AS Mvegz_helye,
   r.Mvegz_iranyitoszam,
   r.Mvegz_telepules,
   r.Mvegz_kozt_neve,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Mvegz_kozt_jellege) AS Mvegz_kozt_jellege,
   r.Mvegz_hazszam,
   r.Mvegz_epulet,
   r.Mvegz_lepcsohaz,
   r.Mvegz_emelet,
   r.Mvegz_ajto,
   r.Fogl_megall_kelte,
   (SELECT DropDownValue FROM DropDownsValues Where DropDownsValueId = r.Anyanyelve) AS Anyanyelve,
   r.Magyar_nyelvismeret,
   r.Dolgozott_Magyarorszagon,
   r.Utlevel_link,
   r.Arckep_link,
   r.Lakasberlet_link,
   r.Lakas_tulajdonjog_link,
   r.Elozetes_megallapodas_link,
   r.Ceges_megh_link,
   r.Szallashely_bej_link,
   r.Postazasi_kerelem_link,
   r.Vizumfelv_ny_link,
   r.Kolcs_szerz_link
   From ExcelRows r
GO