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

-- DropDownTypes, DropDownsValues
--DROP VIEW View_DropDowns
--go
--Drop TABLE DropDownsValues
--go
--Drop TABLE DropDownTypes
--go

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
	DropDownsValuesId INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
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
SELECT dt.DropDownTypeId, dt.DropDownName, dt.ExcelColNames, dv.DropDownsValuesId, dv.DropDownValue
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

-- ** ExcelFiles, ExcelRows. ExcelRows

CREATE TABLE ExcelFiles (
	ExcelFileId int IDENTITY(1,1) NOT NULL ,
	ExcelFileName varchar(50) NOT NULL,
	ExcelType varchar(20) NOT NULL,
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
CREATE INDEX IX2_ExcelFiles On ExcelFiles(ExcelType, QStatusId)
GO

---

-- ExcelSheets

---

CREATE TABLE ExcelRows(
	ExcelRowId int IDENTITY(1,1) NOT NULL,
	ExcelFileId int NOT NULL,
	ExcelRowNum int NOT NULL,
	InputData varchar(max),
	OutputData varchar(max),
	QStatusId int NULL,
	QStatusTime datetime NULL,
 CONSTRAINT PK_ExcelRows PRIMARY KEY NONCLUSTERED 
(
	ExcelRowId ASC
))
GO

ALTER TABLE ExcelRows  WITH CHECK ADD  CONSTRAINT FK_ExcelRows_ExcelFiles FOREIGN KEY(ExcelFileId) REFERENCES ExcelFiles (ExcelFileId)
GO

ALTER TABLE ExcelRows CHECK CONSTRAINT FK_ExcelRows_ExcelFiles
GO