-- USE NCore
-- GO

-- DROP TABLE ExcelRows
-- GO
-- DROP TABLE ExcelFiles
-- GO
-- DROP TABLE QStatuses
-- GO

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

--

CREATE TABLE [dbo].[ExcelFiles](
	[ExcelFileId] [int] IDENTITY(1,1) NOT NULL ,
	[ExcelFileName] [varchar](50) NOT NULL,
	[ExcelType] [varchar](20) NOT NULL,
	[QStatusId] [int] NULL,
	[QStatusTime] [datetime] NULL,
	[RobotName] [varchar](50) NULL,
 CONSTRAINT [PK_ExcelFiles] PRIMARY KEY NONCLUSTERED 
(
	[ExcelFileId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE UNIQUE INDEX IX1_ExcelFiles On ExcelFiles(ExcelFileName)
GO
CREATE INDEX IX2_ExcelFiles On ExcelFiles(ExcelType, QStatusId)
GO

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