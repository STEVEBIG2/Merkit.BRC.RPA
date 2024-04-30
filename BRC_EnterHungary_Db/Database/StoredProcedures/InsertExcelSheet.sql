USE BRC_Hungary_Test 
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================  
-- Author: Steve
-- Create date: 2024.04.30
-- Description: Insert Excel Sheet
-- ============================================= 
CREATE PROCEDURE [dbo].[InsertExcelSheet]
  @ExcelFileId int,
  @ExcelFileName varchar(50),
  @RobotName varchar(50)
AS
BEGIN  
 SET NOCOUNT ON; 
 
  INSERT INTO ExcelSheets (
        ExcelFileId,
        ExcelSheetName,
        QStatusId,
        QStatusTime
     ) VALUES (
	    @ExcelFileId,
	    @ExcelFileName,
        0,
        getdate()
	)

  --SELECT @@IDENTITY AS NewExcelFileId
  RETURN @@IDENTITY
END
GO