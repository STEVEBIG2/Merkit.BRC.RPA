USE BRC_Hungary_Test 
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

DROP PROCEDURE [dbo].[InsertExcelSheet]
GO

-- =============================================  
-- Author: Steve
-- Create date: 2024.05.13
-- Description: Insert Excel Sheet
-- ============================================= 
CREATE PROCEDURE [dbo].[InsertExcelSheet]
  @ExcelFileId int,
  @ExcelSheetName varchar(50),
  @QStatusId int,
  @RobotName varchar(50)
AS
BEGIN  
 SET NOCOUNT ON; 

 DECLARE @ExcelSheetId int
 DECLARE @AktQStatusId int
 SET @ExcelSheetId = -1
 SET @AktQStatusId = 0

 SELECT @ExcelSheetId=ExcelSheetId, @AktQStatusId = QStatusId 
   FROM ExcelSheets
   WHERE @ExcelFileId=ExcelFileId AND @ExcelSheetName=ExcelSheetName
 
 IF @ExcelSheetId<0
    Begin
      INSERT INTO ExcelSheets (
            ExcelFileId,
            ExcelSheetName,
            QStatusId,
            QStatusTime
         ) VALUES (
    	    @ExcelFileId,
    	    @ExcelSheetName,
            @QStatusId,
            getdate()
    	)
	  SET @ExcelSheetId = @@IDENTITY
	  SET @AktQStatusId = @QStatusId
    End

 IF @AktQStatusId NOT IN (0,1)
    Begin
	  SET @ExcelSheetId = -1
    End

  RETURN @ExcelSheetId
END
GO