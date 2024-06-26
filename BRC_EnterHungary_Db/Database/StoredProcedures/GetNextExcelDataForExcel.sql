USE BRC_Hungary_Test
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================  
-- Author: Steve
-- Create date: 2024.05.08
-- Description: Get Next Excel Data For Excel
-- ============================================= 
CREATE PROCEDURE [dbo].[GetNextExcelDataForExcel]
  @ExcelType varchar(max),
  @RobotName varchar(50)
AS
BEGIN  
 SET NOCOUNT ON;
 
 DECLARE @ExcelFileId int
 SET @ExcelFileId = -1 
 
 UPDATE ExcelFiles SET
   RobotName=@RobotName,
   QStatusId=-1,
   QStatusTime=getdate(),
   @ExcelFileId=ExcelFileId   
  WHERE ExcelFileId=(SELECT TOP 1 ExcelFileId FROM ExcelFiles Where QStatusId=3 ORDER BY ExcelFileId)

 IF @ExcelFileId>-1
 Begin
/*
   UPDATE ExcelRows SET
     QStatusId=-1,
     QStatusTime=getdate() 
    WHERE ExcelFileId=@ExcelFileId
*/
	SELECT * FROM View_ExcelRows WHERE ExcelFileId=@ExcelFileId ORDER BY ExcelRowId
 End
 Else
 Begin
   SELECT NULL AS OUTPUTDATA
 End
  
 RETURN @ExcelFileId
END
GO


