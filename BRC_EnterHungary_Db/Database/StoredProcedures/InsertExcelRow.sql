USE BRC_Hungary_Test
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================  
-- Author: Steve
-- Create date: 2024.04.09
-- Description: Insert Excel row
-- ============================================= 
CREATE PROCEDURE [dbo].[InsertExcelRow]
  @ExcelFileId int,
  @ExcelRowNum int,
  @InputData varchar(max),
  @OutputData varchar(max)=null
AS
BEGIN  
 SET NOCOUNT ON; 

 INSERT INTO ExcelRows (
      ExcelFileId,
	  ExcelRowNum,
      InputData,
      OutputData,
      QStatusId,
      QStatusTime
    ) VALUES (
      @ExcelFileId,
	  @ExcelRowNum,
      @InputData,
      @OutputData,
      0,
     getdate()
   )

  --SELECT @@IDENTITY AS NewExcelRowId
  RETURN @@IDENTITY
END
GO