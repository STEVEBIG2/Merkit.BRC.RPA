USE NCore
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================  
-- Author: Steve
-- Create date: 2020.09.22
-- Description: Insert Excel File
-- ============================================= 
CREATE PROCEDURE [dbo].[InsertExcelFile]
  @ExcelFileName varchar(max),
  @ExcelType varchar(20),
  @RobotName varchar(50)
AS
BEGIN  
 SET NOCOUNT ON; 
 
  INSERT INTO ExcelFiles (
        ExcelFileName,
        ExcelType,
        QStatusId,
        QStatusTime,
        RobotName
     ) VALUES (
	    @ExcelFileName,
		@ExcelType,
        0,
        getdate(),
        @RobotName
	)

  --SELECT @@IDENTITY AS NewExcelFileId
  RETURN @@IDENTITY
END
GO