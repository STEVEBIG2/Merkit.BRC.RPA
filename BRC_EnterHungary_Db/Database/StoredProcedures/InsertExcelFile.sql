USE BRC_Hungary_Test 
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

DROP PROCEDURE [dbo].[InsertExcelFile]
GO

-- =============================================  
-- Author: Steve
-- Create date: 2024.05.23
-- Description: Insert Excel File
-- ============================================= 
CREATE PROCEDURE [dbo].[InsertExcelFile]
  @ExcelFileName varchar(50),
  @RobotName varchar(50)
AS
BEGIN  
 SET NOCOUNT ON; 
 
  INSERT INTO ExcelFiles (
        ExcelFileName,
        QStatusId,
        QStatusTime,
	    ErrorExcelsCreated,
        RobotName
     ) VALUES (
	    @ExcelFileName,
        0,
        getdate(),
		0,
        @RobotName
	)

  --SELECT @@IDENTITY AS NewExcelFileId
  RETURN @@IDENTITY
END
GO