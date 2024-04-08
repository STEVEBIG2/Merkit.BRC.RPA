USE NCore
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================  
-- Author: Steve
-- Create date: 2020.09.14
-- Description: Get All Excel Row
-- ============================================= 
CREATE PROCEDURE [dbo].[GetAllExcelRows]
  @ExcelId int
AS
BEGIN  
 SET NOCOUNT ON; 
 
 SELECT f.ExcelFileName, r.*
  FROM ExcelRows r INNER JOIN ExcelFiles f ON (r.ExcelFileId=f.ExcelFileId)
  WHERE f.ExcelFileId = @ExcelId
 
  RETURN 0
END
GO


