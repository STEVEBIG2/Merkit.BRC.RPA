USE BRC_Hungary_Test
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================  
-- Author: Steve
-- Create date: 2024.04.09
-- Description: Set All Status Id
-- ============================================= 
CREATE PROCEDURE [dbo].[SetAllStatusId]
  @ExcelFileId int,
  @QStatusId int,
  @RobotName varchar(50)
AS
BEGIN  
 SET NOCOUNT ON;
 
 UPDATE ExcelFiles SET
   RobotName=@RobotName,
   QStatusId=@QStatusId,
   QStatusTime=getdate()
  WHERE ExcelFileId=@ExcelFileId

/*
 UPDATE ExcelRows SET
   QStatusId=@QStatusId,
   QStatusTime=getdate() 
  WHERE ExcelFileId=@ExcelFileId
*/  
 RETURN 0
END
GO


