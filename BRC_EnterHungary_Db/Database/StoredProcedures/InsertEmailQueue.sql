USE BRC_Hungary_Test 
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

DROP PROCEDURE [dbo].[InsertEmailQueue]
GO


-- =============================================  
-- Author: Steve
-- Create date: 2024.05.22
-- Description: Insert Email Queue
-- ============================================= 
CREATE PROCEDURE [dbo].[InsertEmailQueue]
  @ExcelFileId int,
  @EmailTo varchar(250),
  @EmailCC varchar(250)=null,
  @EmailBCC varchar(250)=null,
  @EmailSubject varchar(250),
  @EmailBody varchar(max),
  @Attachments varchar(250)=null,
  @RobotName varchar(50)

AS
BEGIN  
 SET NOCOUNT ON; 
 
  INSERT INTO EmailQueue
          (ExcelFileId,
           EmailTo,
           EmailCC,
           EmailBCC,
           EmailSubject,
           EmailBody,
           Attachments,
           RobotName,
           CreateTime,
           SentTime
     ) VALUES (
           @ExcelFileId,
           @EmailTo,
           @EmailCC,
           @EmailBCC,
           @EmailSubject,
           @EmailBody,
           @Attachments,
           @RobotName,
		   getdate(),
		   NULL
	)

  --SELECT @@IDENTITY AS NewExcelFileId
  RETURN @@IDENTITY
END
GO