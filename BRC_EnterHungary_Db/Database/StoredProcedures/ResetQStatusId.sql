--BEGIN TRAN
-- update ExcelRows set OutputData = inputdata where OutputData Is null

UPDATE [dbo].[ExcelFiles] SET QStatusId=3 WHERE ExcelFileId in (5,6)
UPDATE [dbo].[ExcelRows] SET QStatusId=4 WHERE ExcelFileId in (5,6)

-- COMMIT

-- ROLLBACK[dbo].[GetNextTransactionData]