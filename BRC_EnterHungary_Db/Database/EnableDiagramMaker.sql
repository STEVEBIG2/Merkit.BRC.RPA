-- diagram készítéshez kell
-- https://stackoverflow.com/questions/25845836/could-not-obtain-information-about-windows-nt-group-user
USE BRC_Hungary
GO 
ALTER DATABASE BRC_Hungary set TRUSTWORTHY ON; 
GO 
EXEC dbo.sp_changedbowner @loginame = N'sa', @map = false 
GO 
sp_configure 'show advanced options', 1; 
GO 
RECONFIGURE; 
GO 
sp_configure 'clr enabled', 1; 
GO 
RECONFIGURE; 
GO