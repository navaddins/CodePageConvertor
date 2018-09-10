/********************************************************
 Script to active code page
 Developer : NavAddIns
 Date : 2018-08-08

 Input: 
 @verbose        		: 0 is disable; 1 is enabled the clr. Default is 0
 @dbName				: Set Your DB Name
 @assemblyPath			: Set CodePageConvertor.dll File Path
*********************************************************/

RAISERROR('',0,1);
-------------------------------------------- Input Start ----------------------------------------
DECLARE @verbose bit = 0;
DECLARE @dbName nvarchar(50) = QUOTENAME('master')
DECLARE @dbUser nvarchar(50) = N''
DECLARE @assemblyPath nvarchar(max) = '''YourCodePageConvertorFolder\CodePageConvertor.dll'''

-------------------------------------------- Input End ----------------------------------------
DECLARE @sql nvarchar(max) = N''
DECLARE @codepagefromName nvarchar(50) = 'Windows_1252';
DECLARE @codepagetoName nvarchar(50) = 'big5'; -- Chinese Traditional
DECLARE @printMsg nvarchar(max) = 'This is a message, no need to do user action.<<%s>>'

SET @sql = N'''show advanced option'',' + '''' + CAST(@verbose AS varchar(1)) + '''';
EXECUTE (N'USE [master]; EXEC sp_configure N' + @sql);

SET @sql = N'''clr enabled'',' + '''' + CAST(@verbose AS varchar(1)) + '''';
EXECUTE (N'USE [master]; EXEC sp_configure N' + @sql);
EXECUTE (N'USE [master]; RECONFIGURE WITH OVERRIDE;');

SELECT @dbUser = QUOTENAME(SL.Name) FROM  master..sysdatabases SD inner join master..syslogins SL on  SD.SID = SL.SID
	WHERE  SD.Name = PARSENAME(@dbName,1)

IF (LTRIM(RTRIM(LOWER(PARSENAME(@dbName,1)))) <> 'master') BEGIN
	SET @sql = 'USE' + SPACE(1) + @dbName + CHAR(10);
	SET @sql = @sql + 'EXEC sp_changedbowner' + SPACE(1) + @dbUser + ';' + SPACE(1) + CHAR(10)
	EXEC sp_executesql @sql
END;

SET @sql = 'USE' + SPACE(1) + @dbName + CHAR(10);
IF (@verbose = 1)
	SET @sql = @sql + 'ALTER DATABASE' + SPACE(1) + @dbName + SPACE(1) + 'SET TRUSTWORTHY ON;';
ELSE
	SET @sql = @sql + 'ALTER DATABASE' + SPACE(1) + @dbName + SPACE(1) + 'SET TRUSTWORTHY OFF;';
EXEC sp_executesql @sql

SET @sql = 'USE' + SPACE(1) + @dbName + CHAR(10);
SET @sql = @sql + 'IF EXISTS (SELECT [name] FROM sys.objects WHERE sys.objects.[object_id] = OBJECT_ID(N''TrimChar'') AND type = N''FS'') DROP FUNCTION [TrimChar]' + ';' + CHAR(10);
SET @sql = @sql + 'IF EXISTS (SELECT [name] FROM sys.objects WHERE sys.objects.[object_id] = OBJECT_ID(N''ReplaceString'') AND type = N''FS'') DROP FUNCTION [ReplaceString]' + ';' + CHAR(10);
SET @sql = @sql + 'IF EXISTS (SELECT [name] FROM sys.objects WHERE sys.objects.[object_id] = OBJECT_ID(N''CodePageConvertor'') AND type = N''FS'') DROP FUNCTION [CodePageConvertor]' + ';' + CHAR(10);
SET @sql = @sql + 'IF EXISTS (SELECT [name] FROM sys.objects WHERE sys.objects.[object_id] = OBJECT_ID(N''GetCodePage'') AND type = N''FT'') DROP FUNCTION [GetCodePage]' + ';' + CHAR(10);
SET @sql = @sql + 'IF EXISTS (SELECT [name] FROM sys.assemblies WHERE sys.assemblies.[name] = N''CodePageConvertor'') DROP ASSEMBLY [CodePageConvertor]' + ';';
EXEC sp_executesql @sql

RAISERROR('********************* UnRegister **********************************',0,1);
RAISERROR(@printMsg,0,1,'TrimChar user defined function is dropped.');
RAISERROR(@printMsg,0,1,'ReplaceString user defined function is dropped.');
RAISERROR(@printMsg,0,1,'CodePageConvertor user defined function is dropped.');
RAISERROR(@printMsg,0,1,'GetCodePage user defined function is dropped.');
RAISERROR(@printMsg,0,1,'CodePageConvertor assembly is unregistered.');

IF (@verbose = 1) BEGIN
	SET @sql = N'';
	SET @sql = 'USE' + SPACE(1) + @dbName + CHAR(10);
	SET @sql = @sql + 'CREATE ASSEMBLY [CodePageConvertor] AUTHORIZATION dbo FROM' + SPACE(1) + @assemblyPath + SPACE(1) + 'WITH PERMISSION_SET = SAFE;'
	EXEC sp_executesql @sql
	
	-- Trim the Char (0 First Char, 1 End Char, 2 Frist & End Char)
	-- Usage. 
	--SELECT master.dbo.TrimChar('?*hello',0,'?,*',',') 
	--	Result : hello
	--SELECT master.dbo.TrimChar('hello*?',1,'?,*',',') 
	--	Result : hello
	--SELECT master.dbo.TrimChar('?hello*?',2,'?,*',',)
	--	Result : hello
	SET @sql = N''
	SET @sql = @sql + 'CREATE FUNCTION TrimChar (@value nvarchar(250), @TrimPos int, @TrimChars nvarchar(250),@TrimCharDelimeter nvarchar(250))' + CHAR(10)
		+ 'RETURNS nvarchar(250)' + CHAR(10)
		+ 'EXTERNAL NAME CodePageConvertor.[CodePageConvertor.Converter].[TrimChar]'
	EXECUTE (N'USE' + @dbName + '; EXEC sp_executesql N'''+ @sql +'''')
	
	-- Replace the String
	-- Usage.
	--SELECT master.dbo.ReplaceString(' hello world ','[\t|\n|\r|\s]',' ',2)
	--	Result : hello world
	SET @sql = N''
	SET @sql = @sql + 'CREATE FUNCTION ReplaceString (@value nvarchar(250), @pattern nvarchar(50), @replaceString nvarchar(50), @TrimPos int)' + CHAR(10)
		+ 'RETURNS nvarchar(250)' + CHAR(10)
		+ 'EXTERNAL NAME CodePageConvertor.[CodePageConvertor.Converter].[ReplaceString]'
	EXECUTE (N'USE' + @dbName + '; EXEC sp_executesql N'''+ @sql +'''')

	-- Change CodePage
	SET @sql = N''
	SET @sql = @sql + 'CREATE FUNCTION CodePageConvertor (@value nvarchar(250), @removeCrLf bit, @codepagefrom int, @codepageto int,@isEncode bit = 1)' + CHAR(10)
		+ 'RETURNS nvarchar(250)' + CHAR(10)
		+ 'EXTERNAL NAME CodePageConvertor.[CodePageConvertor.Converter].[GetString]'
	EXECUTE (N'USE' + @dbName + '; EXEC sp_executesql N'''+ @sql +'''')

	-- Get Avaiable CodePage List
	SET @sql = N''
	SET @sql = @sql + 'CREATE FUNCTION dbo.GetCodePage()' + CHAR(10)
		+ 'RETURNS TABLE (Code nvarchar(50), [Value] int) AS' + CHAR(10)
		+ 'EXTERNAL NAME CodePageConvertor.[CodePageConvertor.Converter].[GetCodePage];'
	EXECUTE (N'USE' + @dbName + '; EXEC sp_executesql N'''+ @sql +'''')
	
	RAISERROR('********************* Register ************************************',0,1);
	RAISERROR(@printMsg,0,1,'CodePageConvertor assembly is registered.');
	RAISERROR(@printMsg,0,1,'TrimChar user defined function is created.');
	RAISERROR(@printMsg,0,1,'ReplaceString user defined function is created.');
	RAISERROR(@printMsg,0,1,'CodePageConvertor user defined function is created.');
	RAISERROR(@printMsg,0,1,'GetCodePage user defined function is created.');

	-- Example of GetCodePage
	SET @sql = 'USE' + SPACE(1) + @dbName + CHAR(10);
	SET @sql = @sql + 'SET NOCOUNT ON' + CHAR(10);
	SET @sql = @sql + 'SELECT * FROM GetCodePage() ORDER BY Value;'
	EXEC sp_executesql @sql

	-- Example of CodePageConvertor
	SET @sql = 'USE' + SPACE(1) + @dbName + CHAR(10);
	SET @sql = @sql + 'SET NOCOUNT ON' + CHAR(10);
	SET @sql = @sql + 'DECLARE @removeCrLf bit = 1;' + CHAR(10);
	SET @sql = @sql + 'DECLARE @value nvarchar(max) = N''¤k«ÄÃM¦Û¦æ¨®'';' + CHAR(10);
	SET @sql = @sql + 'DECLARE @fromCodePage int;' + CHAR(10);
	SET @sql = @sql + 'DECLARE @toCodePage int;' + CHAR(10);
	SET @sql = @sql + 'DECLARE @isEncode bit = 1;' + CHAR(10);
	SET @sql = @sql + 'DECLARE @isDecode bit = 0;' + CHAR(10);

	SET @sql = @sql + 'SELECT @fromCodePage = [Value] FROM' + SPACE(1) + @dbName + '.dbo.GetCodePage() WHERE Code=''' + @codepagefromName + ''';' + CHAR(10);
	SET @sql = @sql + 'SELECT @toCodePage = [Value] FROM' + SPACE(1) + @dbName + '.dbo.GetCodePage() WHERE Code=''' + @codepagetoName + ''';' + CHAR(10);
	SET @sql = @sql + 'SELECT' + SPACE(1) + @dbName + '.dbo.[CodePageConvertor](@value, @removeCrLf, @fromCodePage, @toCodePage, @isEncode) AS [GetString[Encode]]];' + CHAR(10);
	SET @sql = @sql + 'SELECT @value=' + SPACE(1) + @dbName + '.dbo.[CodePageConvertor](@value, @removeCrLf, @fromCodePage, @toCodePage, @isEncode);' + CHAR(10);
	SET @sql = @sql + 'SELECT' + SPACE(1) + @dbName + '.dbo.[CodePageConvertor](@value, @removeCrLf, @fromCodePage, @toCodePage, @isDecode) AS [GetString[Decode]]];'
	EXEC sp_executesql @sql
END;