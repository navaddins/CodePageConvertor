# CodePageConvertor
User Defined Function

	SELECT master.dbo.TrimChar('?*hello?*',0,'?,*',',') -- 0 First Char 
	Result : hello
	SELECT master.dbo.TrimChar('?*hello?*',1,'?,*',',') -- 1 Last Char
	Result : hello
	SELECT master.dbo.TrimChar('?*hello?*',2,'?,*',',') -- 2 First & Last
	Result : hello

Replace the String

	SELECT master.dbo.ReplaceString(' hello world ','[\t|\n|\r|\s]',' ',0) -- 0 Trim First 
	Result : hello world
	SELECT master.dbo.ReplaceString(' hello world ','[\t|\n|\r|\s]',' ',1) -- 1 Trim Last
	Result : hello world
	SELECT master.dbo.ReplaceString(' hello world ','[\t|\n|\r|\s]',' ',2) -- 2 Trim First & Last
	Result : hello world
  
Get Avaiable CodePage List

	SELECT * FROM master.dbo.GetCodePage() ORDER BY Value
Change CodePage
    
    Change Windows-1252 to Chinese Traditional Code Page 950
    SELECT [master].dbo.[CodePageConvertor](''¤k«ÄÃM¦Û¦æ¨®'', 1, 1252, 950, 1) AS [GetString[Encode]]];
