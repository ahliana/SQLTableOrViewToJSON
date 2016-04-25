--USE Temp
IF OBJECT_ID('dbo.SQLTableOrViewToJSON') IS NULL -- Check if SP Exists
 EXEC('CREATE PROCEDURE dbo.SQLTableOrViewToJSON AS SET NOCOUNT ON;') -- Create dummy/empty SP
GO

ALTER PROCEDURE SQLTableOrViewToJSON

@TableName VARCHAR(MAX) = '',
@ExportPath VARCHAR(MAX) = '',
@ExportName VARCHAR(MAX) = ''

AS
BEGIN	--BEGIN OF STORED PROCEDURE

--*******************************************************************
/*
Author: Ahliana Byrd, 20151201
Purpose: To take any table or view, export the data as a JSON file

EXAMPLE: EXEC SQLTableOrViewToJSON 'MyTable', 'C:\WhereIStowMyJSON\'

WARNING: This code will enable xp_cmdshell
WARNING: This code will create 2 tables and leave them so they can be checked later, easier for debugging
WARNING: This code will create a saved stored procedure, SQLTableOrViewToJSON

This creates a string to declare variables, declare a cursor, walk the cursor, and output the resulting strings to a table
The string is then executed, so that the entire procedure is essentially done in a virtual memory space
The string that will be executed is printed in the Messages window, so if something doesn't seem to be working, 
	you can scroll down, grab the generated script, and paste it into a new window so you can debug it

Metadata code by: marc_s on StackOverflow, http://stackoverflow.com/questions/2418527/sql-server-query-to-get-the-list-of-columns-in-a-table-along-with-data-types-no
*/
--*******************************************************************



--*******************************************************************
--This code will enable xp_cmdshell

-- To allow advanced options to be changed.
EXEC sp_configure 'show advanced options', 1
-- To update the currently configured value for advanced options.
RECONFIGURE
-- To enable the feature.
EXEC sp_configure 'xp_cmdshell', 1
-- To update the currently configured value for this feature.
RECONFIGURE
--*******************************************************************




--*******************************************************************
--Create the tables we need for SQL bits and output
PRINT 'Beginning creating tables needed for SQL bits and output'
	IF OBJECT_ID (N'dbo.SQLStatementsForJSONOutput', N'U') IS NULL
		BEGIN
			PRINT '<><><><><><><><><><> SQLStatementsForJSONOutput does not exist. It will be created. <><><><><><><><><><>'
			CREATE TABLE [dbo].[SQLStatementsForJSONOutput](
				[ID] [int] IDENTITY(1,1) NOT NULL,
				--[StatementOrder] [int] NULL,
				[Descriptor] [nvarchar](50) NULL,
				[StatementString] [nvarchar](max) NULL,
			PRIMARY KEY CLUSTERED 
			(
				[ID] ASC
			)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
			) ON [PRIMARY]
		END
	IF OBJECT_ID (N'dbo.StringsForOutput', N'U') IS NULL
		BEGIN
			PRINT '<><><><><><><><><><> StringsForOutput does not exist. It will be created. <><><><><><><><><><>'
			CREATE TABLE [dbo].[StringsForOutput](
				ID INT IDENTITY(1,1) PRIMARY KEY,
				[String] [nvarchar](max) NULL
			) ON [PRIMARY]
		END
PRINT 'Finished creating tables needed for SQL bits and output'
--*******************************************************************






--*******************************************************************
PRINT 'Beginning main transaction'
BEGIN TRANSACTION


	
--*******************************************************************
--Declare variables
	DECLARE @Error INT = 0
	DECLARE @CompletionMessage VARCHAR (256)= 'Completed creating table data to export.'
	DECLARE @Flowerbox VARCHAR (256)= '********************************************************'
	DECLARE @CRLF VARCHAR(20) = char(13) + char(10)
	
	DECLARE @ColName VARCHAR(50)
	DECLARE @ColType VARCHAR(20)
	DECLARE @ColMaxLen INT
	DECLARE @ColID INT
	
	DECLARE @Metadata TABLE (ColName VARCHAR(50), ColType VARCHAR(20), ColID INT, ColMaxLen INT, ColPrecision INT, ColScale INT, ColIsNullable BIT, ColPrimaryKey BIT)
	DECLARE @SQLString_Metadata VARCHAR(MAX)
	DECLARE @SQLString_DeclareStatments VARCHAR(MAX)
	DECLARE @SQLString_Table_Cursor VARCHAR(MAX)
	DECLARE @SQLString_Table_Fetch VARCHAR(MAX)
	DECLARE @SQLString_ToExecute VARCHAR(MAX)
	DECLARE @SQLString_Builder VARCHAR(MAX)
	DECLARE @SQLString_Composite VARCHAR(MAX)
	
	DECLARE @Cursor_Name VARCHAR(20)
	
	DECLARE @StatementString VARCHAR(MAX)
	
	DECLARE @IfIsNotNull VARCHAR(MAX)
	DECLARE @StringBuilderClause VARCHAR(MAX) = 'SET @StringToOutput += '
	DECLARE @Begin VARCHAR(20) = 'BEGIN' + @CRLF
	DECLARE @End VARCHAR(20) = 'END' + @CRLF
	
	DECLARE @PostTest TABLE (ID INT, String VARCHAR(MAX))
	
	PRINT 'Finished declaring variables'
--*******************************************************************


--*******************************************************************
--Set variables	
	PRINT 'Set variables'
	
	SET @Cursor_Name = 'Cur_TableData'
	
	IF UPPER(LEFT(@TableName, 4)) = 'DBO.'
		BEGIN
			SET @TableName = RIGHT(@TableName, LEN(@TableName) - 4)
		END
	
	IF UPPER(LEFT(@ExportPath, 1)) != '\'
		BEGIN
			SET @ExportPath = @ExportPath + '\'
		END
		
	PRINT 'Finished setting variables'
--*******************************************************************


--*******************************************************************
--Pre testing
	PRINT 'Pre testing'
	
--Check for the table to export
	IF COALESCE((OBJECT_ID (N'dbo.' + @TableName, N'U')), (OBJECT_ID (N'dbo.' + @TableName, N'V'))) IS NULL
		BEGIN
			PRINT '<><><><><><><><><><> ' + @TableName + ' does not exist. Nothing to export. <><><><><><><><><><>'
			GOTO EmergencyExitHatch
		END
--*******************************************************************


--*******************************************************************
--If tables exist, truncate them	

--Check for SQLStatementsForJSONOutput, create if not there, truncate
	IF OBJECT_ID (N'dbo.SQLStatementsForJSONOutput', N'U') IS NULL
		BEGIN
			PRINT '<><><><><><><><><><> SQLStatementsForJSONOutput does not exist. It should have been created already. Exiting. <><><><><><><><><><>'
			GOTO EmergencyExitHatch
		END
	ELSE
		BEGIN
			PRINT 'Truncating SQLStatementsForJSONOutput'
			TRUNCATE TABLE SQLStatementsForJSONOutput
		END

--Check for StringsForOutput, truncate
	IF OBJECT_ID (N'dbo.StringsForOutput', N'U') IS NULL
		BEGIN
			PRINT '<><><><><><><><><><> StringsForOutput does not exist. It should have been created already. Exiting. <><><><><><><><><><>'
			GOTO EmergencyExitHatch
		END
	ELSE
		BEGIN
			PRINT 'Truncating StringsForOutput'
			TRUNCATE TABLE StringsForOutput
		END
			
--*******************************************************************


--*******************************************************************
--Do yer stuff

--Metadata
	--http://stackoverflow.com/questions/2418527/sql-server-query-to-get-the-list-of-columns-in-a-table-along-with-data-types-no

	INSERT INTO @Metadata
	SELECT 
		c.name 'Column Name',
		t.Name 'Data type',
		c.column_id,
		c.max_length 'Max Length',
		c.precision ,
		c.scale ,
		c.is_nullable,
		ISNULL(i.is_primary_key, 0) 'Primary Key'
	FROM    
		sys.columns c
	INNER JOIN 
		sys.types t ON c.user_type_id = t.user_type_id
	LEFT OUTER JOIN 
		sys.index_columns ic ON ic.object_id = c.object_id AND ic.column_id = c.column_id
	LEFT OUTER JOIN 
		sys.indexes i ON ic.object_id = i.object_id AND ic.index_id = i.index_id
	WHERE
		c.object_id = OBJECT_ID(@TableName)

	SET @Error += @@ERROR
	PRINT 'Error count = ' + CONVERT(VARCHAR(20), @Error)
	
	SELECT * FROM @Metadata

	SET @SQLString_Table_Cursor = 'DECLARE  ' + @Cursor_Name + '  CURSOR FORWARD_ONLY FOR SELECT '
	SET @SQLString_Table_Fetch = 'FETCH NEXT FROM  ' + @Cursor_Name + '  INTO '
	SET @SQLString_DeclareStatments = 'DECLARE '
	SET @SQLString_Builder = ''
	SET @SQLString_Composite = ''

	
	DECLARE Cur_Metadata CURSOR FOR SELECT ColName, ColType, ColMaxLen, ColID FROM @Metadata

	OPEN Cur_Metadata
	FETCH NEXT FROM Cur_Metadata INTO @ColName, @ColType, @ColMaxLen, @ColID

	WHILE @@FETCH_STATUS = 0
	BEGIN
		SET @SQLString_DeclareStatments += '@' + @ColName + ' ' + @ColType
		
		SET @IfIsNotNull = 'IF ' + '@' + @ColName + ' IS NOT NULL' + @CRLF
		
		
		IF CHARINDEX('CHAR',@ColType) >  0
			BEGIN
				IF @ColMaxLen >0
					BEGIN
						SET @SQLString_DeclareStatments += '(' + CONVERT(NVARCHAR(20), @ColMaxLen) + ')'
					END
				ELSE
					BEGIN
						SET @SQLString_DeclareStatments += '(MAX)'
					END
				SET @SQLString_Builder = '''"' + @ColName + '": "'' + RTRIM(CONVERT(NVARCHAR(MAX), @' + @ColName + ')) + ''",''' + @CRLF
				SET @SQLString_Composite += @IfIsNotNull + @Begin + @StringBuilderClause+ @SQLString_Builder + @End + @CRLF
			END

		IF CHARINDEX('DATE',@ColType) >  0
			BEGIN
				SET @SQLString_Builder = '''"' + @ColName + '": "'' + RTRIM(CONVERT(NVARCHAR(MAX), @' + @ColName + ')) + ''",''' + @CRLF
				SET @SQLString_Composite += @IfIsNotNull + @Begin + @StringBuilderClause+ @SQLString_Builder + @End + @CRLF
			END
			
		IF CHARINDEX('INT',@ColType) >  0
			BEGIN
				SET @SQLString_Builder = '''"' + @ColName + '": '' + RTRIM(CONVERT(NVARCHAR(MAX), @' + @ColName + ')) + '',''' + @CRLF
				SET @SQLString_Composite += @IfIsNotNull + @Begin + @StringBuilderClause+ @SQLString_Builder + @End + @CRLF
			END

		IF CHARINDEX('NUMERIC',@ColType) >  0
			BEGIN
				SET @SQLString_Builder = '''"' + @ColName + '": '' + RTRIM(CONVERT(NVARCHAR(MAX), @' + @ColName + ')) + '',''' + @CRLF
				SET @SQLString_Composite += @IfIsNotNull + @Begin + @StringBuilderClause+ @SQLString_Builder + @End + @CRLF
			END

		IF CHARINDEX('BIT',@ColType) >  0
			BEGIN
				SET @SQLString_Builder = '''"' + @ColName + '": '' + '
				SET @SQLString_Builder += 'CASE WHEN @' + @ColName + ' = 1 THEN ''true'' ELSE ''false'' END + '',''' + @CRLF
				SET @SQLString_Composite += @IfIsNotNull + @Begin + @StringBuilderClause+ @SQLString_Builder + @End + @CRLF
			END

			
		SET @SQLString_DeclareStatments += ', '
		
		SET @SQLString_Table_Cursor += @ColName + ', '
		
		SET @SQLString_Table_Fetch += '@' + @ColName + ', '

		FETCH NEXT FROM Cur_Metadata INTO @ColName, @ColType, @ColMaxLen, @ColID
	END

	CLOSE Cur_Metadata
	DEALLOCATE Cur_Metadata
	
	--Remove trailing space
		SET @SQLString_DeclareStatments = SUBSTRING(@SQLString_DeclareStatments,1, DATALENGTH(@SQLString_DeclareStatments)-2)
	
	--Remove the trailing comma and space
		SET @SQLString_Table_Cursor = SUBSTRING(@SQLString_Table_Cursor,1, DATALENGTH(@SQLString_Table_Cursor)-2)
		SET @SQLString_Table_Fetch = SUBSTRING(@SQLString_Table_Fetch,1, DATALENGTH(@SQLString_Table_Fetch)-2)
	

	--Add the tablename		
		SET @SQLString_Table_Cursor += ' FROM ' + @TableName


	--Build the executable pieces in a table, easier to debug that way
		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Declare String To Output', 'DECLARE @StringToOutput NVARCHAR(MAX)')

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Declare BitValue', 'DECLARE @BitValue BIT')
		
		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Truncate Output Table', 'TRUNCATE TABLE StringsForOutput')
		
		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Declare Statements', @SQLString_DeclareStatments + @CRLF)

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Insert Opening Bracket', 'INSERT INTO StringsForOutput VALUES (''['')' + @CRLF)

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Table Cursor', @SQLString_Table_Cursor)

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Open Cursor', 'OPEN  ' + @Cursor_Name)

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Table Fetch', @SQLString_Table_Fetch + @CRLF)

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('While Fetch', 'WHILE @@FETCH_STATUS = 0')

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('BEGIN', 'BEGIN')

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Open Object', 'SET @StringToOutput = ''{''' + @CRLF)

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Build Object', @SQLString_Composite)

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Remove Trailing Comma', 'SET @StringToOutput = SUBSTRING(@StringToOutput,1, LEN(@StringToOutput)-1)')

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Close Object', 'SET @StringToOutput += ''},''')

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Table Fetch', @SQLString_Table_Fetch)

		--If this is the last record, strip off the trailing comma for the export
			INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
			VALUES ('Strip very last comma for export, Part 1', 'IF @@FETCH_STATUS != 0' + @CRLF)

			INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
			VALUES ('Strip very last comma for export, Part 2', @Begin)

			INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
			VALUES ('Strip very last comma for export, Part 3', 'SET @StringToOutput = SUBSTRING(@StringToOutput,1, LEN(@StringToOutput)-1)' + @CRLF)

			INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
			VALUES ('Strip very last comma for export, Part 4', @End)

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Insert Object', 'INSERT INTO StringsForOutput VALUES (@StringToOutput)')

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('END', 'END')

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('CLOSE', 'CLOSE  ' + @Cursor_Name + ' ')

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('DEALLOCATE', 'DEALLOCATE  ' + @Cursor_Name + ' ')

		INSERT INTO SQLStatementsForJSONOutput (Descriptor, StatementString)
		VALUES ('Insert Closing Bracket', 'INSERT INTO StringsForOutput VALUES ('']'')')


		PRINT 'Finished inserting statements into SQLStatementsForJSONOutput'
		

--Get the statement pieces out of the table, assemble, execute
	PRINT 'Get the statement pieces out of the table, assemble, execute'
	
	SET @SQLString_ToExecute = ''
	DECLARE Cur_SQLStrings CURSOR FORWARD_ONLY STATIC FOR SELECT StatementString FROM SQLStatementsForJSONOutput ORDER BY ID
	OPEN Cur_SQLStrings
	FETCH NEXT FROM Cur_SQLStrings INTO @StatementString

	PRINT 'Building string'
	WHILE @@FETCH_STATUS = 0
	BEGIN
	SET @SQLString_ToExecute += @StatementString + @CRLF
	FETCH NEXT FROM Cur_SQLStrings INTO @StatementString
	END
	CLOSE Cur_SQLStrings
	DEALLOCATE Cur_SQLStrings

	PRINT 'Finished building string'
	PRINT ''
	PRINT ''

	PRINT @Flowerbox
	PRINT @SQLString_ToExecute
	PRINT @Flowerbox
	
	PRINT ''
	PRINT ''
	PRINT 'Executing string'
	
	EXEC(@SQLString_ToExecute) 

--*******************************************************************


--*******************************************************************
--Post testing
	PRINT 'POST Testing'
	INSERT INTO @PostTest
	SELECT ID, String
	FROM StringsForOutput
	
	IF NOT EXISTS(
		SELECT * FROM @PostTest
		)
		BEGIN
			SELECT * FROM @PostTest
			
			SET @Error += 1
			PRINT 'Post testing shows errors. PLEASE CHECK.'
		END
	PRINT 'Finished POST testing.'
--*******************************************************************


--*******************************************************************
--This is the label for just bailing out of the procedure, used above if we know we're done and need to just leave
EmergencyExitHatch:
--*******************************************************************

--*******************************************************************
--Housecleaning
--*******************************************************************


IF @Error > 0 
	BEGIN
		ROLLBACK
		PRINT '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\'
		PRINT '\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/'
		PRINT 'Error count = ' + CONVERT(VARCHAR(20), @Error)
		PRINT 'ROLLING BACK TRANSACTION'
	END
ELSE 
	BEGIN
		COMMIT TRANSACTION 
		PRINT @Flowerbox
		PRINT @Flowerbox
		PRINT 'Completed - ' + @CompletionMessage
	END







--*******************************************************************
--The export has to be outside of the main transaction

DECLARE @DateTime VARCHAR(MAX) = (SELECT CONVERT(VARCHAR(20),GETDATE(),112) + '_' + LEFT(REPLACE(CONVERT(VARCHAR,GETDATE(),114),':',''),6))
DECLARE @ExportFileName VARCHAR(MAX)
DECLARE @DatabaseName VARCHAR(MAX) = (SELECT DB_NAME() AS DataBaseName)

IF @ExportName = ''
	BEGIN
		SET @ExportFileName = @ExportPath + @TableName + '_' + @DateTime + '.json'
	END
ELSE
	BEGIN
		SET @ExportFileName = @ExportPath + @ExportName
	END
--*******************************************************************
--Check for the table to export
	IF COALESCE((OBJECT_ID (N'dbo.' + @TableName, N'U')), (OBJECT_ID (N'dbo.' + @TableName, N'V'))) IS NULL
		BEGIN
			PRINT '<><><><><><><><><><> ' + @TableName + ' does not exist. Nothing to export. <><><><><><><><><><>'
		END
	ELSE
		BEGIN
			--Use BCP to export the table
			--If you don't ORDER BY ID, the closing bracket will end up in the middle, SQL oddity in the INSERT INTO statement
			PRINT 'Starting export.'
			DECLARE @cmd varchar(1000)
			SET @cmd = 'bcp "SELECT String FROM ' + @DatabaseName + '..StringsForOutput ORDER BY ID" queryout ' + @ExportFileName + ' -c -t, -UTF8 -T -S ' + @@servername
			PRINT @cmd
			EXEC master..xp_cmdshell @cmd
			PRINT 'Finished export.'
		END
--*******************************************************************


END	--END OF STORED PROCEDURE