SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER PROCEDURE dbo.PanKostkaMailTMPTable (
	 @TMPTableName Nvarchar(200) NULL -- string with TMP table name. Could be #Local or ##Global. Could not be normal table/view (IMHO TMP table is better). Table name must be unique on SQL server!! 
	,@order_by Nvarchar(200) NULL -- How to sort table. Example: 'Column1 ASC, Column3 Desc'. null -> sort by first column ASC 
	,@profile_name varchar(2000) --  List of profiles: EXEC msdb.dbo.sysmail_help_profile_sp (column Name)
	,@subject varchar(2000) = NULL -- Mail subject. if no subject -> table name and actual date
	,@recipients varchar(2000) -- Mail recipients. For example: 'pankostka@gmail.com;pankostka@gmail.com'
	,@copy_recipients varchar(2000) = NULL -- Mail copy recipients 
	,@from_address varchar(2000) = NULL -- Mail send from adress
	,@reply_to varchar(2000) = NULL -- Mail reply to.
	,@prefix nvarchar(max) = NULL -- HTML text above table. Example: '<p>Table with financial report</p>' 
	,@sufix nvarchar(max) = NULL -- HTML text below table. Signature. Example: '<p>Report created by pankostka@gmail.com</p>' 
	,@PrintOnly char(1) = 'N' -- If 'A'  -> send nothing, only print SQL to table and HTM mail.
	)
AS
BEGIN
SET NOCOUNT ON;


/***************************************************************************************************
Procedure:          dbo.PanKostkaMailTMPTable
Create Date:        2023-06-18
Author:             Jiri Dvorak (pankostka.cz)
Description:        Mail TMP table to users. 
Used for:           Reporting data FROM ERP.
Watch out for:		User must have permissions to send mail (sp_send_dbmail). https://www.mssqltips.com/sqlservertip/1100/SETting-up-database-mail-for-sql-server/
						Test: EXEC msdb.dbo.sp_send_dbmail @profile_name = 'PROFILE_NAME',@recipients = 'pankostka@gmail.com',@body = 'TEXT MAIL',@subject = 'TEST MAIL',@body_format = 'HTML'
						List of profiles: EXEC msdb.dbo.sysmail_help_profile_sp -- Column Name					
					TMP Table name must be unique! For table name #Test - if exists another #Test or #Test1 on SQL server, procedure will return error.
Changes:
					2023-06-18 Published on https://pankostka.cz/
Call:				
					EXEC dbo.PanKostkaMailTMPTable 
						 @TMPTableName  = '#MailSeznamDatabazi'
						,@order_by = NULL 
						,@profile_name =''
						,@subject = NULL
						,@recipients ='pankostka@gmail.com'
						,@copy_recipients = NULL
						,@from_address = NULL
						,@reply_to = NULL
						,@prefix  = NULL
						,@sufix = NULL
						,@PrintOnly = 'N'

For debuging:
						DECLARE @TMPTableName Nvarchar(200)  = '#MailSeznamDatabazi'
						DECLARE @order_by Nvarchar(200) = null 
						DECLARE @subject varchar(2000)  
						DECLARE @recipients varchar(2000)  
						DECLARE @prefix nvarchar(max) = NULL
						DECLARE @sufix nvarchar(max) = NULL
						DECLARE @PrintOnly char(1) = 'A'
--***************************************************************************************************/

--********** TABLE COLUMNS INTO TMP **********
-- Metadata about table FROM tempdb.sys.tables, tempdb.sys.columns, tempdb.sys.types -> #TableColumns
-- SELECT * FROM #TableColumns
-- Check: Table exists? Found only one table? Table has more then one row (no rows ->nothing to send)? Table has more then 1000 rows? -> error
BEGIN
		
	-- Table must exists
	IF (SELECT count(*) FROM tempdb.sys.tables WHERE name LIKE @TMPTableName + '%') = 0
	BEGIN
		DECLARE @ErrorNoTable VARCHAR(2000)
		SET @ErrorNoTable = 'Table ' + @TMPTableName + ' does not exists.'
		RAISERROR (@ErrorNoTable,18,1)
		RETURN
	END

	
	-- Table must be only one
	if ( SELECT count(*) FROM tempdb.sys.tables WHERE name like @TMPTableName + '%' ) > 1
	begin
		DECLARE @ErrorMoreTables varchar (2000)
		SET @ErrorMoreTables = (SELECT 'There are more tables with the name ' + @TMPTableName + '. Found ' + convert(varchar(10),count(*))  FROM tempdb.sys.tables WHERE name like @TMPTableName + '%')
		RAISERROR(@ErrorMoreTables, 18, 1)
		return  
	end


	-- Table rows to vartiable
	DECLARE @SQLCountStatement VARCHAR(2000)
	SET @SQLCountStatement = 'SELECT COUNT(*) FROM ' + @TMPTableName
	DECLARE @SQLCountResult TABLE (countRESULT INT)
	INSERT @SQLCountResult
	EXEC (@SQLCountStatement)
		--print @SQLCountStatement

	-- Table has no rows - there is nothing to send. Exit (no error).
	IF (SELECT TOP 1 countRESULT FROM @SQLCountResult) = 0
	BEGIN
		--RAISERROR('Table has no rows', 18, 1)
		RETURN
	END

	-- Table has more then 1000 rows - Exit with error (maybe send 1000?)
	IF (SELECT TOP 1 countRESULT FROM @SQLCountResult ) > 1000
	BEGIN
		RAISERROR ('Table has more then 1000 rows!',18,1)
		RETURN
	END

	-- Table metadata into #TableColumns
	DROP TABLE
	IF EXISTS #TableColumns
	SELECT @TMPTableName AS TMPTableName
		--,tab.name as table_name
		,ROW_NUMBER() OVER (ORDER BY tab.name ASC) AS OrderID
		,col.name AS column_name
		,t.name AS data_type
	INTO #TableColumns
	-- SELECT top 100 * 
	FROM tempdb.sys.tables AS tab
	JOIN tempdb.sys.columns AS col ON tab.object_id = col.object_id
	LEFT JOIN tempdb.sys.types AS t ON col.user_type_id = t.user_type_id
	WHERE 1 = 1
	AND tab.name LIKE @TMPTableName + '%'
	-- SELECT * FROM #TableColumns

END -- END TABLE COLUMNS INTO TMP

--********** HTML HEADER AND FOOTER **********
-- -> @HTMLHead,HTMLFooter
BEGIN
	
	DECLARE @HTMLHead Nvarchar(max) = ''
	SET @HTMLHead = 
N'<html>
<head><meta http-equiv=Content-Type content="text/html; charSET=iso-8859-2">
<style>
/*body {color: black; background-color: white;font-size: 80%;}*/
body {font-size: 80%;}
p {text-indent: 0px; margin: 0px;}

table.GeneratedTable {
  /*width: 100%;*/
  background-color: #ffffff; 
  border-collapse: collapse;
  border-width: 1px;
  border-color: #FFFFFF; /*#ffcc00*/
  border-style: solid;
  color: #000000;
}

table.GeneratedTable td, table.GeneratedTable th {
  border-width: 2px;
  border-color: #E3E3E3; /*#ffcc00*/
  border-style: solid;
  padding: 3px;
}

table.GeneratedTable th /*thead*/ { 
  background-color: #ffcc00;
}

</style>
</head>
<body>
' + ISNULL(@prefix,'')
	
	DECLARE @HTMLFooter Nvarchar(max)
	SET @HTMLFooter = ISNULL(@sufix,'') + N'

<p>&nbsp</p>

</body>
</html>'
	
	--print @HTMLHead + @HTMLFooter
	
END --END HTML HEADER AND FOOTER

--********** HTML BODY  **********
-- ->@HTMLBody
BEGIN
	
	-- HTML table declaration
	DECLARE @HTMLBody Nvarchar(max)	
	SET @HTMLBody = N'
<TABLE class="GeneratedTable">
<TR><th>'
	
	-- HTML table header
	SET @HTMLBody = @HTMLBody +  
	(SELECT STRING_AGG(column_name, N'</th><th>') 
	-- SELECT * 
	FROM #TableColumns
	)
	SET @HTMLBody = @HTMLBody + N'</th></TR>
'	

	-- HTML table rows
	-- Most tricky one. I read table columns names from #TableColumns into @SQLStatement. 
	DECLARE @SQLStatement varchar(max) = N''
	DECLARE @count INT = 1
	WHILE @count<= (SELECT count(*) FROM #TableColumns)
	BEGIN
		SET @SQLStatement = @SQLStatement  +
		(SELECT 
			-- select distinct name from tempdb.sys.types order by 1
			CASE WHEN t.data_type like 'date%' THEN -- all dates fonvert format
					N'''<td>''' + ' + ISNULL(format(t.[' + t.column_name + '],''yyyy-MM-dd HH:mm''),''&nbsp'') + ''</td>''+ '
				WHEN t.data_type in ('decimal','float','money','numeric','bigint', 'bit','int','smallint') THEN -- numbers to the right
					N'''<td align ="right">''' + ' + ISNULL(convert(varchar(2000),t.[' + t.column_name + ']),''&nbsp'') + ''</td>''+ '
				WHEN t.data_type in ('varchar','nvarchar','text','ntext','char') THEN --add collation 
					N'''<td align ="left">''' + ' + ISNULL(convert(varchar(2000),t.[' + t.column_name +  '] collate ' + convert(varchar(200),SERVERPROPERTY(N'Collation')) + '),''&nbsp'') + ''</td>''+ '
				ELSE	
					N'''<td>''' + ' + ISNULL(convert(varchar(2000),t.[' + t.column_name + ']),''&nbsp'') + ''</td>''+ '
			END
		-- SELECT top 100 * 
		FROM #TableColumns t  
		WHERE t.OrderID = @count		
		) 
	
		SET @count = @count + 1
	END
	
	-- In @SQLStatement is select on columns. Make full SQL query:
	SET @SQLStatement = N'SELECT 
''<TR>''+' + substring(@SQLStatement,1,len(@SQLStatement)-1)+ N'+
''</TR>''
FROM ' + (SELECT distinct TMPTableName FROM #TableColumns) + N' t
ORDER BY ' + ISNULL(@order_by,(SELECT '[' + column_name + '] ASC ' FROM #TableColumns where OrderID = 1 ) )  + '  
'  
	
	-- print sql statement
	if ISNULL(@PrintOnly,'N') = 'A' 
	BEGIN
		print @SQLStatement + '
--------------------------------------------------------------------------'
	END


	DECLARE @SQLResult TABLE (SQLResult varchar(max))
	INSERT @SQLResult
	EXEC (@SQLStatement)

	SET @HTMLBody = @HTMLBody + 
	(SELECT STRING_AGG(SQLResult, '
') 
	-- SELECT * 
	FROM @SQLResult
	)


	SET @HTMLBody = @HTMLBody + '
</TABLE>'


	DECLARE @HTML Nvarchar(max)
	SET @HTML = @HTMLHead + ISNULL(@HTMLBody,'Something wrong') + @HTMLFooter

	-- print sql statement
	if ISNULL(@PrintOnly,'N') = 'A' 
	BEGIN
		print @HTML + '
		'
	END
	

END --END  HTML BODY

--********** MAIL SUBJECT **********
-- If mail subject is NULL then table name + actual datetime. If Not subject + actual datetime.
BEGIN
	if ISNULL(@subject,'') = ''
		SET @subject = '' + @TMPTableName + '   (' + format(getdate(),'yyyy-MM-dd HH:mm') + ')'
	ELSE 
		SET @subject = @subject + ' ' + format(getdate(),'yyyy-MM-dd HH:mm')
	--print @subject

END -- END MAIL SUBJECT

--********** MAIL SEND **********
-- Send mail using msdb.dbo.sp_send_dbmail
BEGIN

	if ISNULL(@PrintOnly,'N') = 'N' 
	EXEC msdb.dbo.sp_send_dbmail
		 @profile_name = @profile_name 
		,@recipients = @recipients
		,@copy_recipients = @copy_recipients
		,@from_address = @from_address
		,@reply_to = @reply_to
		,@body = @HTML
		,@subject = @subject
		,@body_format = 'HTML' 

END --END MAIL SEND

END  -- PROCEDURE END
GO


-- GRANT EXECUTE ON dbo.PanKostkaMailTMPTable TO Admin