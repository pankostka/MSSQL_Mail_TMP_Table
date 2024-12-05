
--********** I need TMP TABLE to be send **********
if 1=1 BEGIN 

	DROP TABLE IF EXISTS #PanKostkaTestTMPTable -- table name must be unique. If another TMP tables (in different session) with same name exists, procedure will return error
	SELECT 	
		 CONVERT(int, 1) as ValueINT
		,CONVERT(date, GETDATE()) as ValueDate
		,CONVERT(datetime, GETDATE()) as ValueDateTime
		,CONVERT(varchar(200), 'TestTestTest') as ValueVarchar
		,CONVERT(varchar(200), 'Èùrala paninka v údolí') as ValueNVarchar
	INTO #PanKostkaTestTMPTable
	select * from #PanKostkaTestTMPTable --just to see table

END

--********** Another TMP table **********
if 1=0 BEGIN 

	DROP TABLE IF EXISTS #PanKostkaDatabasesList
	SELECT 	
		 D.name AS [Jméno databáze]
		,round(sum((F.size * 8) / 1024) / 1024.0, 1) AS [Velikost GB]
		,min(d.create_date) AS [Vytvoøena]
		,datediff(day, min(d.create_date), getdate()) AS [Stáøí dny]
		,min(f.state_desc) AS [State]	
	into #PanKostkaDatabasesList
	FROM sys.master_files f
	JOIN sys.databases D ON D.database_id = F.database_id
	WHERE 1 = 1
	AND D.name NOT IN ('master','model','msdb','tempdb')
	GROUP BY D.name
	select * from #PanKostkaDatabasesList --just to see table

END


--********** Send table **********
if 1=1 BEGIN 

	declare @MyProfileName nvarchar(200)
	set @MyProfileName  = (select name from msdb.dbo.sysmail_profile p) -- EXEC msdb.dbo.sysmail_help_profile_sp

	declare @MySubject varchar(255) -- better not nvarchar?
	SET @MySubject = 'Mail by PanKostkaMailTMPTable. Number of rows: ' + (select convert(varchar(10),COUNT(*)) from #PanKostkaTestTMPTable)


	EXEC PanKostkaMailTMPTable 
		 @TMPTableName  = '#PanKostkaTestTMPTable' -- table you want to send
		,@order_by = 'ValueINT ASC,ValueDate DESC'
		,@profile_name = @MyProfileName --SMTP Profile in MSSQL. List of profiles : EXEC msdb.dbo.sysmail_help_profile_sp
		,@subject = @MySubject
		,@recipients = 'pankostka@gmail.com'
		,@copy_recipients = NULL
		,@from_address = NULL
		,@reply_to = 'pankostka@gmail.com'
		,@prefix  = '<p>Some html formatiing<p>
<p>Link: <A  href="https://pankostka.cz/">pankostka.cz</a></p>
<p>&nbsp</p>'
		,@sufix = '<p>&nbsp</p>
<p>&nbsp</p>
<p><b>PanKostka.cz</b></p>
<a href="mailto:pankostka@gmail.cz">pankostka@gmail.com</a>
<p>&nbsp</p>
<p>Myil by procedure <i>PanKostkaMailTMPTable</i></p>' 
		,@PrintOnly = 'N'

END 