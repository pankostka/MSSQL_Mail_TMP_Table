-- In which databases on this server is procedure PanKostkaMailTMPTable installed?
-- Only ONLINE user databases (system databases skipped) with access for current user.
DECLARE @ProcName SYSNAME = N'PanKostkaMailTMPTable';

DROP TABLE IF EXISTS #Results
CREATE TABLE #Results
(
    DatabaseName SYSNAME,
    SchemaName SYSNAME,
    ProcedureName SYSNAME
);

DECLARE @sql NVARCHAR(MAX)

SELECT @sql = STRING_AGG(CONVERT(NVARCHAR(MAX), N'
USE ' + QUOTENAME(name) + N';

INSERT INTO #Results
SELECT
    DB_NAME(),
    s.name,
    p.name
FROM sys.procedures p
JOIN sys.schemas s ON p.schema_id = s.schema_id
WHERE p.name = @ProcName;
'), N'')
FROM sys.databases
WHERE 1=1
    AND state = 0 -- ONLINE
    AND database_id > 4 -- skip system databases
    AND HAS_DBACCESS(name) = 1; -- skip databases without access

EXEC sp_executesql @sql, N'@ProcName SYSNAME', @ProcName;

SELECT *
FROM #Results
ORDER BY DatabaseName;

DROP TABLE #Results;
