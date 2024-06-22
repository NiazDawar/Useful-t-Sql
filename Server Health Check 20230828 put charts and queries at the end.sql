USE master
Go
--===============================================================================================
/*====================================SCRIPT==================================================*/
/*========================================================================================
-- SCRIPT PURPOSE			: SQL Server Instance/Databases Health check and assesments.
-- CREATED BY				: Niaz Dawar Whatsapp +92 335 844 3565, Email: NiazDawar@yahoo.com
-- STATUS				: complete.
-- CREATED ON				: 2023-01-01
-- OUTPUT				: Table one. Copy to notepad and save it as .html
					: Table two. copy to excel sheet.
-- PARAMETERS				: An excel sheet of 97 parameters set, that enable/disable query executions.

-- DETAIL				: A bunch of 97 queries use in health check and other queries that will look into data
					: Divided into Main/Sub catagories.
-- DESCRIPTION
	
--Total no. of queries : 97
-----------------------------------------------------------------------------------------------------------------*/
DECLARE @includefooter bit =0 -- 0: include
DECLARE @cmd varchar(max)
/*--RUN FOLLOWING QUERY TO SET VALUES AND COPY PRINTED MESSAGE HERE
DECLARE @VariableSet varchar(Max)='SET @ExcludeDbases=''USE [?] IF(db_name() NOT IN (''''Test'''',''''master'''',''''model'''',''''msdb'''',''''tempdb'''','; declare @names varchar(max) DECLARE database_names CURSOR  FOR	select [name] from sys.databases WHERE NAME NOT IN('master','model','msdb','tempdb') ORDER BY NAME OPEN database_names	FETCH NEXT FROM database_names INTO @names; WHILE @@FETCH_STATUS = 0   BEGIN  	FETCH NEXT FROM database_names INTO @names;    SET @VariableSet =@VariableSet +''''''+ @names+''''',' END PRINT @VariableSet+''''' '''') ''' CLOSE database_names   DEALLOCATE database_names
*/
-----------COPY RESULT HERE FROM ABOVE QUERY----------------

/*-----------------------------------------------------------------------------------------------------------------*/
---Main Section AREAS	type		value	descriptions
DECLARE @SQL_Server_Configuration					INT	=	1	--- (0:Enabled, 1:Disabled)
DECLARE @Hardware									INT	=	1	--- (0:Enabled, 1:Disabled)
DECLARE @SQL_Server_Instance_Engine					INT	=	1	--- (0:Enabled, 1:Disabled)
DECLARE @SQL_Agent_Jobs_information					INT	=	1	--- (0:Enabled, 1:Disabled)
DECLARE @tempdb_Check								INT	=	1	--- (0:Enabled, 1:Disabled)
DECLARE @Backup_Check								INT	=	1	--- (0:Enabled, 1:Disabled)
DECLARE @Database_information						INT	=	0	--- (0:Enabled, 1:Disabled)
DECLARE @Database_Performance						INT	=	1	--- (0:Enabled, 1:Disabled)
DECLARE @Index_optimization							INT	=	1	--- (0:Enabled, 1:Disabled)
DECLARE @Queries_and_Stored_Procedures				INT	=	1	--- (0:Enabled, 1:Disabled)
DECLARE @Database_login_users_roles_and_permissions	INT	=	1	--- (0:Enabled, 1:Disabled)
---Main Section area end	-------	-------	-------	-------


---Default Parameters set			
--Parameter	Column1	Value	Comments
DECLARE @Parm1_1 INT 	=	0	----------SQL and OS Version information for current instance (Query 1) (Version Info)
DECLARE @Parm1_2 INT 	=	1	----------SQL Server Services information (Query 8) (SQL Server Services Info)
DECLARE @Parm1_3 INT 	=	0	----------Machine, Server, SQL Edition and Product version information
DECLARE @Parm1_4 INT 	=	1	--error---Check the major product version to see if it is SQL Server 2014 CTP2 or greater
DECLARE @Parm1_5 INT 	=	0	----------Windows information (Query 12) (Windows Info)
DECLARE @Parm1_6 INT 	=	0	----------Instance configuration setups
DECLARE @Parm1_7 INT 	=	0	----------CPU VISIABLE ONLINE CHECK Query 02
DECLARE @Parm1_8 INT 	=	0	----------Physical, core and logical CPUs
DECLARE @Parm1_9 INT 	=	1	----------Get socket, physical core and logical core count from the SQL Server Error log. (Query 2) (Core Counts)
DECLARE @Parm1_10 INT 	=	0	----------CPU Count, Physical memory and other 11 Hardware information
DECLARE @Parm1_11 INT 	=	0	----------Hardware information from SQL Server 2014 (Query 18) (11 Hardware Info
DECLARE @Parm1_12 INT 	=	0	----------Basic information about OS memory amounts and state (Query 14) (System Memory)
DECLARE @Parm1_13 INT 	=	0	----------Volume info for all LUNS that have database files on the current instance (Query 27) (Volume Info)
DECLARE @Parm1_14 INT 	=	0	----------SQL Server NUMA Node information (Query 13) (SQL Server NUMA Info)
DECLARE @Parm1_15 INT 	=	0	----------Disk Space Monitoring -- (Query 06)
DECLARE @Parm1_16 INT 	=	1	--error---How much percent of space used by log files. log file percentage -- (query 08)
DECLARE @Parm1_17 INT 	=	0	----------information about your cluster nodes and their status (Query 15) (Cluster Node Properties)
DECLARE @Parm1_18 INT 	=	0	----------Information about any AlwaysOn AG cluster this instance is a part of (Query 16) (AlwaysOn AG Cluster)
DECLARE @Parm1_19 INT 	=	0	----------Overview of AG health and status (Query 17) (AlwaysOn AG Status)
DECLARE @Parm1_20 INT 	=	0	----------Database Instant File Initialization
DECLARE @Parm1_21 INT 	=	0	----------SQL Server Process Address space info (Query 7) (Process Memory)
DECLARE @Parm1_22 INT 	=	0	----------Resource Governor Resource Pool information (Query 31) (RG Resource Pools)
DECLARE @Parm2_1 INT 	=	1	----------Get processor description from Windows Registry (Query 22) (Processor Description)
DECLARE @Parm2_2 INT 	=	1	----------Get BIOS date from Windows Registry (Query 21) (BIOS Date)
DECLARE @Parm2_3 INT 	=	0	----------System Manufacturer and model number from SQL Server Error log (Query 19) (System Manufacturer)
DECLARE @Parm2_4 INT 	=	1	----------Get pvscsi info from Windows Registry (Query 20) (PVSCSI Driver Parameters)
DECLARE @Parm2_5 INT 	=	0	----------Page Life Expectancy (PLE) value for each NUMA node in current instance (Query 44) (PLE by NUMA Node)
DECLARE @Parm2_6 INT 	=	1	----------CPUs schedulers Visible Online (Query 02)
DECLARE @Parm2_7 INT 	=	0	----------Get CPU Utilization History for last 256 minutes (in one minute intervals) (Query 42) (CPU Utilization History)
DECLARE @Parm3_1 INT 	=	0	----------Look at Suspect Pages Suspect Pages)
DECLARE @Parm3_2 INT 	=	0	----------Get information on location, time and size of any memory dumps from SQL Server (Query 23) (Memory Dump Info)
DECLARE @Parm3_3 INT 	=	0	----------Detect blocking (run multiple times) (Query 41) (Detect Blocking)
DECLARE @Parm3_4 INT 	=	0	----------Isolate top waits for server instance since last restart or wait statistics clear (Query 38) (Top Waits)
DECLARE @Parm3_5 INT 	=	0	----------Memory Clerk Usage for instance (Query 46) (Memory Clerk Usage)
DECLARE @Parm3_6 INT 	=	0	----------Get Average Task Counts (run multiple times) (Query 40) (Avg Task Counts)
DECLARE @Parm3_7 INT 	=	1	----------Look for I/O requests taking longer than 15 seconds in the six most recent SQL Server Error Logs (Query 30) (IO Warnings)
DECLARE @Parm3_8 INT 	=	1	----------Memory Grants Pending value for current instance (Query 45) (Memory Grants Pending)
DECLARE @Parm3_9 INT 	=	1	----------Get a count of SQL connections by IP address (Query 39) (Connection Counts by IP Address)
DECLARE @Parm3_10 INT 	=	1	----------Get a count of SQL connections by IP address (Query 39) (Connection Counts by IP Address)
DECLARE @Parm4_1 INT 	=	0	----------Get SQL Server Agent SQL Agent Jobs and Category information (Query 10)
DECLARE @Parm4_2 INT 	=	1	----------SQL Server Agent SQL Agent Jobs enable status, frequency (Query SQL Agent Jobs information01)
DECLARE @Parm4_3 INT 	=	1	----------List SQL Agent Jobs and schedule info with schedules (Query SQL Agent Jobs information02)
DECLARE @Parm4_4 INT 	=	1	----------SQL Agent Jobs that are executing ssis packages (Query SQL Agent Jobs information ssis 01)
DECLARE @Parm4_5 INT 	=	1	----------Last five days failed SQL Agent Jobs (Query SQL Agent Jobs information ssis 02)
DECLARE @Parm4_6 INT 	=	0	----------SQL Server Agent Alert Information (Query 11) (SQL Server Agent Alerts)
DECLARE @Parm5_1 INT 	=	0	----------Get number of data files in 3. tempdb database (Query 25) (3. tempdb Data Files)
DECLARE @Parm5_2 INT 	=	0	----------tempdb database files informations --- (query 07)
DECLARE @Parm6_1 INT 	=	0	----------Database Backups for all databases in last one month period (Query 0)
DECLARE @Parm6_2 INT 	=	1	----------Last Backup information by database (Query 9) (Last Backup By Database)
DECLARE @Parm6_3 INT 	=	0	----------Look at recent Full 4. Backups for the current database (Query 75) (Recent Full 4. Backups)
DECLARE @Parm7_1 INT 	=	0	----------Recovery model, log reuse wait description, log file size, log usage size (Query 32) (Database Properties)
DECLARE @Parm7_2 INT 	=	0	----------Last successful DBCC CHECKDB that ran on the specified database
DECLARE @Parm7_3 INT 	=	1	----------Get input buffer information for the current database (Query 74) (Input Buffer)
DECLARE @Parm7_4 INT 	=	0	----------Get some key table properties (Query 66) (Table Properties)
DECLARE @Parm7_5 INT 	=	0	----------Get Table names, row counts, and compression status for clustered Index  or heap (Query 65) (Table Sizes, Rows, rowcount row count, total rows)
DECLARE @Parm7_6 INT 	=	0	----------Log space usage for current database (Query 51) (Log Space Usage)
DECLARE @Parm7_7 INT 	=	0	----------Individual File Sizes and space available for current database (Query 50) (File Sizes and Space)
DECLARE @Parm7_8 INT 	=	0	----------File names and paths for all user and system databases on instance (Query 26) (Database Filenames and Paths)
DECLARE @Parm7_9 INT 	=	0	----------Get VLF Counts for all databases on the instance (Query 34) (VLF Counts)
DECLARE @Parm8_10 INT 	=	0	----------Breaks down buffers used by current database by object (table, Index ) in the buffer cache (Query 64 b) (Buffer Usage)
DECLARE @Parm8_11 INT 	=	0	----------I/O Statistics by file for the current database (Query 52) (IO Stats By File)
DECLARE @Parm8_12 INT 	=	0	----------Calculates average stalls per read, per write, and per total input/output for each database file (Query 29) (IO Latency by File)
DECLARE @Parm8_13 INT 	=	0	----------Get total buffer usage by database for current instance (Query 37) (Total Buffer Usage by Database)
DECLARE @Parm8_14 INT 	=	0	----------Get CPU utilization by database (Query 35) (CPU Usage by Database)
DECLARE @Parm8_15 INT 	=	0	----------Drive level latency information (Query 28) (Drive Level Latency)
DECLARE @Parm8_16 INT 	=	1	----------Find missing Index  warnings for cached plans in the current database (Query 64 a) (Missing Index  Warnings)
DECLARE @Parm8_17 INT 	=	0	----------Get I/O utilization by database (Query 36) (IO Usage By Database)
DECLARE @Parm9_1 INT 	=	0	----------Get fragmentation info for all Index es above a certain size in the current database (Query 69) (Index  Fragmentation)
DECLARE @Parm9_2 INT 	=	1	----------Get in-memory OLTP Index  usage (Query 72) (XTP Index  Usage)
DECLARE @Parm9_3 INT 	=	0	----------Get lock waits for current database (Query 73) (Lock Waits)
DECLARE @Parm9_4 INT 	=	1	----------Index es that are being maintained but not used (High Write/Zero Read) (Query 03)
DECLARE @Parm9_5 INT 	=	1	----------Detailed activity information for Index es not used for user reads (Query 04)
DECLARE @Parm9_6 INT 	=	1	----------Pathways for performance improvement.----(Query 05)
DECLARE @Parm9_7 INT 	=	0	----------Index   Read/Write stats (all tables in current DB) ordered by Reads (Query 70) (Overall Index  Usage - Reads)
DECLARE @Parm9_8 INT 	=	0	----------Index  Read/Write stats (all tables in current DB) ordered by Writes (Query 71) (Overall Index  Usage - Writes)
DECLARE @Parm9_9 INT 	=	1	----------Look at most frequently modified Index es and statistics (Query 68) (Volatile Index es)
DECLARE @Parm9_10 INT 	=	0	----------Missing Index es for all databases by Index  Advantage (Query 33) (Missing Index es All Databases)
DECLARE @Parm9_11 INT 	=	1	----------Missing Index es for current database by Index  Advantage (Query 63) (Missing Index es)
DECLARE @Parm9_12 INT 	=	0	----------Possible Bad NC Index es (writes > reads) (Query 62) (Bad NC Index es)
DECLARE @Parm9_13 INT 	=	0	----------Statistics last updated on all Index es? (Query 67) (Statistics Update)
DECLARE @Parm10_1 INT 	=	0	----------Get top average elapsed time queries for entire instance (Query 49) (Top Avg Elapsed Time Queries)
DECLARE @Parm10_2 INT 	=	0	----------Lists the top statements by average input/output usage for the current database (Query 61) (Top IO Statements)
DECLARE @Parm10_3 INT 	=	0	----------Get most frequently executed queries for this database (Query 53) (Query Execution Counts)
DECLARE @Parm10_4 INT 	=	0	----------Find single-use, ad-hoc and prepared queries that are bloating the plan cache (Query 47) (Ad hoc Queries)
DECLARE @Parm10_5 INT 	=	0	----------Top Cached SPs By Avg Elapsed Time (Query 56) (SP Avg Elapsed Time)
DECLARE @Parm10_6 INT 	=	0	----------Top Cached SPs By Execution Count (Query 55) (SP Execution Counts)
DECLARE @Parm10_7 INT 	=	0	----------Top Cached SPs By Total Logical Reads. Logical reads relate to memory pressure (Query 58) (SP Logical Reads)
DECLARE @Parm10_8 INT 	=	0	----------Top Cached SPs By Total Logical Writes (Query 60) (SP Logical Writes)
DECLARE @Parm10_9 INT 	=	0	----------Top Cached SPs By Total Physical Reads. Physical reads relate to disk read I/O pressure (Query 59) (SP Physical Reads)
DECLARE @Parm10_10 INT 	=	0	----------Top Cached SPs By Total Worker time. Worker time relates to CPU cost (Query 57) (SP Worker Time)
DECLARE @Parm10_11 INT 	=	0	----------Get top total logical reads queries for entire instance (Query 48) (Top Logical Reads Queries)
DECLARE @Parm10_12 INT 	=	0	----------Get top total worker time queries for entire instance (Query 43) (Top Worker Time Queries)
DECLARE @Parm11_1 INT 	=	0	----------List of logins in SQL Server
DECLARE @Parm11_2 INT 	=	0	----------List of users in SQL Server
DECLARE @Parm11_3 INT 	=	0	----------List of roles in SQL Server
DECLARE @Parm11_4 INT 	=	0	----------Fix SQL Server orphaned users
DECLARE @Parm11_5 INT 	=	0	----------Determine how many open connections exist to the specific database
DECLARE @Parm11_6 INT 	=	0	---------Determine users/roles and its permissions on each database
---Default Parameters set (End)			




IF OBJECT_ID(N'tempdb..##RunOnce') IS NOT NULL
BEGIN
DROP TABLE ##RunOnce
PRINT 'table ##RunOnce dropped'
END


IF OBJECT_ID(N'tempdb..##DataForSheet') IS NOT NULL
BEGIN
DROP TABLE ##DataForSheet
PRINT 'table ##DataForSheet dropped'
END


IF OBJECT_ID(N'tempdb..##FragmentedIndexData') IS NULL
BEGIN
CREATE TABLE ##FragmentedIndexData([Database] varchar(max),[Schema] nvarchar(max),[Object] varchar(max),[Index Name] varchar(max), index_id int, index_type_desc nvarchar(max), avg_fragmentation_in_percent float, fragment_count bigint, page_count bigint, fill_factor tinyint,has_filter bit, filter_definition nvarchar(max), allow_page_locks bit)
PRINT 'Table ##FragmentedIndexData created successfully'
END


/*-----------------------------------------------------------------------------------------------------------------*/
--Comments (1.1) (NZ) -- Dated: 2023-01-01

--Description:
	--Collecting fragmentation data.
if(@Index_optimization=0 and @Parm9_1=0)
BEGIN
	if(select count(1) from ##FragmentedIndexData)=0
	begin
		INSERT INTO ##FragmentedIndexData	
		EXEC sp_MSforeachdb  'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) 
			SELECT DB_NAME(ps.database_id) AS [Database], SCHEMA_NAME(o.[schema_id]) AS [Schema], OBJECT_NAME(ps.OBJECT_ID) AS [Object], i.[name] AS [Index Name], ps.index_id,  ps.index_type_desc, ps.avg_fragmentation_in_percent,  ps.fragment_count, ps.page_count, i.fill_factor, i.has_filter,  i.filter_definition, i.[allow_page_locks] FROM sys.dm_db_index_physical_stats(DB_ID(),NULL, NULL, NULL , N''LIMITED'') AS ps INNER JOIN sys.indexes AS i WITH (NOLOCK) ON ps.[object_id] = i.[object_id]  AND ps.index_id = i.index_id INNER JOIN sys.objects AS o WITH (NOLOCK) ON i.[object_id] = o.[object_id] WHERE ps.database_id = DB_ID() AND ps.page_count > 1000  ORDER BY ps.avg_fragmentation_in_percent DESC OPTION (RECOMPILE);'
	end
END
-- SELECT * FROM ##FragmentedIndexData
--/*--Run below manually selecting single database every time.
--INSERT INTO ##FragmentedIndexData	
--	SELECT Quotename(DB_NAME(ps.database_id)) AS [Database], SCHEMA_NAME(o.[schema_id]) AS [Schema], OBJECT_NAME(ps.OBJECT_ID) AS [Object], i.[name] AS [Index Name], ps.index_id,  ps.index_type_desc, ps.avg_fragmentation_in_percent,  ps.fragment_count, ps.page_count, i.fill_factor, i.has_filter,  i.filter_definition, i.[allow_page_locks] FROM sys.dm_db_index_physical_stats(DB_ID(),NULL, NULL, NULL , N'LIMITED') AS ps INNER JOIN sys.indexes AS i WITH (NOLOCK) ON ps.[object_id] = i.[object_id]  AND ps.index_id = i.index_id INNER JOIN sys.objects AS o WITH (NOLOCK) ON i.[object_id] = o.[object_id] WHERE ps.database_id = DB_ID() AND ps.page_count > 2000  ORDER BY ps.avg_fragmentation_in_percent DESC OPTION (RECOMPILE);

/*---TIPS 
select @@Servername
SELECT distinct [Database] FROM ##FragmentedIndexData ORDER BY [Database] DESC
--TOp n most fragmented records
SELECT * FROM (
SELECT ROW_NUMBER() OVER (Partition by ([Database]) Order by avg_fragmentation_in_percent desc) AS TOP_Frag,
* FROM ##FragmentedIndexData WHERE index_id>2
) TOP5 WHERE TOP_Frag <=10
ORDER BY avg_fragmentation_in_percent desc

--create queries for fragmention collection from each database
select 'USE ', Quotename( name), '  ', ' INSERT INTO ##FragmentedIndexData	',' SELECT DB_NAME(ps.database_id) AS [Database], SCHEMA_NAME(o.[schema_id]) AS [Schema], OBJECT_NAME(ps.OBJECT_ID) AS [Object], i.[name] AS [Index Name], ps.index_id,  ps.index_type_desc, ps.avg_fragmentation_in_percent,  ps.fragment_count, ps.page_count, i.fill_factor, i.has_filter,  i.filter_definition, i.[allow_page_locks] FROM sys.dm_db_index_physical_stats(DB_ID(),NULL, NULL, NULL , N''LIMITED'') AS ps INNER JOIN sys.indexes AS i WITH (NOLOCK) ON ps.[object_id] = i.[object_id]  AND ps.index_id = i.index_id INNER JOIN sys.objects AS o WITH (NOLOCK) ON i.[object_id] = o.[object_id] WHERE ps.database_id = DB_ID() AND ps.page_count > 2000  ORDER BY ps.avg_fragmentation_in_percent DESC OPTION (RECOMPILE);' from sys.databases
WHERE name not  in ('master','msdb','tempdb','distribution','DWQueue' ,'DWDiagnostics','model')  order by name

--*/
---USE IN CHART
declare  @DataForChart as table(id int identity, xName varchar(max), y1 decimal(15,2), y2 decimal(15,2), y3 decimal(15,2))
declare @xParms varchar(max) 
declare @ChartTitle varchar(max) 
DECLARE @outputScript NVARCHAR(MAX);
DECLARE @Labels VARCHAR(MAX);
DECLARE @xValues VARCHAR(MAX);
---------------------------------
DECLARE @AllowFewToReport int
SELECT @AllowFewToReport =COUNT(name) from sys.databases
SET @AllowFewToReport =CASE WHEN @AllowFewToReport>10 THEN 10 ELSE @AllowFewToReport END
set @AllowFewToReport=10
DECLARE @No_of_database_having_full_backup int=0
DECLARE @No_of_database_having_Transaction_log_backup int = 0
DECLARE @TotalDatabases int=0
DECLARE @ExcelSheetNo int=1
DECLARE @HugeDataCounter int=0
DECLARE @CPUs int =1, @NoofTempDbfiles int=1
DECLARE  @DateFormatNo INT =20, @QueryNo int=0
DECLARE  @QueryDesc  VARCHAR(max)
DECLARE  @Sql_exec  VARCHAR(max)
DECLARE  @QSectionNo int=0,  @QSectionSubNo int=0, @QSectionSplitNo int=0, @QRunTime  VARCHAR(max)
DECLARE  @QTotalNo int=1
DECLARE @QHeadingTitle VARCHAR(max)
DECLARE @QHeadingOne VARCHAR(max)
DECLARE @QHeadingTw0 VARCHAR(max)
DECLARE @QGoodPoint VARCHAR(max)
DECLARE @QBadPoint VARCHAR(max)
 
declare @ExpandStart varchar(256)=''--
declare @ExpandEnd varchar(130)=''--
declare @Owner varchar(15)='NiazD'
---tbl1
CREATE TABLE  ##RunOnce(
R1					INT DEFAULT 0, 
R2					INT DEFAULT 0, 
S1 					INT DEFAULT 0, 
QuerySort			INT DEFAULT 0, 
R3					INT DEFAULT 0, 
R4					VARCHAR(MAX) DEFAULT '', 
R5					VARCHAR(MAX) DEFAULT '', 
TRTH_TAG			VARCHAR(400) DEFAULT '',
TD					VARCHAR(400) DEFAULT '',
TDR					VARCHAR(400) DEFAULT '',
TDC					VARCHAR(400) DEFAULT '',
C5					VARCHAR(MAX) DEFAULT '', 
C6					VARCHAR(MAX) DEFAULT '', 
C7					VARCHAR(MAX) DEFAULT '', 
C8					VARCHAR(MAX) DEFAULT '', 
C9					VARCHAR(MAX) DEFAULT '', 
C10					VARCHAR(MAX) DEFAULT '', 
C11					VARCHAR(MAX) DEFAULT '', 
C12					VARCHAR(MAX) DEFAULT '', 
C13					VARCHAR(MAX) DEFAULT '', 
C14					VARCHAR(MAX) DEFAULT '', 
C15					VARCHAR(MAX) DEFAULT '', 
C16					VARCHAR(MAX) DEFAULT '', 
C17					VARCHAR(MAX) DEFAULT '', 
C18					VARCHAR(MAX) DEFAULT '', 
C19					VARCHAR(MAX) DEFAULT '', 
C20					VARCHAR(MAX) DEFAULT '', 
C21					VARCHAR(MAX) DEFAULT '', 
C22					VARCHAR(MAX) DEFAULT '', 
C23					VARCHAR(MAX) DEFAULT '', 
C24					VARCHAR(MAX) DEFAULT '', 
C25					VARCHAR(MAX) DEFAULT '', 
C26					VARCHAR(MAX) DEFAULT '', 
C27					VARCHAR(MAX) DEFAULT '', 
C28					VARCHAR(MAX) DEFAULT '', 
C29					VARCHAR(MAX) DEFAULT '', 
C30					VARCHAR(MAX) DEFAULT '', 
C31					VARCHAR(MAX) DEFAULT '', 
C32					VARCHAR(MAX) DEFAULT '', 
C33					VARCHAR(MAX) DEFAULT '', 
C34					VARCHAR(MAX) DEFAULT '', 
C35					VARCHAR(MAX) DEFAULT '', 
C36					VARCHAR(MAX) DEFAULT '',
C37					VARCHAR(MAX) DEFAULT '',
C38					VARCHAR(MAX) DEFAULT '',
C39					VARCHAR(MAX) DEFAULT '',
C40					VARCHAR(MAX) DEFAULT '',
H1					VARCHAR(100) DEFAULT '',
H2					VARCHAR(100) DEFAULT ''
)
--tbl2
CREATE TABLE  ##DataForSheet(
  R1	INT DEFAULT 0 
, R2	INT DEFAULT 0 
, RR2 INT DEFAULT 0
, S1 		INT DEFAULT 0 
, QuerySort			INT DEFAULT 0
, R3	INT DEFAULT 0 
, R4	VARCHAR(MAX) DEFAULT ''
, R5	VARCHAR(MAX) DEFAULT ''
, H2	VARCHAR(MAX) DEFAULT ''
, C5	VARCHAR(MAX) DEFAULT ''
, C6	VARCHAR(MAX) DEFAULT ''
, C7	VARCHAR(MAX) DEFAULT ''
, C8	VARCHAR(MAX) DEFAULT ''
, C9	VARCHAR(MAX) DEFAULT ''
, C10	VARCHAR(MAX) DEFAULT ''
, C11	VARCHAR(MAX) DEFAULT ''
, C12	VARCHAR(MAX) DEFAULT ''
, C13	VARCHAR(MAX) DEFAULT ''
, C14	VARCHAR(MAX) DEFAULT ''
, C15	VARCHAR(MAX) DEFAULT ''
, C16	VARCHAR(MAX) DEFAULT ''
, C17	VARCHAR(MAX) DEFAULT ''
, C18	VARCHAR(MAX) DEFAULT ''
, C19	VARCHAR(MAX) DEFAULT ''
, C20	VARCHAR(MAX) DEFAULT ''
, C21	VARCHAR(MAX) DEFAULT ''
, C22	VARCHAR(MAX) DEFAULT ''
, C23	VARCHAR(MAX) DEFAULT ''
, C24	VARCHAR(MAX) DEFAULT ''
, C25	VARCHAR(MAX) DEFAULT ''
, C26	VARCHAR(MAX) DEFAULT ''
, C27	VARCHAR(MAX) DEFAULT ''
, C28	VARCHAR(MAX) DEFAULT ''
, C29	VARCHAR(MAX) DEFAULT ''
, C30	VARCHAR(MAX) DEFAULT ''
, C31	VARCHAR(MAX) DEFAULT ''
, C32	VARCHAR(MAX) DEFAULT ''
, C33	VARCHAR(MAX) DEFAULT ''
, C34	VARCHAR(MAX) DEFAULT ''
, C35	VARCHAR(MAX) DEFAULT ''
, C36	VARCHAR(MAX) DEFAULT ''
, C37	VARCHAR(MAX) DEFAULT ''
, C38	VARCHAR(MAX) DEFAULT ''
, C39	VARCHAR(MAX) DEFAULT ''
, C40	VARCHAR(MAX) DEFAULT ''
, ExclSheet int default 0
)

INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5, H1) 
SELECT 0, 0, 0, 0, 'HTML',
'<!DOCTYPE html>
<html>
<head>

<meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="D:\2. Freelance\00. Sql Server DBA\01 An sql server health and performance\New way for healh check\Health Check Latest\prism.css">

    <script src="https://unpkg.com/sql-formatter@2.3.3/dist/sql-formatter.min.js"></script>
    <script>
    document.addEventListener("DOMContentLoaded", function() {    
    var format = window.sqlFormatter.format;    
    document.getElementById("Query").innerText=format(document.getElementById("Query").innerText);
});
</script>

<link rel="stylesheet" href="style01.css">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.9.4/Chart.js"></script>
<script src="https://cdn.plot.ly/plotly-2.25.2.min.js"></script>
<style>
.intro {
    background-color: #f2eeee;
    PADDING: 2PX;
    MARGIN-LEFT: 5PX;
}


.summary {
    background-color: #ffeb3b;
    color: #000;
    margin-left: 30px;
    font-size: 18px;
    padding: 3px;
    margin-top: 0px;
}
.critical{
    color:red;
    text-decoration:underline;
}
.success {
    color: #1ac41a;
    border-left: 3px solid #068706;
    padding-left: 10px;
}
.safe {
    color: #07b707;
    font-weight: bold;
}
h1 {
    margin-top: 20px;
    margin-bottom: 1px;
}
h2 {
    margin-bottom: 0px;
}

textarea {
    width: 100%;
    height: 150px;
    padding: 12px 20px;
    box-sizing: border-box;
    border: 1px solid #fff;
    border-radius: 4px;
    background-color: #ffffff;
    font-size: 14px;
	resize:vertical;
}
textarea:focus {
    outline: none !important;
    border:1px solid red;
    box-shadow: 0 0 10px #719ECE;
  }

h3 {
    margin-bottom: 1px;
}
.numbering {
    color: #9f9b9b;
}
.no_data {
    display: block;
    text-align: left;
    font-size: 12pt;
    color: #eaa60f;
}
.timeStamp {
    font-size: 10pt;
    color: #9f9b9b;
}
.lessDataOnReport {
    font-size: 10pt;
    color: #9f9b9b;
}
.data_on_sheet {
    font-size: 10pt;
    color: #e92f2f;
}
.Index_Details table {
    border-collapse: collapse;
    width: 40px;
    border: 0px solid #ccc;
}
.Index_Details tr {
    border: 0px solid #ccc;
}


.Sumry_Report table {
    border-collapse: collapse;
    width: 2000px;
    border: 1px solid #ccc;
}
.fields {
    margin: 0;
    margin-left: 40px;
    font-size: x-small;
    color: #e53935;
}
.Sumry_Report th {
    /*background-image: linear-gradient(#4e86ba, #ffffff);*/
    background: #aacff0;
    font-weight: bold;
    color: black;
    border: 0px;
    font-size: 14pt;
    text-align: left;
}
.Sumry_Report tr {
    border: 0px solid #ccc;
    font-size: 12pt;
}
table {
    border-collapse: collapse;
    width: 900px;
}
tr {
    border-bottom: 1px solid #000;
    font-size: 12pt;
}
td
{
	border:0px;
    font-size: 10pt;
}
th {
    font-weight: bold;
    color: white;
    border: 0px;
    font-size: 10pt;
    text-align: left;
    color: white;
    border: 0px;
    background: #4472c4;
}
.ShowHideButton{
	color: #ffffff;
    background-color: #5a9bd7;
    padding: 5px;
    cursor: pointer;
    margin-left: 0px;
    display: inline-block;
    width: 50px;
    text-align: center;
    border-radius: 5px;
}
.AlternateQuery
{
    background-color: #fff;
}
.error {
    color: red;
    border-left: 3px solid #b71c1c;
    padding-left: 10px;
}

</style>
</head>
', 'Niaz'

declare @configName varchar(100), @minvalue int, @maxvalue int, @config_Value int, @run_value int
declare @MaxP as table([name] varchar(100),minimum int, maximum int, config_value int, run_value int)
insert into @MaxP
EXEC sp_configure 'max degree of parallelism'
select @configName=[name], @minvalue=minimum, @maxvalue=maximum, @config_Value=config_value, @run_value=run_value from @MaxP

DECLARE @CurrentDate datetime=getdate()

INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) 
SELECT 0, 0, 0, 0, 'HTML',
'<font size="29">
<b>Health Check & Assessment Report</b>
</font>'+
'
<br><br><br>
<font size="3">
<span><b>Report Date </b></span>'+	CAST(GetDATE() AS VARCHAR)+'</td></tr>'+
'<table class="Sumry_Report">'+
'<tr><td>Server<td>'+	@@SERVERNAME +'</td></tr>'+
'<tr><td> SQL Start Started <td>'+ 
		FORMAT(sqlserver_start_time, 'yyyy-MM-dd hh:mm')+' -- (Uptime '+
		CAST(DATEDIFF(MONTH, sqlserver_start_time, @CurrentDate)  AS VARCHAR)+' months & '+
		CAST(DATEDIFF(DAY, sqlserver_start_time, @CurrentDate) % 30 AS VARCHAR)+' days & '+
		CAST( DATEDIFF(HOUR, sqlserver_start_time, @CurrentDate) % 24 AS VARCHAR)+' hrs. '
+'</td></tr>'+
'<tr><td>Version<td>'+ SUBSTRING(@@VERSION,0, CHARINDEX('ft Corporation', @@VERSION )+15) +'</td></tr>'+
'<tr><td> Edition <td>'+	cast(SERVERPROPERTY('Edition') AS  VARCHAR) +'</td></tr>'+
'<tr><td> Machine <td>'+cast(SERVERPROPERTY('MachineName') AS  VARCHAR)+'</td></tr>'+
'<tr><td>Window<td>'+ SUBSTRING(@@VERSION,CHARINDEX(' Windows', @@VERSION ),4000)+'</td></tr>'+
'<tr><td>physical memory(GB)<td>'+  CAST((physical_memory_kb/(1000)/1000) AS VARCHAR)+'</td></tr>'+
'<tr><td> number of logical CPUs<td>'+CAST(cpu_count AS VARCHAR)+'</td></tr>'+
'<tr><td> '+@configName+'<td>min: '+CAST(@minvalue AS VARCHAR)+', max: '+CAST(@maxvalue AS VARCHAR)+', configured value: '+CAST(@config_Value AS VARCHAR)+', Running value: '+CAST(@run_value AS VARCHAR)+'</td></tr>'+

'</table>'+

'</font>'  FROM [sys].[dm_os_sys_info]

--INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) 
--SELECT 0, 0, 0, 0, 'HTML',
--'<font size="29"><b>Health Check & Assessment Report</b>
--</font>'+
--'<br>
--<br><br>
--<font size="3"><span>Report Date (' + 
--	CAST(GetDATE() AS VARCHAR)+')'+
--'<br><br><span><b>Server</b>&#160&#160&#160&#160&#160&#160&#160&#160&#160&#160&#160&#160&#160:&#160'+ 
--	@@SERVERNAME +
--'<br><span><b>Version</b>&#160&#160&#160&#160&#160&#160&#160&#160&#160&#160&#160:&#160'+ 
--	SUBSTRING(@@VERSION,0, CHARINDEX('on Windows', @@VERSION )) +
--'<br><span><b>Window</b>&#160&#160&#160&#160&#160&#160&#160&#160&#160:&#160'+ 
--	SUBSTRING(@@VERSION,CHARINDEX(' Windows', @@VERSION ),4000)+
--'<br><span>Physical memory :'+  (physical_memory_kb/(1000)/1000)+
--'<br><span> CPUs </span> :'+cpu_count+
--'<br><span> SQL Started : '+ sqlserver_start_time+
--'</font>'  FROM [sys].[dm_os_sys_info]


--SELECT [number_of_physical_cpus] 	,[number_of_cores_per_cpu] 	,[total_number_of_cores] 	,[number_of_virtual_cpus] 	,LTRIM(RIGHT([cpu_category], CHARINDEX('x', [cpu_category]) - 1)) AS [cpu_category] FROM [ProcessorInfo]
	
-----start----sss-----code--------code start----------------
-- 1-22 queries are related to SQL_Server_Configuration
IF(@SQL_Server_Configuration=0)
BEGIN
	SET @QSectionNo=@QSectionNo+1
	SET @QSectionSubNo=1
	SET @QSectionSplitNo=1
	SET @QTotalNo=@QTotalNo+1
    Set  @QTotalNo=1
    Set @QHeadingOne='SQL server configuration setting'
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, -1, 'H1', CAST(@QSectionNo AS VARCHAR)+'. '+@QHeadingOne) 
	--1---
/*    IF(@Parm1_1=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='SQL and OS Version information for current instance (Version Info)'
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo,@QSectionSplitNo,30,'<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Server Name',	'SQL Server and OS Version Info')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6)
	SELECT @QSectionNo,@QSectionSubNo, @QSectionSplitNo, 32, 'D', @@SERVERNAME AS [Server Name], @@VERSION AS [SQL Server and OS Version Info];
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1,R2,S1,R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm*/
	--2---
    IF(@Parm1_2=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='SQL Server Services information'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')

	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'servicename',	'process_id',	'startup_type_desc',	'status_desc',	'last_startup_time')
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'Text-Align:',	'L',	'R',	'L',	'L',	'R')

	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) 
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', servicename, process_id, startup_type_desc, status_desc,  last_startup_time FROM sys.dm_server_services WITH (NOLOCK) OPTION (RECOMPILE);
	SET @QSectionSplitNo=@QSectionSplitNo+1
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'servicename',	'service_account',	'is_clustered',	'cluster_nodename',	'filename')
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'servicename',	'L',	'L',	'L',	'L')
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) 
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', servicename, service_account, is_clustered, cluster_nodename, [filename] FROM sys.dm_server_services WITH (NOLOCK) OPTION (RECOMPILE);


	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	IF(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1,R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm

	---3--
    IF(@Parm1_3=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Machine, Server, SQL Edition and Product version information'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)

/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/ 

	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
---	1-5
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9)
	VALUES (@QSectionNo,	@QSectionSubNo, @QSectionSplitNo,  31, 'H',	'Machine', 'Server',	'Instance',	'Edition',	'Product')	

	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9)
	VALUES (@QSectionNo,	@QSectionSubNo, @QSectionSplitNo,  231, 'Text-Align:',	'L', 'R',	'L',	'L',	'R')	


	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',  cast(SERVERPROPERTY('MachineName') AS  VARCHAR) AS [MachineName],   cast( SERVERPROPERTY('ServerName')as  VARCHAR) AS [ServerName],  cast(SERVERPROPERTY('InstanceName') AS  VARCHAR) AS [Instance],  cast(SERVERPROPERTY('Edition') AS  VARCHAR) AS [Edition],  cast(SERVERPROPERTY('ProductLevel') AS   VARCHAR) AS [ProductLevel];	

----6 - 10
	SET @QSectionSplitNo=@QSectionSplitNo+1
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9)
	VALUES (@QSectionNo,	@QSectionSubNo, @QSectionSplitNo,  31, 'H', 'update',	'version',	'major',	'minor',	'build')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',  cast(SERVERPROPERTY('ProductUpdateLevel') AS  VARCHAR) AS [ProductUpdateLevel],	 cast(SERVERPROPERTY('ProductVersion') AS  VARCHAR) AS [ProductVersion],  cast(SERVERPROPERTY('ProductMajorVersion') AS  VARCHAR) AS [ProductMajorVersion],   cast(SERVERPROPERTY('ProductMinorVersion') AS  VARCHAR) AS [ProductMinorVersion],  cast(SERVERPROPERTY('ProductBuild') AS  VARCHAR) AS [ProductBuild];
------ 10 -15
	SET @QSectionSplitNo=@QSectionSplitNo+1
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9)
	VALUES (@QSectionNo,	@QSectionSubNo, @QSectionSplitNo,  31, 'H', 'type',	'full text', 'log path',	'clr version',	'xtp support ?')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',  cast(SERVERPROPERTY('ProductBuildType') AS  VARCHAR)  AS [ProductBuildType], cast(SERVERPROPERTY('IsFullTextInstalled')  AS  VARCHAR) AS [IsFullTextInstalled],  cast(SERVERPROPERTY('InstanceDefaultLogPath')  AS  VARCHAR) AS [InstanceDefaultLogPath], cast(SERVERPROPERTY('BuildClrVersion')  AS  VARCHAR) AS [Build CLR Version],  cast(SERVERPROPERTY('IsXTPSupported')  AS  VARCHAR) AS [IsXTPSupported];

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1,R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	---4--
    IF(@Parm1_4=0)
    BEGIN
    BEGIN TRY 
	SET @QHeadingTw0='Check the major product version to see if it is SQL Server 2014 CTP2 or greater'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	IF NOT EXISTS (SELECT * WHERE CONVERT( VARCHAR(128), SERVERPROPERTY('ProductVersion')) LIKE '12%')	
	BEGIN		
	DECLARE  @ProductVersion  VARCHAR(128) = CONVERT( VARCHAR(128), SERVERPROPERTY('ProductVersion'));		
	RAISERROR ('Script does not match the ProductVersion [%s] of this instance. Many of these queries may not work on this version.' , 18 , 16 , @ProductVersion);	END	ELSE		PRINT N'You have the correct major version of SQL Server for this diagnostic information script';
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	---5--
    IF(@Parm1_5=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Windows information (Windows Info)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
    SET @QTotalNo=@QTotalNo+1
    if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
    ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'win_release',	'win_service_pack_level',	'windows_sku',	'os_language_version')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', windows_release, windows_service_pack_level,   windows_sku, os_language_version FROM sys.dm_os_windows_info WITH (NOLOCK) OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN 
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	---6--
    IF(@Parm1_6=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Instance configuration setups'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	--Huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'config_id',	'name',		'minimum',	'maximum',	'Current value',	'description',	'is_dyn',	'is_adv', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'R',	'L',		'R',	'R',	'C',	'L',	'C',	'C', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', cast(configuration_id AS  VARCHAR),	cast(name	as  VARCHAR),  cast(minimum	as  VARCHAR), cast(maximum	as  VARCHAR),cast(value_in_use	as  VARCHAR), cast(description	as  VARCHAR), cast(is_dynamic	as  VARCHAR), cast(is_advanced  AS  VARCHAR),@ExcelSheetNo FROM sys.configurations WITH (NOLOCK)  ORDER BY name OPTION (RECOMPILE);

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')

	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
 IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7) 
		SELECT R1, R2, S1, R3,  R4,	C6, C10, C9 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (31,231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C6, C10, C9 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		--DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		IF @HugeDataCounter<=@AllowFewToReport
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5,C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5,C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		SET @ExcelSheetNo=@ExcelSheetNo+1
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	---7--
    IF(@Parm1_7=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='CPU VISIABLE ONLINE CHECK'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @OnlineCpuCount INT DECLARE  @LogicalCpuCount INT 
	SELECT @OnlineCpuCount = COUNT(*) FROM sys.dm_os_schedulers WHERE status = 'VISIBLE ONLINE'; SELECT @LogicalCpuCount = cpu_count FROM sys.dm_os_sys_info ;
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4,	C5,	C6,	C7) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'ASSIGNED ONLINE CPU #',	'VISIBLE ONLINE CPU #',	'CPU Usage Desc')
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4,	C5,	C6,	C7)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', @LogicalCpuCount AS 'ASSIGNED ONLINE CPU #', @OnlineCpuCount AS 'VISIBLE ONLINE CPU #',   CASE WHEN @OnlineCpuCount < @LogicalCpuCount THEN 'You are not using all CPU assigned to O/S! If it is VM, review your VM configuration to make sure you are not maxout Socket'     ELSE 'You are using all CPUs assigned to O/S. GOOD!'    END AS 'CPU Usage Desc'
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	---8--
    IF(@Parm1_8=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Physical, core and logical CPUs'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'number_of_physical_cpus',	'number_of_cores_per_cpu',	'total_number_of_cores',	'number_of_virtual_cpus',	'cpu_category')
	DECLARE  @xp_msver TABLE ([idx] [int] NULL, [c_name] VARCHAR(100) NULL, [int_val] [float] NULL,[c_val]  VARCHAR(128) NULL) INSERT INTO @xp_msver EXEC ('[master]..[xp_msver]');; WITH [ProcessorInfo] AS ( SELECT ([cpu_count] / [hyperthread_ratio]) AS [number_of_physical_cpus] 		,CASE  			WHEN hyperthread_ratio = cpu_count 				THEN cpu_count 			ELSE (([cpu_count] - [hyperthread_ratio]) / ([cpu_count] / [hyperthread_ratio])) 			END AS [number_of_cores_per_cpu] 		,CASE  			WHEN hyperthread_ratio = cpu_count 				THEN cpu_count 			ELSE ([cpu_count] / [hyperthread_ratio]) * (([cpu_count] - [hyperthread_ratio]) / ([cpu_count] / [hyperthread_ratio])) 			END AS [total_number_of_cores] 		,[cpu_count] AS [number_of_virtual_cpus] 		,( SELECT [c_val] 			FROM @xp_msver 			WHERE [c_name] = 'Platform' 			) AS [cpu_category] 	FROM [sys].[dm_os_sys_info] 	)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', [number_of_physical_cpus] 	,[number_of_cores_per_cpu] 	,[total_number_of_cores] 	,[number_of_virtual_cpus] 	,LTRIM(RIGHT([cpu_category], CHARINDEX('x', [cpu_category]) - 1)) AS [cpu_category] FROM [ProcessorInfo]
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
	----delete
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21, 60)
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	---9--
    IF(@Parm1_9=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Socket, physical core and logical core count from the SQL Server Error log.  (Core Counts)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'LogDate',	'ProcessInfo',	'Text')
	IF OBJECT_ID(N'tempdb..#temptbl3') IS NOT NULL
	BEGIN
	DROP TABLE #temptbl3
	END
	CREATE TABLE  #temptbl3(logdate  VARCHAR(max), procinfo  VARCHAR(max), textd  VARCHAR(max))
	INSERT INTO #temptbl3
	EXEC sys.xp_readerrorlog 0, 1, N'detected', N'socket';
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', * FROM #temptbl3
	drop table #temptbl3
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
		END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	---10--
    IF(@Parm1_10=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='CPU Count, Physical memory and other hardware information'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'cpu_count',	'Physical_mem_Gb',	'os_quantum',	'max_workers_count',	'scheduler_count',	'scheduler_total_count',	'sqlserver_start_time',	'virtual_machine_type',	'virtual_machine_type_desc',	'socket_count',	'cores_per_socket')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',  cpu_count, (physical_memory_kb/(1000)/1000) AS Physical_mem_Gb,  os_quantum, max_workers_count, scheduler_count, scheduler_total_count, sqlserver_start_time, virtual_machine_type, virtual_machine_type_desc,  socket_count, cores_per_socket	
	FROM [sys].[dm_os_sys_info]
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
		END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	---11--
    IF(@Parm1_11=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Hardware information from SQL Server  (Hardware Info)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) VALUES (@QSectionNo,	@QSectionSubNo, @QSectionSplitNo,  31, 'H',	'Logical cpu count',	'Scheduler count',	'Hyperthread ratio',	'Physical cpu count');
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', cpu_count AS [Logical CPU Count], scheduler_count,   hyperthread_ratio AS [Hyperthread Ratio], cpu_count/hyperthread_ratio AS [Physical CPU Count] FROM sys.dm_os_sys_info WITH (NOLOCK) OPTION (RECOMPILE);

	SET @QSectionSplitNo=@QSectionSplitNo+1
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) VALUES (@QSectionNo,	@QSectionSubNo, @QSectionSplitNo,  31, 'H',	'Physical memory (mb)',	'Committed memory (mb)',	'Committed target memory (mb)',	'Max workers count')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',  physical_memory_kb/1024 AS [Physical Memory (MB)], 	   committed_kb/1024 AS [Committed Memory (MB)],  committed_target_kb/1024 AS [Committed Target Memory (MB)],  max_workers_count AS [Max Workers Count]   FROM sys.dm_os_sys_info WITH (NOLOCK) OPTION (RECOMPILE);

	SET @QSectionSplitNo=@QSectionSplitNo+1
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) VALUES (@QSectionNo,	@QSectionSubNo, @QSectionSplitNo,  31, 'H','Max Worker Count',  'Sql server start time',	'Sql server up time (hrs)',	'Virtual machine type')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',   
	max_workers_count AS [Max Workers Count],   sqlserver_start_time AS [SQL Server Start Time],	   DATEDIFF(hour, sqlserver_start_time, GETDATE()) AS [SQL Server Up Time (hrs)],	   virtual_machine_type_desc AS [Virtual Machine Type] FROM sys.dm_os_sys_info WITH (NOLOCK) OPTION (RECOMPILE);


	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		IF(SELECT COUNT(C5) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=4 AND CAST(C5 AS INT)>CAST(C6 AS INT))>0
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5, H2)	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D','All CPUs should be scheduled', '2'
		ELSE
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5, H2)	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D','Good! All cpu are scheduled','1'	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	---12--
    IF(@Parm1_12=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Basic information about OS memory amounts and state (System Memory)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Physical Memory (MB)',	'Available Memory (MB)')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', total_physical_memory_kb/1024 AS [Physical Memory (MB)],   available_physical_memory_kb/1024 AS [Available Memory (MB)] FROM sys.dm_os_sys_memory WITH (NOLOCK) OPTION (RECOMPILE);

	SET @QSectionSplitNo=@QSectionSplitNo+1
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H', 'Total Page File (MB)',	'Available Page File (MB)')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', total_page_file_kb/1024 AS [Total Page File (MB)], 	   available_page_file_kb/1024 AS [Available Page File (MB)] FROM sys.dm_os_sys_memory WITH (NOLOCK) OPTION (RECOMPILE);

	SET @QSectionSplitNo=@QSectionSplitNo+1
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6 ) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'System Cache (MB)',	'System Memory State')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',  system_cache_kb/1024 AS [System Cache (MB)],  system_memory_state_desc AS [System Memory State] FROM sys.dm_os_sys_memory WITH (NOLOCK) OPTION (RECOMPILE);

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, H2) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D',	C10, 
		CASE WHEN C10='Available physical memory is high' THEN '1' Else '2' END FROM
			##RunOnce WHERE R1=@QSectionNo AND R2= @QSectionSubNo AND R3=4

		---Logice goes here	
					INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	---13--
    IF(@Parm1_13=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Volume info for all LUNS that have database files on the current instance (Volume Info)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,c9,c10) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'volume_mount_point',	'file_system_type',	'logical_volume_name',	'Total Size (GB)',	'Available Size (GB)',	'Space Free %')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8, C9, C10)
	SELECT DISTINCT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', vs.volume_mount_point, vs.file_system_type, vs.logical_volume_name,  CONVERT(DECIMAL(18,2), vs.total_bytes/1073741824.0) AS [Total Size (GB)], CONVERT(DECIMAL(18,2), vs.available_bytes/1073741824.0) AS [Available Size (GB)],   CONVERT(DECIMAL(18,2), vs.available_bytes * 1. / vs.total_bytes * 100.) AS [Space Free %] FROM sys.master_files AS f WITH (NOLOCK) CROSS APPLY sys.dm_os_volume_stats(f.database_id, f.[file_id]) AS vs   ORDER BY vs.volume_mount_point OPTION (RECOMPILE);


	SET @QSectionSplitNo=@QSectionSplitNo+1
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9, c10) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'volume_mount_point','logical_volume_name', 'supports_compression',	'is_compressed',	'supports_sparse_files',	'supports_alternate_streams')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,c10)
	SELECT DISTINCT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', vs.volume_mount_point,  vs.logical_volume_name,   vs.supports_compression, vs.is_compressed,  vs.supports_sparse_files, vs.supports_alternate_streams FROM sys.master_files AS f WITH (NOLOCK) CROSS APPLY sys.dm_os_volume_stats(f.database_id, f.[file_id]) AS vs   ORDER BY vs.volume_mount_point  OPTION (RECOMPILE);


	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, H2) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D',	'Being low on free space can negatively affect performance','0')
		--1. Check if space is enough are tight
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, H2) 
		SELECT  @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D', 
		CASE WHEN CONVERT(DECIMAL(18,2), C10) > 50 THEN 'Enough space available on ' +C5   
		WHEN CONVERT(DECIMAL(18,2), C10) < 15 THEN 'Capacity is tight enough on ' +C5 END,   
		CASE WHEN CONVERT(DECIMAL(18,2), C10) > 50 THEN '1'   
		WHEN CONVERT(DECIMAL(18,2), C10) < 15 THEN '2' END  
		FROM ##RunOnce WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=4 and R4='D'

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	---14--
    IF(@Parm1_14=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='SQL Server NUMA Node information (SQL Server NUMA Info)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'node_desc',	'memory_node_id',	'processor_group',	'online_scheduler_count',	'idle_scheduler_count',	'active_worker_count',	'avg_load_balance',	'resource_monitor_state')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', cast(node_id as varchar)+'. '+ node_state_desc, memory_node_id, processor_group, online_scheduler_count,   idle_scheduler_count, active_worker_count, avg_load_balance, resource_monitor_state FROM sys.dm_os_nodes WITH (NOLOCK)  WHERE node_state_desc <> N'ONLINE DAC' OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	----15--
	IF(@Parm1_15=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Disk Space Monitoring'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Disk Mount Point',	'File System Type',	'Logical Drive Name',	'Total Size in GB',	'Available Size in GB',	'Space Free %')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'R',	'R',	'R')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10)
	SELECT DISTINCT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', volume_mount_point [Disk Mount Point],  file_system_type [File System Type],  logical_volume_name AS [Logical Drive Name],  CONVERT(DECIMAL(18,2),total_bytes/1073741824.0) AS [Total Size in GB],  CONVERT(DECIMAL(18,2),available_bytes/1073741824.0) AS [Available Size in GB],   CAST(CAST(available_bytes AS FLOAT)/ CAST(total_bytes AS FLOAT) AS DECIMAL(18,2)) * 100 AS [Space Free %]  FROM sys.master_files  CROSS APPLY sys.dm_os_volume_stats(database_id, file_id)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----16---
    IF(@Parm1_16=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='How much percent of space used by log files. log file percentage'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')

	IF OBJECT_ID(N'tempdb..#temptbl2') IS NOT NULL
	BEGIN
	DROP TABLE #temptbl2
	END
	CREATE TABLE  #temptbl2(dname  VARCHAR(max), logsize  VARCHAR(max), sused  VARCHAR(max), status  VARCHAR(max))
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'R',	'R',	'L')
	INSERT INTO #temptbl2
	exec ('DBCC sqlperf(logspace)')
	--huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Log Size (MB)',	'Log Space Used (%)',	'Status', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'R',	'R',	'C', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM #temptbl2
	drop table #temptbl2
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 -- AND  
	BEGIN --
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 = 32 
		ORDER BY C7 DESC
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		if(@HugeDataCounter>@AllowFewToReport)
		BEGIN 
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
			Set @ExcelSheetNo=@ExcelSheetNo+1;
		END
		ELSE
			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	----17----
    IF(@Parm1_17=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='information about your cluster nodes and their status (Cluster Node Properties)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'NodeName',	'status_description',	'is_current_owner')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', NodeName, status_description, is_current_owner FROM sys.dm_os_cluster_nodes WITH (NOLOCK) OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/

	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	---18--
    IF(@Parm1_18=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Information about any AlwaysOn AG cluster on this instance'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'cluster_name',	'quorum_type_desc',	'quorum_state_desc')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', cluster_name, quorum_type_desc, quorum_state_desc FROM sys.dm_hadr_cluster WITH (NOLOCK) OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	---19--
    IF(@Parm1_19=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Overview of AG health and status'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20,	C21,	C22,	C23,	C24,	C25,	C26,	C27,	C28,	C29,	C30,	C31,	C32) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'AG Name',	'replica_server_name',	'availability_mode_desc',	'database_name',	'is_local',	'is_primary_replica',	'synchronization_state_desc',	'is_commit_participant',	'synchronization_health_desc',	'recovery_lsn',	'truncation_lsn',	'last_sent_lsn',	'last_sent_time',	'last_received_lsn',	'last_received_time',	'last_hardened_lsn',	'last_hardened_time',	'last_redone_lsn',	'last_redone_time',	'log_send_queue_size',	'log_send_rate',	'redo_queue_size',	'redo_rate',	'filestream_send_rate',	'end_of_log_lsn',	'last_commit_lsn',	'last_commit_time',	'database_state_desc')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20,	C21,	C22,	C23,	C24,	C25,	C26,	C27,	C28,	C29,	C30,	C31,	C32)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', ag.name AS [AG Name], ar.replica_server_name, ar.availability_mode_desc, adc.[database_name],   drs.is_local, drs.is_primary_replica, drs.synchronization_state_desc, drs.is_commit_participant, 	   drs.synchronization_health_desc, drs.recovery_lsn, drs.truncation_lsn, drs.last_sent_lsn, 	   drs.last_sent_time, drs.last_received_lsn, drs.last_received_time, drs.last_hardened_lsn, 	   drs.last_hardened_time, drs.last_redone_lsn, drs.last_redone_time, drs.log_send_queue_size, 	   drs.log_send_rate, drs.redo_queue_size, drs.redo_rate, drs.filestream_send_rate, 	   drs.end_of_log_lsn, drs.last_commit_lsn, drs.last_commit_time, drs.database_state_desc   FROM sys.dm_hadr_database_replica_states AS drs WITH (NOLOCK)  INNER JOIN sys.availability_databases_cluster AS adc WITH (NOLOCK)  ON drs.group_id = adc.group_id   AND drs.group_database_id = adc.group_database_id  INNER JOIN sys.availability_groups AS ag WITH (NOLOCK)  ON ag.group_id = drs.group_id  INNER JOIN sys.availability_replicas AS ar WITH (NOLOCK)  ON drs.group_id = ar.group_id   AND drs.replica_id = ar.replica_id  ORDER BY ag.name, ar.replica_server_name, adc.[database_name] OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	---20--
    IF(@Parm1_20=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Database Instant File Initialization'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,C5,C6,C7) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H', 'logDate', 'processinfo', 'text')
	DECLARE @Db_IFI TABLE(logDate varchar(max), processinfo varchar(max), text varchar(max))
	INSERT INTO @Db_IFI 
	EXEC sys.xp_readerrorlog 0, 1, N'Database Instant File Initialization';
	update @Db_IFI set [text]=[text]+' <a href="https://learn.microsoft.com/en-us/sql/relational-databases/databases/database-instant-file-initialization?view=sql-server-ver16"> (Learn more)</a>'
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, C6, C7) 
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', * FROM @Db_IFI 
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	----21---
    IF(@Parm1_21=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='SQL Server Process Address space info (Process Memory)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8)VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Memory Usage (MB)',	'Locked Pages Alloc (MB)',	'Large Pages Alloc (MB)',	'page_fault_count')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', physical_memory_in_use_kb/1024 AS [SQL Server Memory Usage (MB)], locked_page_allocations_kb/1024 AS [SQL Server Locked Pages Allocation (MB)], large_page_allocations_kb/1024 AS [SQL Server Large Pages Allocation (MB)], 	   page_fault_count FROM sys.dm_os_process_memory WITH (NOLOCK) OPTION (RECOMPILE);
	SET @QSectionSplitNo=@QSectionSplitNo+1
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'mem_utilization_%',	'available_commit_limit_kb',	'process_physical_memory_low',	'process_virtual_memory_low')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',  memory_utilization_percentage, available_commit_limit_kb, process_physical_memory_low, process_virtual_memory_low FROM sys.dm_os_process_memory WITH (NOLOCK) OPTION (RECOMPILE);

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm

	---22--
    IF(@Parm1_22=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Resource Governor Resource Pool information (RG Resource Pools)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'pool_id',	'Name',	'statistics_start_time',	'[min-max] memory %',	'max_memory_mb',	'used_memory_mb',	'target_memory_mb',	'[min - max] iops_per_volume')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'R',	'R',	'R',	'R',	'R',	'R')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12) 
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', pool_id, [Name], CONVERT(char(10), statistics_start_time,126),  cast(min_memory_percent as varchar)+'% -'+cast(max_memory_percent as varchar)+'%',     max_memory_kb/1024 AS [max_memory_mb],     used_memory_kb/1024 AS [used_memory_mb],      target_memory_kb/1024 AS [target_memory_mb], min_iops_per_volume+'-'+ max_iops_per_volume FROM sys.dm_resource_governor_resource_pools WITH (NOLOCK) OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
END
if(@Hardware=0)
BEGIN
	SET @QSectionNo=@QSectionNo+1
	SET @QSectionSubNo=1
	SET @QSectionSplitNo=1
	SET @QTotalNo=@QTotalNo+1
    Set @QHeadingOne='Harding configuration setting'
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, -1, 'H1', CAST(@QSectionNo AS VARCHAR)+'. '+@QHeadingOne) 
	----23----
    IF(@Parm2_1=0)
	BEGIN
	BEGIN TRY 
	SET @QHeadingTw0='Processor description from Windows Registry (Processor Description)'
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, C6) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H','VALUE','DATA')
	DECLARE @ProcessorName TABLE(value varchar(max), data varchar(max))
	INSERT INTO @ProcessorName
	EXEC sys.xp_instance_regread N'HKEY_LOCAL_MACHINE', N'HARDWARE\DESCRIPTION\System\CentralProcessor\0', N'ProcessorNameString';
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, C6) 
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',  * FROM @ProcessorName
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')	
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	---24--
    IF(@Parm2_2=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='BIOS date from Windows Registry (BIOS Date)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE @BiosRelease table(value varchar(max), date varchar(max))
	insert into @BiosRelease
	EXEC sys.xp_instance_regread N'HKEY_LOCAL_MACHINE', N'HARDWARE\DESCRIPTION\System\BIOS', N'BiosReleaseDate';
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H', 'VALUE','DATA')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',* FROM @BiosRelease
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	---25--
    IF(@Parm2_3=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='System Manufacturer and model number from SQL Server Error log (System Manufacturer)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7)VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'LogDate',	'ProcessInfo',	'Text')
	IF OBJECT_ID(N'tempdb..#temptbl4') IS NOT NULL
	BEGIN
	DROP TABLE #temptbl4
	END
	CREATE TABLE  #temptbl4(R1  VARCHAR(max), R2  VARCHAR(max), R3  VARCHAR(max))
	INSERT INTO #temptbl4
	EXEC sys.xp_readerrorlog 0, 1, N'Manufacturer';
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', * FROM #temptbl4
	drop table #temptbl4
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	---26--
    IF(@Parm2_4=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='pvscsi info from Windows Registry (PVSCSI Driver Parameters)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	EXEC sys.xp_instance_regread N'HKEY_LOCAL_MACHINE', N'SYSTEM\CurrentControlSet\services\pvscsi\Parameters\Device', N'DriverParameter';
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm

	---27--
    IF(@Parm2_5=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Page Life Expectancy (PLE) value for each NUMA node in current instance (PLE by NUMA Node)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Server Name',	'Object',	'instance_name',	'Page Life Expectancy')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'R')
	INSERT INTO ##RunOnce(R1, R2, S1, QuerySort, R3, R4,	C5,	C6,	C7,	C8)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, ROW_NUMBER() OVER(ORDER BY cntr_value DESC), 32, 'D', @@SERVERNAME AS [Server Name], RTRIM([object_name]) AS [Object],    instance_name, cntr_value AS [Page Life Expectancy] FROM sys.dm_os_performance_counters WITH (NOLOCK) WHERE [object_name] LIKE N'%Buffer Node%' -- Handles named instances AND counter_name = N'Page life expectancy' OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	---28--
    IF(@Parm2_6=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='CPUs schedulers Visible Online'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'scheduler_id',	'cpu_id',	'status',	'is_online')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',  scheduler_id, cpu_id, status, is_online FROM sys.dm_os_schedulers 
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	---29--
    IF(@Parm2_7=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='CPU Utilization History for last 256 minutes (in one minute intervals)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @ts_now bigint = (SELECT cpu_ticks/(cpu_ticks/ms_ticks) FROM sys.dm_os_sys_info WITH (NOLOCK));  
	INSERT INTO  ##RunOnce(R1, R2, S1, R3, R4, C5, C6) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H', 'Event Time', '[SQL Server Process CPU Utilization],  [System Idle Process], [Other Process CPU Utilization]');
	INSERT INTO  ##RunOnce(R1, R2, S1, R3, R4, C5, C6) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S', 'L', 'R');
	INSERT INTO  ##RunOnce(R1, R2, S1, R3, R4, C5, C6) 
	SELECT TOP(10) @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', DATEADD(ms, -1 * (@ts_now - [timestamp]), GETDATE()) AS [Event Time], Cast(SQLProcessUtilization as varchar)+',&nbsp&nbsp&nbsp&nbsp'+ CAST(SystemIdle AS VARCHAR)+',&nbsp&nbsp&nbsp&nbsp'+ CAST(100 - SystemIdle - SQLProcessUtilization AS VARCHAR) FROM (SELECT record.value('(./Record/@id)[1]', 'int') AS record_id, record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]', 'int') AS [SystemIdle], record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]', 'int') AS [SQLProcessUtilization], [timestamp]  	  FROM (SELECT [timestamp], CONVERT(xml, record) AS [record]  			FROM sys.dm_os_ring_buffers WITH (NOLOCK) 			WHERE ring_buffer_type = N'RING_BUFFER_SCHEDULER_MONITOR'  			AND record LIKE N'%<SystemHealth>%') AS x) AS y   ORDER BY record_id DESC OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
END -- END IF @Hardware check

IF(@SQL_Server_Instance_Engine=0)
BEGIN
	SET @QSectionNo=@QSectionNo+1
	SET @QSectionSubNo=1
	SET @QSectionSplitNo=1
	SET @QTotalNo=@QTotalNo+1
    Set @QHeadingOne='SQL server running instance configurations'
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, -1, 'H1', CAST(@QSectionNo AS VARCHAR)+'. '+@QHeadingOne) 
	---30--
    IF(@Parm3_1=0)
	BEGIN
	BEGIN TRY 
    SET @QHeadingTw0='Look at Suspect Pages Suspect Pages)'
    SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,C5, C6, C7, C8, C9, C10) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H', 'Database','file_id','page_id','event_type','error_count','last_update_date')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,C5, C6, C7, C8, C9, C10) 
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', DB_NAME(database_id) AS [Database], [file_id], page_id,    event_type, error_count, CONVERT(VARCHAR, last_update_date, 20)  FROM msdb.dbo.suspect_pages WITH (NOLOCK)  ORDER BY database_id OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
    if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
    BEGIN
    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
    END
    END TRY
    BEGIN CATCH
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo,0, 'E',CAST(ERROR_NUMBER() AS VARCHAR)+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----31----
    IF(@Parm3_2=0)
	BEGIN
	BEGIN TRY 
    SET @QHeadingTw0='information on location, time and size of any memory dumps from SQL Server (Memory Dump Info)'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
    SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,C5, C6, C7) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H', 'filename', 'creation_time','size(MB)')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,C5, C6, C7) 
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', [filename], CONVERT(VARCHAR, creation_time, 20) , size_in_bytes/1048576.0 AS [Size (MB)] FROM sys.dm_server_memory_dumps WITH (NOLOCK)   ORDER BY creation_time DESC OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
    if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
    BEGIN
    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
    END
    END TRY
    BEGIN CATCH
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	----32---
    IF(@Parm3_3=0)
	BEGIN
	BEGIN TRY 
    SET @QHeadingTw0='Detect blocking'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
    SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'lock type',	'database',	'blk object',	'lock req',	'waiter sid',	'wait time',	'waiter_batch',	'waiter_stmt',	'blocker sid',	'blocker_batch')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', t1.resource_type AS [lock type], DB_NAME(resource_database_id) AS [database], 	t1.resource_associated_entity_id AS [blk object],t1.request_mode AS [lock req], t1.request_session_id AS [waiter sid], t2.wait_duration_ms AS [wait time], (SELECT [text] FROM sys.dm_exec_requests AS r WITH (NOLOCK)        	CROSS APPLY sys.dm_exec_sql_text(r.[sql_handle])  	WHERE r.session_id = t1.request_session_id) AS [waiter_batch], 	(SELECT SUBSTRING(qt.[text],r.statement_start_offset/2,  		(CASE WHEN r.statement_end_offset = -1  		THEN LEN(CONVERT(NVARCHAR(max), qt.[text])) * 2  		ELSE r.statement_end_offset END - r.statement_start_offset)/2)  	FROM sys.dm_exec_requests AS r WITH (NOLOCK) 	CROSS APPLY sys.dm_exec_sql_text(r.[sql_handle]) AS qt 	WHERE r.session_id = t1.request_session_id) AS [waiter_stmt],					 	t2.blocking_session_id AS [blocker sid],										 	(SELECT [text] FROM sys.sysprocesses AS p										 	CROSS APPLY sys.dm_exec_sql_text(p.[sql_handle])  	WHERE p.spid = t2.blocking_session_id) AS [blocker_batch] 	FROM sys.dm_tran_locks AS t1 WITH (NOLOCK) 	INNER JOIN sys.dm_os_waiting_tasks AS t2 WITH (NOLOCK) 	ON t1.lock_owner_address = t2.resource_address OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
    if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
    BEGIN
    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
    END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
  
	---33--
    IF(@Parm3_4=0)
	BEGIN
	BEGIN TRY	
	SET @QHeadingTw0='Isolate top waits for server instance since last restart or wait statistics clear'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
    SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'WaitType',	'Wait(%)',	'AvgWait_Sec',	'AvgRes_Sec',	'AvgSig_Sec',	'Wait_Sec',	'Resource_Sec',	'Signal_Sec',	'Wait Count',	'Help/Info URL');
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'L');
    WITH [Waits]  AS (SELECT wait_type, wait_time_ms/ 1000.0 AS [WaitS],      (wait_time_ms - signal_wait_time_ms) / 1000.0 AS [ResourceS],  signal_wait_time_ms / 1000.0 AS [SignalS],  waiting_tasks_count AS [WaitCount],  100.0 *  wait_time_ms / SUM (wait_time_ms) OVER() AS [Percentage],  ROW_NUMBER() OVER( ORDER BY wait_time_ms DESC) AS [RowNum]     FROM sys.dm_os_wait_stats WITH (NOLOCK)     WHERE [wait_type] NOT IN (    N'BROKER_EVENTHANDLER', N'BROKER_RECEIVE_WAITFOR', N'BROKER_TASK_STOP', 		N'BROKER_TO_FLUSH', N'BROKER_TRANSMITTER', N'CHECKPOINT_QUEUE',    N'CHKPT', N'CLR_AUTO_EVENT', N'CLR_MANUAL_EVENT', N'CLR_SEMAPHORE',    N'DBMIRROR_DBM_EVENT', N'DBMIRROR_EVENTS_QUEUE', N'DBMIRROR_WORKER_QUEUE', 		N'DBMIRRORING_CMD', N'DIRTY_PAGE_POLL', N'DISPATCHER_QUEUE_SEMAPHORE',    N'EXECSYNC', N'FSAGENT', N'FT_IFTS_SCHEDULER_IDLE_WAIT', N'FT_IFTSHC_MUTEX',    N'HADR_CLUSAPI_CALL', N'HADR_FILESTREAM_IOMGR_IOCOMPLETION', N'HADR_LOGCAPTURE_WAIT',  		N'HADR_NOTIFICATION_DEQUEUE', N'HADR_TIMER_TASK', N'HADR_WORK_QUEUE',    N'KSOURCE_WAKEUP', N'LAZYWRITER_SLEEP', N'LOGMGR_QUEUE', N'ONDEMAND_TASK_QUEUE',    N'PWAIT_ALL_COMPONENTS_INITIALIZED',  		N'PREEMPTIVE_OS_AUTHENTICATIONOPS', N'PREEMPTIVE_OS_CREATEFILE', N'PREEMPTIVE_OS_GENERICOPS', 		N'PREEMPTIVE_OS_LIBRARYOPS', N'PREEMPTIVE_OS_QUERYREGISTRY', 		N'PREEMPTIVE_HADR_LEASE_MECHANISM', N'PREEMPTIVE_SP_SERVER_DIAGNOSTICS', 		N'QDS_PERSIST_TASK_MAIN_LOOP_SLEEP',    N'QDS_CLEANUP_STALE_QUERIES_TASK_MAIN_LOOP_SLEEP', N'QDS_SHUTDOWN_QUEUE', N'REQUEST_FOR_DEADLOCK_SEARCH', 		N'RESOURCE_QUEUE', N'SERVER_IDLE_CHECK', N'SLEEP_BPOOL_FLUSH', N'SLEEP_DBSTARTUP', 		N'SLEEP_DCOMSTARTUP', N'SLEEP_MASTERDBREADY', N'SLEEP_MASTERMDREADY',    N'SLEEP_MASTERUPGRADED', N'SLEEP_MSDBSTARTUP', N'SLEEP_SYSTEMTASK', N'SLEEP_TASK',    N'SLEEP_TEMPDBSTARTUP', N'SNI_HTTP_ACCEPT', N'SP_SERVER_DIAGNOSTICS_SLEEP', 		N'SQLTRACE_BUFFER_FLUSH', N'SQLTRACE_INCREMENTAL_FLUSH_SLEEP', N'SQLTRACE_WAIT_ENTRIES', 		N'WAIT_FOR_RESULTS', N'WAITFOR', N'WAITFOR_TASKSHUTDOWN', N'WAIT_XTP_HOST_WAIT', 		N'WAIT_XTP_OFFLINE_CKPT_NEW_LOG', N'WAIT_XTP_CKPT_CLOSE', N'XE_DISPATCHER_JOIN',    N'XE_DISPATCHER_WAIT', N'XE_TIMER_EVENT')     AND waiting_tasks_count > 0) 
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',  MAX (W1.wait_type) AS [WaitType], 	CAST (MAX (W1.Percentage) AS DECIMAL (5,2)) AS [Wait Percentage], 	CAST ((MAX (W1.WaitS) / MAX (W1.WaitCount)) AS DECIMAL (16,4)) AS [AvgWait_Sec],     CAST ((MAX (W1.ResourceS) / MAX (W1.WaitCount)) AS DECIMAL (16,4)) AS [AvgRes_Sec],     CAST ((MAX (W1.SignalS) / MAX (W1.WaitCount)) AS DECIMAL (16,4)) AS [AvgSig_Sec], CAST (MAX (W1.WaitS) AS DECIMAL (16,2)) AS [Wait_Sec],     CAST (MAX (W1.ResourceS) AS DECIMAL (16,2)) AS [Resource_Sec],     CAST (MAX (W1.SignalS) AS DECIMAL (16,2)) AS [Signal_Sec],     MAX (W1.WaitCount) AS [Wait Count], 	CAST (N'https://www.sqlskills.com/help/waits/' + W1.wait_type AS  VARCHAR) AS [Help/Info URL] FROM Waits AS W1 INNER JOIN Waits AS W2 ON W2.RowNum <= W1.RowNum GROUP BY W1.RowNum, W1.wait_type HAVING SUM (W2.Percentage) - MAX (W1.Percentage) < 99  OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
    if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
    BEGIN
    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
    END
    END TRY
    BEGIN CATCH
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	----34---
	IF(@Parm3_5=0)
	BEGIN
    BEGIN TRY 
	SET @QHeadingTw0='Memory Clerk Usage for instance'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
    SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Memory Clerk Type',	'Memory Usage (MB)')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'Memory Clerk Type',	'R')
    INSERT INTO ##RunOnce(R1, R2, S1, QuerySort, R3, R4,	C5,	C6)
    SELECT TOP(10)  @QSectionNo, @QSectionSubNo, @QSectionSplitNo, ROW_NUMBER() OVER(ORDER BY SUM(mc.pages_kb) DESC), 32, 'D', mc.[type] AS [Memory Clerk Type],   CAST((SUM(mc.pages_kb)/1024.0) AS DECIMAL (15,2)) AS [Memory Usage (MB)] FROM sys.dm_os_memory_clerks AS mc WITH (NOLOCK)GROUP BY mc.[type]   ORDER BY SUM(mc.pages_kb) DESC OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
    if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
    BEGIN
    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
    END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	----35---
    IF(@Parm3_6=0)
	BEGIN
	BEGIN TRY 
	SET @QHeadingTw0='Average Task Counts'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
    SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Avg Task Count',	'Avg Work Queue Count',	'Avg Runnable Task Count',	'Avg Pending DiskIO Count')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', AVG(current_tasks_count) AS [Avg Task Count],  AVG(work_queue_count) AS [Avg Work Queue Count], AVG(runnable_tasks_count) AS [Avg Runnable Task Count], AVG(pending_disk_io_count) AS [Avg Pending DiskIO Count] FROM sys.dm_os_schedulers WITH (NOLOCK) WHERE scheduler_id < 255 OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
    if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
    BEGIN
    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
    END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	----36---
    IF(@Parm3_7=0)
	BEGIN
	BEGIN TRY 
    SET @QHeadingTw0='Look for I/O requests taking longer than 15 seconds in the six most recent SQL Server Error Logs (IO Warnings)'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    CREATE TABLE  #IOWarningResults(LogDate datetime, ProcessInfo sysname, LogText NVARCHAR(1000)); 	
    INSERT INTO #IOWarningResults  	EXEC xp_readerrorlog 0, 1, N'taking longer than 15 seconds'; 
    INSERT INTO #IOWarningResults  	EXEC xp_readerrorlog 1, 1, N'taking longer than 15 seconds';	
    INSERT INTO #IOWarningResults  	EXEC xp_readerrorlog 2, 1, N'taking longer than 15 seconds';	
    INSERT INTO #IOWarningResults  	EXEC xp_readerrorlog 3, 1, N'taking longer than 15 seconds'; 	
    INSERT INTO #IOWarningResults 	EXEC xp_readerrorlog 4, 1, N'taking longer than 15 seconds';	
    INSERT INTO #IOWarningResults 	EXEC xp_readerrorlog 5, 1, N'taking longer than 15 seconds'; 
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'LogDate',	'ProcessInfo',	'LogText')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', * FROM #IOWarningResults  ORDER BY LogDate DESC; DROP TABLE #IOWarningResults;
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
    if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
    BEGIN
    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
    END
    END TRY
    BEGIN CATCH
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	----37---
	IF(@Parm3_8=0)
	BEGIN
    BEGIN TRY 
    SET @QHeadingTw0='Memory Grants Pending value for current instance'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Server Name',	'Object',	'Memory Grants Pending')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'R')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', @@SERVERNAME AS [Server Name], RTRIM([object_name]) AS [Object], cntr_value AS [Memory Grants Pending]  FROM sys.dm_os_performance_counters WITH (NOLOCK) WHERE [object_name] LIKE N'%Memory Manager%'   ORDER BY cntr_value desc -- Handles named instances AND counter_name = N'Memory Grants Pending' OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
    if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
    BEGIN
    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
    END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
    
	---38--
	IF(@Parm3_9=0)
	BEGIN
    BEGIN TRY 
    SET @QHeadingTw0='Count of SQL connections by IP address '
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'client_net_address',	'program_name',	'host_name',	'login_name',	'connection count');
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', ec.client_net_address, es.[program_name], es.[host_name], es.login_name,  COUNT(ec.session_id) AS [connection count]  FROM sys.dm_exec_sessions AS es WITH (NOLOCK)  INNER JOIN sys.dm_exec_connections AS ec WITH (NOLOCK)  ON es.session_id = ec.session_id  GROUP BY ec.client_net_address, es.[program_name], es.[host_name], es.login_name    ORDER BY ec.client_net_address, es.[program_name] OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
    if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
    BEGIN
    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
    END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
    
	----39---
	IF(@Parm3_10=0)
	BEGIN
    BEGIN TRY 
    SET @QHeadingTw0='SQL connections by IP address (Connection Counts by IP Address)'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20,	C21,	C22,	C23,	C24,	C25,	C26) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'session_id',	'most_recent_session_id',	'connect_time',	'net_transport',	'protocol_type',	'protocol_version',	'endpoint_id',	'encrypt_option',	'auth_scheme',	'node_affinity',	'num_reads',	'num_writes',	'last_read',	'last_write',	'net_packet_size',	'client_net_address',	'client_tcp_port',	'local_net_address',	'local_tcp_port',	'connection_id',	'parent_connection_id',	'most_recent_sql_handle');
    INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20,	C21,	C22,	C23,	C24,	C25,	C26)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', * FROM sys.dm_exec_connections
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
    if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
    BEGIN
    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
    END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
END -- END OF IF(@SQL_Server_Instance_Engine=0)

IF(@SQL_Agent_Jobs_information =0)
BEGIN
	SET @QSectionNo=@QSectionNo+1
	SET @QSectionSubNo=1
	SET @QSectionSplitNo=1
	SET @QTotalNo=@QTotalNo+1
    Set @QHeadingOne='SQL agent jobs information'
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, -1, 'H1', CAST(@QSectionNo AS VARCHAR)+'. '+@QHeadingOne) 

	----40---
    IF(@Parm4_1=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='SQL Server Agent jobs and Category information'
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	--huge data expected	done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Name',	'Description',	'Owner',	'Date Created',	'Enabled',	'email_operator',	'notify_level',	'Category Name',	'Sched Enabled',	'next_run_date',	'next_run_time', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'R',	'C',	'R',	'C',	'L',	'C',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15, ExclSheet)
	--splite data
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', sj.name AS [Name], sj.[description] AS [Description], SUSER_SNAME(sj.owner_sid) AS [Owner], sj.date_created AS [Date Created], sj.[enabled] AS [Job Enabled], sj.notify_email_operator_id, sj.notify_level_email, sc.name AS [CategoryName], s.[enabled] AS [Sched Enabled], CONVERT(VARCHAR, js.next_run_date, 20), js.next_run_time, @ExcelSheetNo FROM msdb.dbo.sysjobs AS sj WITH (NOLOCK) INNER JOIN msdb.dbo.syscategories AS sc WITH (NOLOCK) ON sj.category_id = sc.category_id LEFT OUTER JOIN msdb.dbo.sysjobschedules AS js WITH (NOLOCK) ON sj.job_id = js.job_id LEFT OUTER JOIN msdb.dbo.sysschedules AS s WITH (NOLOCK) ON js.schedule_id = s.schedule_id  ORDER BY sj.name OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
 IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14,	C15) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14,	C15 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14,	C15) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14,	C15 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
		IF @HugeDataCounter>@AllowFewToReport
		BEGIN
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
			Set @ExcelSheetNo=@ExcelSheetNo+1;
		END
		ELSE
			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END

    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	---41--
    IF(@Parm4_2=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='SQL Server Agent Jobs enable status, frequency'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	--huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'name',	'enabled',	'freq_type',	'freq_interval',	'freq_subday_type',	'freq_subday_interval',	'freq_recurrence_factor', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'C',	'C',	'R',	'R',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', name ,enabled ,freq_type ,freq_interval ,freq_subday_type ,freq_subday_interval ,freq_recurrence_factor, @ExcelSheetNo FROM msdb.dbo.sysschedules
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
 IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11) 
		SELECT R1, R2, s1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11) 
		SELECT TOP(@AllowFewToReport) R1, R2, s1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
		IF @HugeDataCounter>@AllowFewToReport
			BEGIN
				INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
				Set @ExcelSheetNo=@ExcelSheetNo+1;
			END
		ELSE
			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END

    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	----42---
    IF(@Parm4_3=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='List of jobs as per schedules'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11)  	VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'name',	'enabled',	'schedule_name',	'freq_recurrence_factor',	'frequency',	'Days',	'time')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',  sysjobs.name job_name ,sysjobs.enabled job_enabled ,sysschedules.name schedule_name ,sysschedules.freq_recurrence_factor ,case  WHEN freq_type = 4 THEN 'Daily' end frequency , 'every ' + cast (freq_interval AS  VARCHAR(3)) + ' day(s)'  Days , CASE  WHEN freq_subday_type = 2 THEN ' every ' + cast(freq_subday_interval AS  VARCHAR(7))   + ' seconds' + ' starting at '  + stuff(stuff(RIGHT(replicate('0', 6) +  cast(active_start_time AS  VARCHAR(6)), 6), 3, 0, ':'), 6, 0, ':')  WHEN freq_subday_type = 4 THEN ' every ' + cast(freq_subday_interval AS  VARCHAR(7))   + ' minutes' + ' starting at '  + stuff(stuff(RIGHT(replicate('0', 6) +  cast(active_start_time AS  VARCHAR(6)), 6), 3, 0, ':'), 6, 0, ':')  WHEN freq_subday_type = 8 THEN ' every ' + cast(freq_subday_interval AS  VARCHAR(7))   + ' hours'   + ' starting at '  + stuff(stuff(RIGHT(replicate('0', 6) +  cast(active_start_time AS  VARCHAR(6)), 6), 3, 0, ':'), 6, 0, ':')  else ' starting at '   +stuff(stuff(RIGHT(replicate('0', 6) +  cast(active_start_time AS  VARCHAR(6)), 6), 3, 0, ':'), 6, 0, ':') end time FROM msdb.dbo.sysjobs inner join msdb.dbo.sysjobschedules on sysjobs.job_id = sysjobschedules.job_id inner join msdb.dbo.sysschedules on sysjobschedules.schedule_id = sysschedules.schedule_id where freq_type = 4 
	union 
	SELECT  @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', sysjobs.name job_name ,sysjobs.enabled job_enabled ,sysschedules.name schedule_name ,sysschedules.freq_recurrence_factor ,case  WHEN freq_type = 8 THEN 'Weekly' end frequency , replace (  CASE WHEN freq_interval&1 = 1 THEN 'Sunday, ' ELSE '' END +CASE WHEN freq_interval&2 = 2 THEN 'Monday, ' ELSE '' END +CASE WHEN freq_interval&4 = 4 THEN 'Tuesday, ' ELSE '' END +CASE WHEN freq_interval&8 = 8 THEN 'Wednesday, ' ELSE '' END +CASE WHEN freq_interval&16 = 16 THEN 'Thursday, ' ELSE '' END +CASE WHEN freq_interval&32 = 32 THEN 'Friday, ' ELSE '' END +CASE WHEN freq_interval&64 = 64 THEN 'Saturday, ' ELSE '' END ,', ' ,'' ) Days , CASE  WHEN freq_subday_type = 2 THEN ' every ' + cast(freq_subday_interval AS  VARCHAR(7))   + ' seconds' + ' starting at '  + stuff(stuff(RIGHT(replicate('0', 6) +  cast(active_start_time AS  VARCHAR(6)), 6), 3, 0, ':'), 6, 0, ':')   WHEN freq_subday_type = 4 THEN ' every ' + cast(freq_subday_interval AS  VARCHAR(7))   + ' minutes' + ' starting at '  + stuff(stuff(RIGHT(replicate('0', 6) +  cast(active_start_time AS  VARCHAR(6)), 6), 3, 0, ':'), 6, 0, ':')  WHEN freq_subday_type = 8 THEN ' every ' + cast(freq_subday_interval AS  VARCHAR(7))   + ' hours'   + ' starting at '  + stuff(stuff(RIGHT(replicate('0', 6) +  cast(active_start_time AS  VARCHAR(6)), 6), 3, 0, ':'), 6, 0, ':')  else ' starting at '   + stuff(stuff(RIGHT(replicate('0', 6) +  cast(active_start_time AS  VARCHAR(6)), 6), 3, 0, ':'), 6, 0, ':') end time FROM msdb.dbo.sysjobs inner join msdb.dbo.sysjobschedules on sysjobs.job_id = sysjobschedules.job_id inner join msdb.dbo.sysschedules on sysjobschedules.schedule_id = sysschedules.schedule_id where freq_type = 8 
	UNION 
	SELECT  @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', sysjobs.name job_name ,sysjobs.enabled job_enabled ,sysschedules.name schedule_name ,sysschedules.freq_recurrence_factor ,case  WHEN freq_type = 4 THEN 'Daily'  WHEN freq_type = 8 THEN 'Weekly'  WHEN freq_type = 16 THEN 'Monthly'  WHEN freq_type = 32 THEN 'Monthly' end frequency , CASE  WHEN freq_type = 32 THEN (  CASE   WHEN freq_relative_interval = 1 THEN 'First '   WHEN freq_relative_interval = 2 THEN 'Second '   WHEN freq_relative_interval = 4 THEN 'Third '   WHEN freq_relative_interval = 8 THEN 'Fourth '   WHEN freq_relative_interval = 16 THEN 'Last '  end   +  replace  (   CASE WHEN freq_interval = 1 THEN 'Sunday, ' ELSE '' END  +case WHEN freq_interval = 2 THEN 'Monday, ' ELSE '' END  +case WHEN freq_interval = 3 THEN 'Tuesday, ' ELSE '' END  +case WHEN freq_interval = 4 THEN 'Wednesday, ' ELSE '' END  +case WHEN freq_interval = 5 THEN 'Thursday, ' ELSE '' END  +case WHEN freq_interval = 6 THEN 'Friday, ' ELSE '' END  +case WHEN freq_interval = 7 THEN 'Saturday, ' ELSE '' END  +case WHEN freq_interval = 8 THEN 'Day of Month, ' ELSE '' END  +case WHEN freq_interval = 9 THEN 'Weekday, ' ELSE '' END  +case WHEN freq_interval = 10 THEN 'Weekend day, ' ELSE '' END  ,', '  ,''  ) ) else cast(freq_interval AS  VARCHAR(3)) END Days , CASE  WHEN freq_subday_type = 2 THEN ' every ' + cast(freq_subday_interval AS  VARCHAR(7))   + ' seconds' + ' starting at '  + stuff(stuff(RIGHT(replicate('0', 6) +  cast(active_start_time AS  VARCHAR(6)), 6), 3, 0, ':'), 6, 0, ':')   WHEN freq_subday_type = 4 THEN ' every ' + cast(freq_subday_interval AS  VARCHAR(7))   + ' minutes' + ' starting at '  + stuff(stuff(RIGHT(replicate('0', 6) +  cast(active_start_time AS  VARCHAR(6)), 6), 3, 0, ':'), 6, 0, ':')  WHEN freq_subday_type = 8 THEN ' every ' + cast(freq_subday_interval AS  VARCHAR(7))   + ' hours'   + ' starting at '  + stuff(stuff(RIGHT(replicate('0', 6) +  cast(active_start_time AS  VARCHAR(6)), 6), 3, 0, ':'), 6, 0, ':')  else ' starting at '   + stuff(stuff(RIGHT(replicate('0', 6) +  cast(active_start_time AS  VARCHAR(6)), 6), 3, 0, ':'), 6, 0, ':') end time FROM msdb.dbo.sysjobs inner join msdb.dbo.sysjobschedules on sysjobs.job_id = sysjobschedules.job_id inner join msdb.dbo.sysschedules on sysjobschedules.schedule_id = sysschedules.schedule_id where freq_type in (16, 32)
	ORDER BY job_enabled desc
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	------43----
    IF(@Parm4_4=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Jobs that are executing SSIS packages'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, C6, C7, C8, C9, C10, C11) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H', 'SQLInstance',	'job', 'Enabled', 'step', 'SSIS_Package', 'StorageType', 'Server')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, C6, C7, C8, C9, C10, C11) 
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', SQLInstance = @@ServerName, [job]=j.name, j.Enabled, [step]=s.step_name ,  SSIS_Package= CASE WHEN charindex('/ISSERVER', s.command)=1 THEN substring(s.command, len('/ISSERVER "\"')+1, charindex('" /SERVER ', s.command)-len('/ISSERVER "\"')-3) WHEN charindex('/FILE', s.command)=1 THEN substring(s.command, len('/FILE "')+1, charindex('.dtsx', s.command)-len('/FILE "\"')+6) WHEN charindex('/SQL', s.command)=1 THEN substring(s.command, len('/SQL "\"')+1, charindex('" /SERVER ', s.command)-len('/SQL "\"')-3) else s.command end, StorageType = CASE WHEN charindex('/ISSERVER', s.command) = 1 THEN 'SSIS Catalog' WHEN charindex('/FILE', s.command)=1 THEN 'File System' WHEN charindex('/SQL', s.command)=1 THEN 'MSDB'else 'OTHER' end ,  [Server] = CASE WHEN charindex('/ISSERVER', s.command) = 1 THEN replace(replace(substring(s.command, charindex('/SERVER ', s.command)+len('/SERVER ')+1, charindex(' /', s.command, charindex('/SERVER ', s.command)+len('/SERVER '))-charindex('/SERVER ', s.command)-len('/SERVER ')-1), '"\"',''), '\""', '') WHEN charindex('/FILE', s.command)=1 THEN substring(s.command, charindex('"\\', s.command)+3, CHARINDEX('\', s.command, charindex('"\\', s.command)+3)-charindex('"\\', s.command)-3) WHEN charindex('/SQL', s.command)=1 THEN replace(replace(substring(s.command, charindex('/SERVER ', s.command)+len('/SERVER ')+1, charindex(' /', s.command, charindex('/SERVER ', s.command)+len('/SERVER '))-charindex('/SERVER ', s.command)-len('/SERVER ')-1), '"\"',''), '\""', '') else 'OTHER' END  FROM msdb.dbo.sysjobsteps s inner join msdb.dbo.sysjobs j on s.job_id = j.job_id and s.subsystem ='SSIS'
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60) AND R4='H' 
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----44----
    IF(@Parm4_5=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Last seven days failed jobs'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @FinalDate INT; SET @FinalDate = CONVERT(int     , CONVERT( VARCHAR(10), DATEADD(DAY, -7, GETDATE()), 112)     )  
	--huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31,		'H',		'step_id',			'Job Name',		'Step Name',	'run_date',	'run_time',	'sql_severity',	'message',	'server', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9, 	C10,	C11,	C12,	ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',		'L',				'L',			 'L',			'R',		'R',		'C',			'L',		'L',		@ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',  h.step_id,   j.[name],      s.step_name,    h.run_date,      h.run_time,      h.sql_severity,      h.message,  h.server , @ExcelSheetNo  FROM    msdb.dbo.sysjobhistory h      INNER JOIN msdb.dbo.sysjobs j     ON h.job_id = j.job_id      INNER JOIN msdb.dbo.sysjobsteps s     ON j.job_id = s.job_id   AND h.step_id = s.step_id   WHERE    h.run_status = 0  AND h.run_date > @FinalDate    ORDER BY h.instance_id DESC;
	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')	
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN 
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9, c10, c11, c12) -- all fields data must be on sheet for detail error message
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9, c10, c11, c12 FROM ##DataForSheet  WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (31,  231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9, c10, c11, c12) -- all fields
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9, c10, substring(C11,0,100), c12 FROM ##DataForSheet  WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
		--AND (C7 IS NOT NULL OR C8 IS NOT NULL OR C9 IS NOT NULL) AND (C8 IS NOT NULL OR C9 IS NOT NULL)
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
			IF(@HugeDataCounter<=@AllowFewToReport)
				INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo,'{To see error messages details} '
			ELSE
				INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo,'rfr'
			Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	----45---
    IF(@Parm4_6=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='SQL Server Agent alert information '
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'name',	'event_source',	'message_id',	'severity',	'enabled',	'has_notification',	'delay_between_responses',	'occurrence_count',	'last_occurrence_date',	'last_occurrence_time')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14)
	SELECT top 10 @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', name, event_source, message_id, severity, [enabled], has_notification,   delay_between_responses, occurrence_count, last_occurrence_date, last_occurrence_time FROM msdb.dbo.sysalerts WITH (NOLOCK)  ORDER BY last_occurrence_time desc OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
END -- END OF IF(@SQL_Agent_Jobs_information=0) 

IF(@tempdb_check=0)
BEGIN
	SET @QSectionNo=@QSectionNo+1
	SET @QSectionSubNo=1
	SET @QSectionSplitNo=1
	SET @QTotalNo=@QTotalNo+1
    Set @QHeadingOne='tempDb configuration setting'
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, -1, 'H1', CAST(@QSectionNo AS VARCHAR)+'. '+@QHeadingOne) 

	---46--
    IF(@Parm5_1=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Number of data files in tempdb database (Tempdb Data Files)'
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,C5,C6,C7) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H', 'logDate', 'processinfo', 'text')
	DECLARE @filesInTempdb TABLE(logDate varchar(max), processinfo varchar(max), text varchar(max))
	INSERT INTO @filesInTempdb
	EXEC sys.xp_readerrorlog 0, 1, N'The tempdb database has';
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,C5,C6,C7) 
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', * FROM @filesInTempdb
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')

		SELECT @CPUs=COUNT(*) FROM sys.dm_os_schedulers WHERE status = 'VISIBLE ONLINE'		
		SELECT @NoofTempDbfiles=COUNT(mf.file_id) from sys.master_files mf WHERE TYPE=0 AND DB_NAME(mf.database_id)='tempDB'
		if(@CPUs>@NoofTempDbfiles OR @NoofTempDbfiles<8) 
			INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, H2)	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D',
			'According to <a href="https://learn.microsoft.com/en-US/troubleshoot/sql/database-engine/performance/recommendations-reduce-allocation-contention">Microsoft Support</a>, 
			the best approach is to create one tempdb data file per logical processor up to 8 data files
			Here we have '+cast(@CPUs as varchar)+' CPUS so for better performance, add more files','2'			
		else
			INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, H2)	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D',
			'Thumb up, to have '+Cast(@NoofTempDbfiles as varchar)+' files thats match to <a href="https://learn.microsoft.com/en-US/troubleshoot/sql/database-engine/performance/recommendations-reduce-allocation-contention">Microsoft Support</a> recommendations','2'			
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----47---
    IF(@Parm5_2=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Tempdb database files informations'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'FileID',	'Name',	'Path',	'Size (MB)',	'Growth(MB)')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'R',	'R')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',  FileID=[file_id], Name=[name], Path=physical_name, CONVERT(bigint, size/128.0) AS 'Size (MB)', CONVERT(bigint, growth/128.0) AS 'Growth(MB)' FROM sys.master_files WITH (NOLOCK) WHERE DB_NAME(database_id) LIKE '%tempDB%' OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		SELECT @CPUs=COUNT(*) FROM sys.dm_os_schedulers WHERE status = 'VISIBLE ONLINE'		
		SELECT @NoofTempDbfiles=COUNT(mf.file_id) from sys.master_files mf WHERE TYPE=0 AND DB_NAME(mf.database_id)='tempDB'
		if(@CPUs>@NoofTempDbfiles OR @NoofTempDbfiles<8) 
			INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, H2)	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D',
			'According to <a href="https://learn.microsoft.com/en-US/troubleshoot/sql/database-engine/performance/recommendations-reduce-allocation-contention">Microsoft Support</a>, 
			the best approach is to create one tempdb data file per logical processor up to 8 data files
			Here we have '+cast(@CPUs as varchar)+' CPUS so for better performance, add more files','-'			
		else
			INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, H2)	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D',
			'Thumb up, to have '+Cast(@NoofTempDbfiles as varchar)+' files thats match to <a href="https://learn.microsoft.com/en-US/troubleshoot/sql/database-engine/performance/recommendations-reduce-allocation-contention">Microsoft Support</a> recommendations','+'			
					INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
END --END OF IF(@tempDB_Check=0)

IF(@Backup_Check=0)
BEGIN
	SET @QSectionNo=@QSectionNo+1
	SET @QSectionSubNo=1
	SET @QSectionSplitNo=1
	SET @QTotalNo=@QTotalNo+1
    Set @QHeadingOne='Backup policy, latest backup and history information '
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, -1, 'H1', CAST(@QSectionNo AS VARCHAR)+'. '+@QHeadingOne) 
	----48----
    IF(@Parm6_1=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Database Backups for all databases in last one month period'
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4, R5,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12, C13, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H', @QHeadingTw0,	'database',	'type',	'Device',	'Start Date',	'size_mb', 'snapshot?',	 'Latest Backup Location',		'password?',	'recovery', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12, C13, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S', 	'L',	'L',	'L',	'R',	'R', 'R',	 'L',		'R',	'L', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12, C13, RR2, ExclSheet) 
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', bs.database_name,	backuptype = CASE	WHEN bs.type = 'D'			AND bs.is_copy_only = 0 THEN 'Full Database'			WHEN bs.type = 'D'			AND bs.is_copy_only = 1 THEN 'Full Copy-Only Database'			WHEN bs.type = 'I' THEN 'Differential database backup'			WHEN bs.type = 'L' THEN 'Transaction Log' WHEN bs.type = 'F' THEN 'File or filegroup' WHEN bs.type = 'G' THEN 'Differential file' WHEN bs.type = 'P' THEN 'Partial' WHEN bs.type = 'Q' THEN 'Differential partial' END + ' Backup',	
	CASE bf.device_type WHEN 2 THEN 'Disk'			WHEN 5 THEN 'Tape'			WHEN 7 THEN 'Virtual device'			WHEN 9 THEN 'Azure Storage'			WHEN 105 THEN 'A permanent backup device'			ELSE 'Other Device'		END AS DeviceType,	BackupStartDate = CONVERT(VARCHAR, bs.Backup_Start_Date, 20),	backup_size_mb = CONVERT(decimal(10, 2), bs.backup_size/1024./1024.),   CASE bs.is_snapshot WHEN 0 THEN 'NO' ELSE 'YES' END AS [snapshot?],	LatestBackupLocation = bf.physical_device_name,	CASE bms.is_password_protected WHEN 0 THEN 'NO' ELSE 'YES' END AS password_protected,   bs.recovery_model, row_number() over(order by bs.Backup_Start_Date desc), @ExcelSheetNo FROM msdb.dbo.backupset bs LEFT OUTER JOIN msdb.dbo.backupmediafamily bf ON bs.[media_set_id] = bf.[media_set_id] INNER JOIN msdb.dbo.backupmediaset bms ON bs.[media_set_id] = bms.[media_set_id] WHERE (CONVERT(datetime, bs.Backup_Start_Date, 102) >= GETDATE() - 30)   ORDER BY  bs.Backup_Start_Date DESC, bs.database_name ASC;	
	--Observations start
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		---Logic goes here	
		IF (SELECT COUNT(C6) FROM ##DataForSheet WHERE C6 ='Full Database Backup' AND R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo)=0
			INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5, H2)	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D','No full database backup found','2'			
		IF (SELECT COUNT(C6) FROM ##DataForSheet WHERE C6 ='Transaction Log Backup' AND R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo)=0
			INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5, H2)	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D','No transcation log backup found','2'			
		IF (SELECT COUNT(C6) FROM ##DataForSheet WHERE C6 LIKE 'Differential%' AND R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo)=0
			INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5, H2)	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D','No differential file/partial backup found','2'			
		select @No_of_database_having_full_backup=count(distinct C5)  from ##DataForSheet where R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R4='D' AND C6 ='Full Database Backup'--) DistinctDatabases
		select @No_of_database_having_Transaction_log_backup=count(distinct C5)  from ##DataForSheet where R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R4='D' AND C6 like 'Transaction%'--) DistinctDatabases
		select @TotalDatabases=(count(1)-1) from sys.databases
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5, H2)	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D','Full backup found for '+cast(@No_of_database_having_full_backup as varchar)+' databases out of '+ cast(@TotalDatabases as varchar)+' (including master, msdb and model)','0'			
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5, H2)	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D','Transaction log backup found for '+cast(@No_of_database_having_Transaction_log_backup as varchar)+' databases out of '+ cast(@TotalDatabases as varchar)+' (including master, msdb and model)','0'			
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0  
	BEGIN 
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, C9) --- specific fields
		SELECT  R1, R2, S1, R3,  R4,	C5,	C6,	C8,	C9, C11 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, C9) --- specific fields
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C8,	C9, C11 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32 order by RR2
		--ORDER BY C8  desc, C9 Desc 
		IF @HugeDataCounter<=@AllowFewToReport
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5,C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5,C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	----49---
    IF(@Parm6_2=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Last backup information by database'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	--huge data done 	
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Recovery',		'Full Backup',	'Diff; Backup',	'Log Backup', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S', 'L',		'C',			'L',			'L',			'L', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', ISNULL(d.[name], bs.[database_name]) AS [Database], d.recovery_model_desc AS [Recovery Model],  MAX(CASE WHEN [type] = 'D' THEN CONVERT(VARCHAR, bs.backup_finish_date, 20) ELSE NULL END) AS [Last Full Backup],    MAX(CASE WHEN [type] = 'I' THEN CONVERT(VARCHAR, bs.backup_finish_date, 20) ELSE NULL END) AS [Last Differential Backup],    MAX(CASE WHEN [type] = 'L' THEN bs.backup_finish_date ELSE NULL END) AS [Last Log Backup], @ExcelSheetNo FROM sys.databases AS d WITH (NOLOCK) LEFT OUTER JOIN msdb.dbo.backupset AS bs WITH (NOLOCK) ON bs.[database_name] = d.[name]  AND bs.backup_finish_date > GETDATE()- 30 WHERE d.name <> N'tempdb' GROUP BY ISNULL(d.[name], bs.[database_name]), d.recovery_model_desc, d.log_reuse_wait_desc, d.[name]  ORDER BY d.recovery_model_desc, d.[name] OPTION (RECOMPILE);

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN 
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9) -- all fields
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9 FROM ##DataForSheet  WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (31,  231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9) -- all fields
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9 FROM ##DataForSheet  WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
		--AND (C7 IS NOT NULL OR C8 IS NOT NULL OR C9 IS NOT NULL) AND (C8 IS NOT NULL OR C9 IS NOT NULL)
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		IF @HugeDataCounter>@AllowFewToReport
		BEGIN
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
			Set @ExcelSheetNo=@ExcelSheetNo+1;
		END
		ELSE
			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 

	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	----50---
    IF(@Parm6_3=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Look at recent full backups for each database'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset26 TABLE(C5  VARCHAR(MAX),	C6  VARCHAR(MAX),	C7  VARCHAR(MAX),	C8  VARCHAR(MAX),	C9  VARCHAR(MAX),	C10  VARCHAR(MAX),	C11  VARCHAR(MAX),	C12  VARCHAR(MAX),	C13  VARCHAR(MAX),	C14  VARCHAR(MAX),	C15  VARCHAR(MAX),	C16  VARCHAR(MAX),	C17  VARCHAR(MAX),	C18  VARCHAR(MAX))
	INSERT INTO @Dataset26
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) 
	SELECT TOP (20) bs.machine_name, bs.server_name, bs.database_name AS [Database], bs.recovery_model, CONVERT (BIGINT, bs.backup_size / 1048576 ) AS [Uncompressed Backup Size (MB)], CONVERT (BIGINT, bs.compressed_backup_size / 1048576 ) AS [Compressed Backup Size (MB)], CONVERT (NUMERIC (20,2), (CONVERT (FLOAT, bs.backup_size) / CONVERT (FLOAT, bs.compressed_backup_size))) AS [Compression Ratio], bs.has_backup_checksums, bs.is_copy_only, bs.encryptor_type, DATEDIFF (SECOND, bs.backup_start_date, bs.backup_finish_date) AS [Backup Elapsed Time (sec)], CONVERT(VARCHAR, bs.backup_finish_date, 20) AS [Backup Finish Date], bmf.physical_device_name AS [Backup Location], bmf.physical_block_size FROM msdb.dbo.backupset AS bs WITH (NOLOCK) INNER JOIN msdb.dbo.backupmediafamily AS bmf WITH (NOLOCK) ON bs.media_set_id = bmf.media_set_id   WHERE bs.database_name = DB_NAME(DB_ID()) AND bs.[type] = ''D'' -- Change to L if you want Log backups  ORDER BY bs.backup_finish_date DESC OPTION (RECOMPILE);'

	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'machine_name',	'server',	'Name',	'recovery_model',	'Uncompr Backup Size (MB)',	'Compressed Size (MB)',	'Comp Ratio',	'has_checksums?',	'copy_only?',	'encrypt',	'Elapsed Time (sec)',	'Finish Date',	'Backup Location',	'physical block size', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'L',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'L',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset26
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8) --specific
		SELECT R1, R2, S1, R3,  R4, C7, C8, C16, C17 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4, C7, C8, C16, C17 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		--DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		IF @HugeDataCounter<=@AllowFewToReport
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5,C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5,C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'

		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
END --END OF IF(@Backup_Check=0)

IF(@Database_information=0)
BEGIN
	SET @QSectionNo=@QSectionNo+1
	SET @QSectionSubNo=1
	SET @QSectionSplitNo=1
	SET @QTotalNo=@QTotalNo+1
    Set @QHeadingOne='Database Details Information Query Suite:'
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, -1, 'H1', CAST(@QSectionNo AS VARCHAR)+'. '+@QHeadingOne) 
	DECLARE @dbNamesWithPageVerify varchar(max)=''

	set @dbNamesWithPageVerify +='<font size="2">
		<span><b><br>Query Execution Timestamp Display: </b></span>'+	CAST(GetDATE() AS VARCHAR)

	set @dbNamesWithPageVerify +='<table><tr style="border-top-style: dashed;border-bottom-width: medium;"><td style="vertical-align:top">'
	set @dbNamesWithPageVerify +='<font size="4"><b>Database Settings summary</b></font><font size="3"><br>'
	set @dbNamesWithPageVerify +='<font size="3"><table class="Sumry_Report" style="width: 360px;">'
	set @dbNamesWithPageVerify += '<tr><td>Snapshot Isolation ON (database#)<td>'+cast((SELECT count(1) FROM sys.databases WHERE snapshot_isolation_state_desc = 'ON') as varchar)
	set @dbNamesWithPageVerify += '<tr><td>Read Committed Snapshot Isolation (database#)<td>'+ cast((SELECT count(1) FROM sys.databases WHERE is_read_committed_snapshot_on= 1) as varchar)
	if(SELECT count(1) FROM sys.databases WHERE is_read_only= 1)>0
		set @dbNamesWithPageVerify += '<tr><td>Read only (database#)<td>'+ cast((SELECT count(1) FROM sys.databases WHERE is_read_only= 1) as varchar)
	if(SELECT count(1) FROM sys.databases WHERE is_query_store_on= 1)>0
		set @dbNamesWithPageVerify += '<tr><td>Query store On (database#)<td>'+ cast((SELECT count(1) FROM sys.databases WHERE is_query_store_on= 1) as varchar)
	if(SELECT count(1) FROM sys.databases WHERE is_cdc_enabled= 1)>0
		set @dbNamesWithPageVerify += '<tr><td>CDC Enabled (database#)<td>'+ cast((SELECT count(1) FROM sys.databases WHERE is_cdc_enabled= 1) as varchar)
	if(SELECT count(1) FROM sys.databases WHERE is_distributor= 1)>0
		set @dbNamesWithPageVerify += '<tr><td>Distributor Databases<td>'+ cast((SELECT count(1) FROM sys.databases WHERE is_distributor= 1) as varchar)
	if(SELECT count(1) FROM sys.databases WHERE is_published= 1)>0
		set @dbNamesWithPageVerify += '<tr><td>Publishing Databases<td>'+ cast((SELECT count(1) FROM sys.databases WHERE is_published= 1) as varchar)
	if(SELECT count(1) FROM sys.databases WHERE is_encrypted= 1)>0
		set @dbNamesWithPageVerify += '<tr><td>Encrypted Databases<td>'+ cast((SELECT count(1) FROM sys.databases WHERE is_encrypted= 1) as varchar)
	if(SELECT count(1) FROM sys.databases WHERE is_subscribed= 1)>0
		set @dbNamesWithPageVerify += '<tr><td>Subscribed Databases<td>'+ cast((SELECT count(1) FROM sys.databases WHERE is_subscribed= 1) as varchar)
	if(SELECT count(1) FROM sys.databases WHERE is_memory_optimized_enabled= 1)>0
		set @dbNamesWithPageVerify += '<tr><td>Memory optimized Databases<td>'+ cast((SELECT count(1) FROM sys.databases WHERE is_memory_optimized_enabled= 1) as varchar)
	SET @dbNamesWithPageVerify +='</table></font>'  
	set @dbNamesWithPageVerify +='<td style="vertical-align:top">'

	set @dbNamesWithPageVerify +='<font size="4"><b>Page Verification Options</b></font><font size="3"><br>'
	set @dbNamesWithPageVerify +='<table class="Sumry_Report" style="width: 360px;">'
	set @dbNamesWithPageVerify +='<tr><td>Total Database(s) on server<td>'+ cast((SELECT count(1) FROM sys.databases) as varchar)
	set @dbNamesWithPageVerify += '<tr><td>Checksum (database#)<td>'+cast((SELECT count(1) FROM sys.databases WHERE page_verify_option_desc = 'CHECKSUM') as varchar)
	set @dbNamesWithPageVerify += '<tr><td>Torn_page_detection (database#)<td>'+ cast((SELECT count(1) FROM sys.databases WHERE page_verify_option_desc = 'TORN_PAGE_DETECTION') as varchar)
	if(SELECT count(1) FROM sys.databases WHERE page_verify_option_desc = 'None')>0
		set @dbNamesWithPageVerify += '<tr style="color: red;"><td>None (database#)<td>'+cast((SELECT count(1) FROM sys.databases WHERE page_verify_option_desc = 'None') as varchar)
	SET @dbNamesWithPageVerify +='</table></font>'  

	declare @bigdbName varchar(50),@smalldbName varchar(50), @largesize decimal(18,2), @smallsize decimal(18,2)
	SELECT TOP 1 @bigdbName=DB_NAME(database_id) , @largesize=SUM(size * 8.0 / 1024/1024 ) FROM sys.master_files WHERE type_desc = 'ROWS' and database_id > 4 GROUP BY  database_id ORDER BY  SUM(size) desc
	SELECT TOP 1 @smalldbName=DB_NAME(database_id) , @smallsize=SUM(size * 8.0 / 1024 ) FROM sys.master_files WHERE type_desc = 'ROWS' and database_id > 4 GROUP BY  database_id ORDER BY  SUM(size) asc

	set @dbNamesWithPageVerify +='<font size="4"><b><br>Database Sizes</b></font><font size="3"><br>'
	set @dbNamesWithPageVerify +='<table class="Sumry_Report" style="width: 360px;">'
	set @dbNamesWithPageVerify +='<tr><td>Sum of all database size <td><b><u>'+ cast((SELECT convert(decimal(18,2),SUM(size * 8.0 / 1024 / 1024)) AS [Total Size (GB)] FROM sys.master_files WHERE type_desc = 'ROWS') as varchar) +' GB</b></u>'
	set @dbNamesWithPageVerify += '<tr><td>Biggest Databases (<b>'+@bigdbName+'</b>) <td> '+cast(@largesize as varchar)+ ' GB'
	set @dbNamesWithPageVerify += '<tr><td>Smallest Database (<b>'+@smalldbName+'</b>)<td>'+cast(@smallsize as varchar)+ ' MB'
	SET @dbNamesWithPageVerify +='</table></font>'  

	set @dbNamesWithPageVerify +='</table>'
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) 
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0, 'HTML',@dbNamesWithPageVerify

	----51----
	IF(@Parm7_1=0)
	BEGIN
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1

    BEGIN TRY
	SET @QHeadingTw0='Recovery model, log reuse wait description, log file size, log usage size (Database Properties)'
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')


	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20,	C21,	C22,	C23,	C24,	C25,	C26,	C27,	C28,	C29,	C30,	C31,	C32,	C33,	C34,	C35,	C36,	C37,	C38, ExclSheet)	VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Database Owner',	'Recovery Model',	'state_desc',	'containment_desc',	'Log Reuse Wait Description',	'Log Size (MB)',	'Log Used (MB)',	'Log Used %',	'DB Compatibility Level',	'Page Verify Option',	'is_auto_create_stats_on',	'is_auto_update_stats_on',	'is_auto_update_stats_async_on',	'is_parameterization_forced',	'snapshot_isolation_state_desc',	'is_read_committed_snapshot_on',	'is_auto_close_on',	'is_auto_shrink_on',	'target_recovery_time_in_seconds',	'is_cdc_enabled',	'is_published',	'is_distributor',	'is_encrypted',	'group_database_id',	'replica_id',	'is_memory_optimized_elevate_to_snapshot_on',	'delayed_durability_desc',	'is_auto_create_stats_incremental_on',	'is_encrypted',	'encryption_state',	'percent_complete',	'key_algorithm',	'key_length', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20,	C21,	C22,	C23,	C24,	C25,	C26,	C27,	C28,	C29,	C30,	C31,	C32,	C33,	C34,	C35,	C36,	C37,	C38, ExclSheet)	VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231,'S',	'L',		'L',				'C',				'L',			'L',				'L',							'R',				'R',				'R',			'C',						'L',					'R',						'R',						'R',								'R',							'L',								'R',								'R',				'R',					'R',								'R',				'R',			'R',				'R',	'L',	'L',	'R',	'L',	'R',	'R',	'L',	'R',	'L',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20,	C21,	C22,	C23,	C24,	C25,	C26,	C27,	C28,	C29,	C30,	C31,	C32,	C33,	C34,	C35,	C36,	C37,	C38, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', db.[name] AS [Database], SUSER_SNAME(db.owner_sid) AS [Database Owner], db.recovery_model_desc AS [Recovery Model],  db.state_desc, db.containment_desc, db.log_reuse_wait_desc AS [Log Reuse Wait Description],  CONVERT(DECIMAL(18,2), ls.cntr_value/1024.0) AS [Log Size (MB)], CONVERT(DECIMAL(18,2), lu.cntr_value/1024.0) AS [Log Used (MB)], CAST(CAST(lu.cntr_value AS FLOAT) / CAST(ls.cntr_value AS FLOAT)AS DECIMAL(18,2)) * 100 AS [Log Used %],  db.[compatibility_level] AS [DB Compatibility Level], db.page_verify_option_desc AS [Page Verify Option],  db.is_auto_create_stats_on, db.is_auto_update_stats_on, db.is_auto_update_stats_async_on, db.is_parameterization_forced,  db.snapshot_isolation_state_desc, db.is_read_committed_snapshot_on, db.is_auto_close_on, db.is_auto_shrink_on,  db.target_recovery_time_in_seconds, db.is_cdc_enabled, db.is_published, db.is_distributor, db.is_encrypted, db.group_database_id, db.replica_id,db.is_memory_optimized_elevate_to_snapshot_on,  db.delayed_durability_desc, db.is_auto_create_stats_incremental_on, db.is_encrypted, de.encryption_state, de.percent_complete, de.key_algorithm, de.key_length, @ExcelSheetNo  FROM sys.databases AS db WITH (NOLOCK) INNER JOIN sys.dm_os_performance_counters AS lu WITH (NOLOCK) ON db.name = lu.instance_name INNER JOIN sys.dm_os_performance_counters AS ls WITH (NOLOCK) ON db.name = ls.instance_name LEFT OUTER JOIN sys.dm_database_encryption_keys AS de WITH (NOLOCK) ON db.database_id = de.database_id WHERE db_name(de.database_id) not  in ('master','msdb','tempdb','distribution','DWQueue' ,'DWDiagnostics','model')  and lu.counter_name LIKE N'Log File(s) Used Size (KB)%'  AND ls.counter_name LIKE N'Log File(s) Size (KB)%' AND ls.cntr_value > 0   ORDER BY db.[name] OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, C6, C7, C8, C9, C10, C11, C12) --specific
		SELECT  R1, R2, S1, R3, R4, C5, C6, C7, C8, C11, C13, C14, C30 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, C6, C7, C8, C9, C10, C11, C12) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3, R4, C5, C6, C7, C8, C11, C13, C14, C30 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
		IF @HugeDataCounter<=@AllowFewToReport 
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		SET @ExcelSheetNo=@ExcelSheetNo+1
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm

	----52---
    IF(@Parm7_2=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Last successful DBCC CHECKDB that ran on the specified database'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'name',	'recovery',	'status',	'DBCheckLast')
	
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'R')
	IF OBJECT_ID('tempdb..#tempDbCheckList') IS NOT NULL
	DROP TABLE #tempDbCheckList
	SELECT name, 
	cast(DATABASEPROPERTYEX(name, 'Recovery') AS  VARCHAR) AS [recovery], 
	cast(DATABASEPROPERTYEX(name, 'Status') AS  VARCHAR) AS [status], 
	TRY_CAST(DATABASEPROPERTYEX(name, 'LastGoodCheckDbTime') AS  datetime) AS [DBCheckLast],
	DATEDIFF(m, TRY_CAST(DATABASEPROPERTYEX(name, 'LastGoodCheckDbTime') AS  datetime), Getdate()) as NoOfMonths
	INTO #tempDbCheckList
	FROM master.dbo.sysdatabases where name not in('master','msdb', 'model', 'tempDB')  
	ORDER BY DBCheckLast 
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',
	name,recovery, status, CASE WHEN NoOfMonths <= 1 THEN '<span class=safe>'+cast(DBCheckLast as varchar)+'</span>' ELSE '<span class=critical>'+cast(DBCheckLast as varchar) +'</span>' END as DBCheckLast FROM #tempDbCheckList

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		---Logice goes here	
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm

	----53---
    IF(@Parm7_3=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Input buffer information for the following databases'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset25 TABLE(C5  VARCHAR(MAX),	C6  VARCHAR(MAX),	C7  VARCHAR(MAX),	C8  VARCHAR(MAX),	C9  VARCHAR(MAX),	C10  VARCHAR(MAX),	C11  VARCHAR(MAX))
	INSERT INTO @Dataset25
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) SELECT es.session_id, DB_NAME(es.database_id) AS [Database], es.login_time, es.cpu_time, es.logical_reads, es.[status], ib.event_info AS [Input Buffer] FROM sys.dm_exec_sessions AS es WITH (NOLOCK) CROSS APPLY sys.dm_exec_input_buffer(es.session_id, NULL) AS ib WHERE es.database_id = DB_ID() AND es.session_id > 50 AND es.session_id <> @@SPID OPTION (RECOMPILE);'
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'session_id',	'Database',	'login_time',	'cpu_time',	'logical_reads',	'status',	'Input Buffer')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'R',	'R',	'R',	'L',	'R')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', * FROM @Dataset25
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		---Logice goes here	
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm

	----54----
    IF(@Parm7_4=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Key table properties (Table Properties)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset17 TABLE(C5  VARCHAR(MAX),	C6  VARCHAR(MAX),	C7  VARCHAR(MAX),	C8  VARCHAR(MAX),	C9  VARCHAR(MAX),	C10  VARCHAR(MAX),	C11  VARCHAR(MAX),	C12  VARCHAR(MAX),	C13  VARCHAR(MAX),	C14  VARCHAR(MAX),	C15  VARCHAR(MAX),	C16  VARCHAR(MAX),	C17  VARCHAR(MAX))
	INSERT INTO @Dataset17
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) SELECT OBJECT_NAME(t.[object_id]) AS [ObjectName], p.[rows] AS [Table Rows], p.index_id,    p.data_compression_desc AS [Index Data Compression],   CONVERT(VARCHAR, create_date, 20), t.lock_on_bulk_load, t.is_replicated, t.has_replication_filter,    t.is_tracked_by_cdc, t.lock_escalation_desc, t.is_filetable,  	   t.is_memory_optimized, t.durability_desc   FROM sys.tables AS t WITH (NOLOCK) INNER JOIN sys.partitions AS p WITH (NOLOCK) ON t.[object_id] = p.[object_id] WHERE OBJECT_NAME(t.[object_id]) NOT LIKE N''sys%''  ORDER BY OBJECT_NAME(t.[object_id]), p.index_id OPTION (RECOMPILE);'
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'ObjectName',	'Table Rows',	'index_id',	'Index Data Compression',	'create_date',	'lock_on_bulk_load',	'is_replicated',	'has_replication_filter',	'is_tracked_by_cdc',	'lock_escalation_desc',	'is_filetable',	'is_memory_optimized',	'durability_desc', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo,231,	'S', 'L',			'R',			'L',		'C',						'R',			'L',					'R',				'R',						'R',					'R',					'R',			'R',					'L', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset17
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) 
		SELECT  R1, R2, S1, R3, R4,	C5,	C6,	C8,	C9 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3, R4,	C5,	C6,	C8,	C9 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
		IF @HugeDataCounter<=@AllowFewToReport 
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		SET @ExcelSheetNo=@ExcelSheetNo+1
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----55----
    IF(@Parm7_5=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Tables, row counts, and compression status for clustered index or heap (Table Sizes, Rows, rowcount row count, total rows)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset16 TABLE([Database]  VARCHAR(255),[Schema]  VARCHAR(255),[ObjectName]  VARCHAR(255), [RowCount] BIGINT,[Compression Type]  VARCHAR(255))
	INSERT INTO @Dataset16 
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) SELECT DB_NAME(DB_ID()) AS [Database], SCHEMA_NAME(o.Schema_ID) AS [Schema], OBJECT_NAME(p.object_id) AS [ObjectName], SUM(p.Rows) AS [RowCount], p.data_compression_desc AS [Compression Type]FROM sys.partitions AS p WITH (NOLOCK)INNER JOIN sys.objects AS o WITH (NOLOCK)ON p.object_id = o.object_id WHERE index_id < 2  AND OBJECT_NAME(p.object_id) NOT LIKE N''sys%'' AND OBJECT_NAME(p.object_id) NOT LIKE N''spt_%'' AND OBJECT_NAME(p.object_id) NOT LIKE N''queue_%''  AND OBJECT_NAME(p.object_id) NOT LIKE N''filestream_tombstone%''  AND OBJECT_NAME(p.object_id) NOT LIKE N''fulltext%'' AND OBJECT_NAME(p.object_id) NOT LIKE N''ifts_comp_fragment%'' AND OBJECT_NAME(p.object_id) NOT LIKE N''filetable_updates%'' AND OBJECT_NAME(p.object_id) NOT LIKE N''xml_index_nodes%'' AND OBJECT_NAME(p.object_id) NOT LIKE N''sqlagent_job%'' AND OBJECT_NAME(p.object_id) NOT LIKE N''plan_persist%'' GROUP BY  SCHEMA_NAME(o.Schema_ID), p.object_id, data_compression_desc  ORDER BY SUM(p.Rows) DESC OPTION (RECOMPILE);'
	--huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Schema',	'ObjectName',	'RowCount',	'Compression Type', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'R',	'C', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset16
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logice goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32

	 IF @HugeDataCounter>0 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) -- all
		SELECT  R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
		IF @HugeDataCounter<=@AllowFewToReport 
			DELETE FROM  ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----56---
    IF(@Parm7_6=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Log space usage of each database'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @command  VARCHAR(max)
	SET @command='USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model''))  
	SELECT DB_NAME(lsu.database_id) AS [Database], db.recovery_model_desc AS [Recovery Model], 		CAST(lsu.total_log_size_in_bytes/1048576.0 AS DECIMAL(10, 2)) AS [Total Log Space (MB)], 		CAST(lsu.used_log_space_in_bytes/1048576.0 AS DECIMAL(10, 2)) AS [Used Log Space (MB)],  		CAST(lsu.used_log_space_in_percent AS DECIMAL(10, 2)) AS [Used Log Space %], 		CAST(lsu.log_space_in_bytes_since_last_backup/1048576.0 AS DECIMAL(10, 2)) AS [Used Log Space Since Last Backup (MB)], 		db.log_reuse_wait_desc		  FROM sys.dm_db_log_space_usage AS lsu WITH (NOLOCK) INNER JOIN sys.databases AS db WITH (NOLOCK) ON lsu.database_id = db.database_id OPTION (RECOMPILE)';
	DECLARE  @Dataset2 TABLE(  [Database]  VARCHAR(255),  [Recovery Model]  NVARCHAR(255),  [Total Log Space (MB)]  DECIMAL(16,2),  [Used Log Space (MB)]  DECIMAL(16,2),  [Used Log Space %]  DECIMAL(16,2),  [Used Log Space Since Last Backup (MB)] DECIMAL(16,2),  [log_reuse_wait_desc]  VARCHAR(255))
	INSERT  INTO @Dataset2
	EXEC sp_MSforeachdb @command
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Recovery Model',	'Total Log Space (MB)',	'Used Log Space (MB)',	'Used Log Space %',	'Used Log Space Since Last Backup (MB)',	'log_reuse_wait_desc', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',				'R',					'R',					'R',				'R',										'C', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',*, @ExcelSheetNo FROM @Dataset2

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11  FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)		
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11  FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32 		
		IF @HugeDataCounter<=@AllowFewToReport 
			DELETE FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;

	--highlight Red Used log space % < 20 or > 70
	UPDATE DFS SET DFS.C9='<span class=critical>'+DFS.C9+'</span>' 
	FROM ##RunOnce DFS 
	WHERE DFS.R1=@QSectionNo AND DFS.R2=@QSectionSubNo AND DFS.S1=@QSectionSplitNo AND DFS.R3=32
	AND TRY_CAST(DFS.C9 AS DECIMAL(16,2))>70.0 
	UPDATE DFS SET DFS.C9='<span class=safe>'+DFS.C9+'</span>' 
	FROM ##RunOnce DFS 
	WHERE DFS.R1=@QSectionNo AND DFS.R2=@QSectionSubNo AND DFS.S1=@QSectionSplitNo AND DFS.R3=32
	AND TRY_CAST(DFS.C9 AS DECIMAL(16,2))<20.0 

	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----57----
    IF(@Parm7_7=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Individual File Sizes and space available for current database'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset1 TABLE(C5  VARCHAR(max),	C6  VARCHAR(max),	C7  VARCHAR(max),	C8  VARCHAR(max),	C9  VARCHAR(max),	C10  VARCHAR(max),	C11  VARCHAR(max),	C12  VARCHAR(max),	C13  VARCHAR(max),	C14  VARCHAR(max),	C15  VARCHAR(max))
	INSERT INTO @Dataset1
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) SELECT DB_name(db_id()) AS dbname, f.name AS [File Name] , f.physical_name AS [Physical Name],  CAST((f.size/128.0) AS DECIMAL(15,2)) AS [Total Size in MB], CAST(f.size/128.0 - CAST(FILEPROPERTY(f.name, ''SpaceUsed'') AS int)/128.0 AS DECIMAL(15,2))  AS [Available Space In MB], f.[file_id], fg.name AS [Filegroup Name], f.is_percent_growth, f.growth, fg.is_default, fg.is_read_only FROM sys.database_files AS f WITH (NOLOCK)  LEFT OUTER JOIN sys.filegroups AS fg WITH (NOLOCK) ON f.data_space_id = fg.data_space_id  ORDER BY f.[file_id] OPTION (RECOMPILE)';
	--huge data expected done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'dbname',	'File Name',	'Physical Name',	'Total Size in MB',	'Available Space In MB',	'file_id',	'Filegroup Name',	'is_percent_growth',	'growth',	'is_default',	'is_read_only', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'R',	'R',	'L',	'L',	'R',	'R',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset1 
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9) --specific
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32)
		
		IF @HugeDataCounter<=@AllowFewToReport 
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----58---
    IF(@Parm7_8=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='File names and paths for all user and system databases on instance'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'file_id',	'name',	'physical_name',	'type_desc',	'state_desc',	'is_percent_growth',	'growth',	'Growth in MB',	'Total Size in MB',	'max_size', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'L',	'L',	'R',	'R',	'R',	'R',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', DB_NAME([database_id]) AS [Database],    [file_id], [name], physical_name, [type_desc], state_desc, 	   is_percent_growth, growth,  	   CONVERT(bigint, growth/128.0) AS [Growth in MB],    CONVERT(bigint, size/128.0) AS [Total Size in MB], max_size, @ExcelSheetNo FROM sys.master_files WITH (NOLOCK)  ORDER BY [Total Size in MB] OPTION (RECOMPILE);
	--huge data done
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0  
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11) --specific
		SELECT R1, R2, S1, R3,  R4,	C5,	C7,	C8,	C9, C13,	C14,	C15 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C7,	C8,	C9, C13,	C14,	C15 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32)
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		IF @HugeDataCounter<=@AllowFewToReport
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----59---
    IF(@Parm7_9=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Virtual log files (VLF) Counts for all databases on the instance'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	CREATE TABLE  #VLFInfo (RecoveryUnitID INT, FileID  INT,   FileSize BIGINT, StartOffset BIGINT,   FSeqNo BIGINT, [Status]    BIGINT, 					   Parity BIGINT, CreateLSN   NUMERIC(38)); 	  
	CREATE TABLE  #VLFCountResults(DatabaseName sysname, VLFCount int);  EXEC sp_MSforeachdb N'Use [?];  				INSERT INTO #VLFInfo  				EXEC sp_executesql N''DBCC LOGINFO([?])'';  	  				INSERT INTO #VLFCountResults  				SELECT DB_NAME(), COUNT(*)  				FROM #VLFInfo;  				TRUNCATE TABLE #VLFInfo;' 	  
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'DatabaseName',	'VLFCount')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'H',	'L',	'R')
	INSERT INTO ##RunOnce(R1, R2, S1, QuerySort, R3, R4,	C5,	C6)
	SELECT TOP 10 @QSectionNo, @QSectionSubNo, @QSectionSplitNo, ROW_NUMBER() OVER(ORDER BY VLFCount DESC), 32, 'D', DatabaseName, VLFCount   FROM #VLFCountResults  ORDER BY VLFCount DESC; 
	DROP TABLE #VLFInfo; DROP TABLE #VLFCountResults;
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		---Logice goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
END -- END OF IF(@Database_information=0)

IF(@Database_Performance=0)
BEGIN
	SET @QSectionNo=@QSectionNo+1
	SET @QSectionSubNo=1
	SET @QSectionSplitNo=1
	SET @QTotalNo=@QTotalNo+1
    Set @QHeadingOne='Database performance information'
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, -1, 'H1', CAST(@QSectionNo AS VARCHAR)+'. '+@QHeadingOne) 
	----60---
    IF(@Parm8_10=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Breaks down buffers used by each database by object (table, index) in the buffer cache (Buffer Usage)'
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset15 TABLE([Database]  VARCHAR(255),[Schema]  VARCHAR(255),[Object]  VARCHAR(255),index_id INT,[Buffer size(MB)] DECIMAL(8,2),BufferCount INT, [Row Count] BIGINT, [Compression Type]  VARCHAR(120))
	INSERT INTO @Dataset15 
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) SELECT DB_NAME(DB_ID()) AS [Database], SCHEMA_NAME(o.Schema_ID) AS [Schema], OBJECT_NAME(p.[object_id]) AS [Object], p.index_id,  CAST(COUNT(*)/128.0 AS DECIMAL(10, 2)) AS [Buffer size(MB)],   COUNT(*) AS [BufferCount], p.[Rows] AS [Row Count], p.data_compression_desc AS [Compression Type] FROM sys.allocation_units AS a WITH (NOLOCK) INNER JOIN sys.dm_os_buffer_descriptors AS b WITH (NOLOCK) ON a.allocation_unit_id = b.allocation_unit_id INNER JOIN sys.partitions AS p WITH (NOLOCK) ON a.container_id = p.hobt_id INNER JOIN sys.objects AS o WITH (NOLOCK) ON p.object_id = o.object_id WHERE b.database_id = CONVERT(int, DB_ID()) AND p.[object_id] > 100 AND OBJECT_NAME(p.[object_id]) NOT LIKE N''plan_%'' AND OBJECT_NAME(p.[object_id]) NOT LIKE N''sys%'' AND OBJECT_NAME(p.[object_id]) NOT LIKE N''xml_index_nodes%'' GROUP BY o.Schema_ID, p.[object_id], p.index_id, p.data_compression_desc, p.[Rows]  ORDER BY [BufferCount] DESC OPTION (RECOMPILE);'
	--huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Schema',	'Object',	'index_id',	'Buffer size(MB)',	'BufferCount',	'Row Count',	'Compression Type', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'R',	'R',	'R',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset15 order by [Row Count] desc
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0  
	BEGIN
		--Separate Header from Data for the sake of Order By clause
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9, C10) --specific
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C9, C10, C11 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (31, 231) 
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9, C10) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C9, C10, C11 FROM ##DataForSheet 
		WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
		ORDER BY TRY_CAST(C11 AS INT) DESC
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		IF(@HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo,'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo,'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;

	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----61---
    IF(@Parm8_11=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='I/O Statistics by file for each database'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset3 TABLE([Database]  VARCHAR(255),[Logical Name]  VARCHAR(255),file_id INT,[type_desc]  VARCHAR(50),[Physical Name]  VARCHAR(400),[Size on Disk (MB)] DECIMAL(16,2),[num_of_reads] DECIMAL(16,2),[num_of_writes] DECIMAL(16,2),[io_stall_read_ms] DECIMAL(16,2),[io_stall_write_ms] DECIMAL(16,2),[IO Stall Reads Pct] DECIMAL(16,2),[IO Stall Writes Pct] DECIMAL(16,2),[Writes + Reads] DECIMAL(16,2),[MB Read] DECIMAL(16,2),[MB Written]DECIMAL(16,2),[# Reads Pct] DECIMAL(16,2),[# Write Pct] DECIMAL(16,2),[Read Bytes Pct] DECIMAL(16,2),[Written Bytes Pct] DECIMAL(16,2))
	INSERT INTO @Dataset3 EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model''))  SELECT DB_NAME(DB_ID()) AS [Database], df.name AS [Logical Name], vfs.[file_id], df.type_desc, df.physical_name AS [Physical Name], CAST(vfs.size_on_disk_bytes/1048576.0 AS DECIMAL(16, 2)) AS [Size on Disk (MB)], vfs.num_of_reads, vfs.num_of_writes, vfs.io_stall_read_ms, vfs.io_stall_write_ms, CAST(100. * vfs.io_stall_read_ms/(vfs.io_stall_read_ms + vfs.io_stall_write_ms) AS DECIMAL(16,1)) AS [IO Stall Reads Pct], CAST(100. * vfs.io_stall_write_ms/(vfs.io_stall_write_ms + vfs.io_stall_read_ms) AS DECIMAL(16,1)) AS [IO Stall Writes Pct], (vfs.num_of_reads + vfs.num_of_writes) AS [Writes + Reads],  CAST(vfs.num_of_bytes_read/1048576.0 AS DECIMAL(16, 2)) AS [MB Read],  CAST(vfs.num_of_bytes_written/1048576.0 AS DECIMAL(16, 2)) AS [MB Written], CAST(100. * vfs.num_of_reads/(vfs.num_of_reads + vfs.num_of_writes) AS DECIMAL(16,1)) AS [# Reads Pct], CAST(100. * vfs.num_of_writes/(vfs.num_of_reads + vfs.num_of_writes) AS DECIMAL(16,1)) AS [# Write Pct], CAST(100. * vfs.num_of_bytes_read/(vfs.num_of_bytes_read + vfs.num_of_bytes_written) AS DECIMAL(16,1)) AS [Read Bytes Pct], CAST(100. * vfs.num_of_bytes_written/(vfs.num_of_bytes_read + vfs.num_of_bytes_written) AS DECIMAL(16,1)) AS [Written Bytes Pct] FROM sys.dm_io_virtual_file_stats(DB_ID(), NULL) AS vfs INNER JOIN sys.database_files AS df WITH (NOLOCK) ON vfs.[file_id]= df.[file_id] OPTION (RECOMPILE)';
	--split data
	--huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20,	C21,	C22,	C23, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Logical Name',	'file_id',	'type_desc',	'Physical Name',	'Size on Disk (MB)',	'num_of_reads',	'num_of_writes',	'io_stall_read_ms',	'io_stall_write_ms',	'IO Stall Reads Pct',	'IO Stall Writes Pct',	'Writes + Reads',	'MB Read',	'MB Written',	'# Reads Pct',	'# Write Pct',	'Read Bytes Pct',	'Written Bytes Pct', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20,	C21,	C22,	C23, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231,	'S',	'L',	'L',			'L',		'L',			'L',				'R',					'R',			'R',				'R',				'R',					'R',					'R',					'R',				'R',		'R',			'R',			'R',			'R',				'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20,	C21,	C22,	C23, RR2, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, [Writes + Reads], @ExcelSheetNo FROM @Dataset3 
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0  
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12) --specifics
		SELECT  R1, R2, S1, R3,  R4,	C5,	C6,	C10,	C11,	C12, C13,	C14,	C17 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C10,	C11,	C12, C13,	C14,	C17 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32 --, 32, 231)
		ORDER BY RR2 DESC


		IF @HugeDataCounter<=@AllowFewToReport
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'

		Set @ExcelSheetNo=@ExcelSheetNo+1;


	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	-----62--
    IF(@Parm8_12=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Calculates average stalls per read, per write, and per total input/output for each database file (IO Latency by File)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	--huge data done
	--split data
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'avg_read_latency_ms',	'avg_write_latency_ms',	'avg_io_latency_ms',	'File Size (MB)',	'physical_name',	'type_desc',	'io_stall_read_ms',	'num_of_reads',	'io_stall_write_ms',	'num_of_writes',	'io_stalls',	'total_io',	'Resource Governor Total Read IO Latency (ms)',	'Resource Governor Total Write IO Latency (ms)', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',				'L',		'R',					'R',					'R',					'R',				'L',				'L',			'R',				'R',			'R',					'R',				'R',			'R',		'R',											'R', @ExcelSheetNo)

	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19, rr2, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',DB_NAME(fs.database_id) AS [Database], CAST(fs.io_stall_read_ms/(1.0 + fs.num_of_reads) AS NUMERIC(10,1)) AS [avg_read_latency_ms],  CAST(fs.io_stall_write_ms/(1.0 + fs.num_of_writes) AS NUMERIC(10,1)) AS [avg_write_latency_ms],  CAST((fs.io_stall_read_ms + fs.io_stall_write_ms)/(1.0 + fs.num_of_reads + fs.num_of_writes) AS NUMERIC(10,1)) AS [avg_io_latency_ms],  CONVERT(DECIMAL(18,2), mf.size/128.0) AS [File Size (MB)], mf.physical_name, mf.type_desc, fs.io_stall_read_ms, fs.num_of_reads,   fs.io_stall_write_ms, fs.num_of_writes, fs.io_stall_read_ms + fs.io_stall_write_ms AS [io_stalls], fs.num_of_reads + fs.num_of_writes AS [total_io],  io_stall_queued_read_ms AS [Resource Governor Total Read IO Latency (ms)], io_stall_queued_write_ms AS [Resource Governor Total Write IO Latency (ms)], (fs.io_stall_read_ms + fs.io_stall_write_ms)/(1.0 + fs.num_of_reads + fs.num_of_writes), @ExcelSheetNo   FROM sys.dm_io_virtual_file_stats(null,null) AS fs  INNER JOIN sys.master_files AS mf WITH (NOLOCK)  ON fs.database_id = mf.database_id  AND fs.[file_id] = mf.[file_id]  where DB_NAME(fs.database_id) not  in ('master','msdb','tempdb','distribution','DWQueue' ,'DWDiagnostics','model')   OPTION (RECOMPILE);
	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11) --specific
		SELECT  R1, R2, S1, R3,  R4,	C5,	C11, C9, C8,  C13,	C15,	C16 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31,  231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C11, C9, C8,  C13,	C15,	C16 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32 order by rr2 desc
		--ORDER BY TRY_CAST(C16 AS BIGINT) DESC
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		IF @HugeDataCounter<=@AllowFewToReport 
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'

		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----63---
    IF(@Parm8_13=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Total buffer usage by database for current instance (Buffer Usage by Database)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Buffer Pool Rank',	'Database',	'Cached Size (MB)',	'Buffer Pool Percent');
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S', 'R',				'L',		'R',				'R');
	WITH AggregateBufferPoolUsage  AS  (SELECT DB_NAME(database_id) AS [Database],  CAST(COUNT(*) * 8/1024.0 AS DECIMAL (10,2))  AS [CachedSize]  FROM sys.dm_os_buffer_descriptors WITH (NOLOCK)  WHERE database_id <> 32767   GROUP BY DB_NAME(database_id))  
	INSERT INTO ##RunOnce(R1, R2, S1, QuerySort, R3, R4,	C5,	C6,	C7,	C8)
	SELECT TOP 10 @QSectionNo, @QSectionSubNo, @QSectionSplitNo, ROW_NUMBER() OVER( ORDER BY CachedSize DESC), 32, 'D', ROW_NUMBER() OVER( ORDER BY CachedSize DESC) AS [Buffer Pool Rank], [Database], CachedSize AS [Cached Size (MB)],    CAST(CachedSize / SUM(CachedSize) OVER() * 100.0 AS DECIMAL(5,2)) AS [Buffer Pool Percent]  FROM AggregateBufferPoolUsage where [Database] not  in ('master','msdb','tempdb','distribution','DWQueue' ,'DWDiagnostics','model')    ORDER BY [Buffer Pool Rank] OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		---Logice goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----64---
    IF(@Parm8_14=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='CPU utilization by database'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1,QuerySort, R3, R4,	C5,	C6,	C7,	C8) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo,-1, 31, 'H',	'CPU Rank',	'Database',	'CPU Time (ms)',	'CPU Percent');
	INSERT INTO ##RunOnce(R1, R2, S1, QuerySort, R3, R4,	C5,	C6,	C7,	C8) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo,0, 231, 'S',	'R',	'L',	'R',	'R');
	WITH DB_CPU_Stats AS (
	SELECT pa.DatabaseID, DB_Name(pa.DatabaseID) AS [Database], SUM(qs.total_worker_time/1000) AS [CPU_Time_Ms]  FROM sys.dm_exec_query_stats AS qs WITH (NOLOCK)  CROSS APPLY (SELECT CONVERT(int, VALUE  ) AS [DatabaseID] FROM sys.dm_exec_plan_attributes(qs.plan_handle) WHERE attribute = N'dbid') AS pa  GROUP BY DatabaseID)  
	INSERT INTO ##RunOnce(R1, R2, S1,QuerySort, R3,  R4,	C5,	C6,	C7,	C8)
	SELECT TOP 10 @QSectionNo, @QSectionSubNo, @QSectionSplitNo, ROW_NUMBER() OVER( ORDER BY [CPU_Time_Ms] DESC) AS [CPU Rank], 32, 'D', ROW_NUMBER() OVER( ORDER BY [CPU_Time_Ms] DESC) AS [CPU Rank], [Database], [CPU_Time_Ms] AS [CPU Time (ms)], CAST([CPU_Time_Ms] * 1.0 / SUM([CPU_Time_Ms]) OVER() * 100.0 AS DECIMAL(5, 2)) AS [CPU Percent] FROM DB_CPU_Stats WHERE [Database] not  in ('master','msdb','tempdb','distribution','DWQueue' ,'DWDiagnostics','model')   order by CAST([CPU_Time_Ms] * 1.0 / SUM([CPU_Time_Ms]) OVER() * 100.0 AS DECIMAL(5, 2)) DESC

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		---Logice goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END


	--code for chart
	declare @databaseList varchar(max)='', @ResourceUsage varchar(max)='', @BarColor varchar(max)='"red", "green","blue","orange","brown"'
	DECLARE @cnt INT = 1;
	WHILE @cnt <= 10
	BEGIN
		SET @BarColor =@BarColor+', "rgb('+CAST(floor(rand()*100) + 150 as varchar)+','+CAST(floor(rand()*100) + 150 as varchar)+','+CAST(floor(rand()*100) + 150 as varchar)+')"' 
	   SET @cnt = @cnt + 1;
	END;
	;WITH DB_CPU_Stats AS (
	SELECT pa.DatabaseID, DB_Name(pa.DatabaseID) AS [Database], SUM(qs.total_worker_time/1000) AS [CPU_Time_Ms]  FROM sys.dm_exec_query_stats AS qs WITH (NOLOCK)  CROSS APPLY (SELECT CONVERT(int, VALUE  ) AS [DatabaseID] FROM sys.dm_exec_plan_attributes(qs.plan_handle) WHERE attribute = N'dbid')  AS pa where db_name([DatabaseID]) not  in ('master','msdb','tempdb','distribution','DWQueue' ,'DWDiagnostics','model')  GROUP BY DatabaseID)  
	SELECT top 10 
	@databaseList =@databaseList +', "'+ISNULL([Database],'')+'"', @ResourceUsage =@ResourceUsage +', '+ CAST(CAST([CPU_Time_Ms] * 1.0 / SUM([CPU_Time_Ms]) OVER() * 100.0 AS DECIMAL(5, 2)) as varchar) FROM DB_CPU_Stats  WHERE DatabaseID <> 32767 order by CAST([CPU_Time_Ms] * 1.0 / SUM([CPU_Time_Ms]) OVER() * 100.0 AS DECIMAL(5, 2)) DESC
	set @QSectionSplitNo=@QSectionSplitNo+1
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 33, 'chart',		
		'<canvas id="myChart" style="width:100%;max-width:600px"></canvas>
		<script>
		var xValues = ['+substring(@databaseList,2, len(@databaseList))+'];
		var yValues = ['+substring(@ResourceUsage,2, len(@ResourceUsage))+'];
		var barColors = ['+@BarColor+'];
		new Chart("myChart", {
		  type: "bar",
		  data: {
			labels: xValues,
			datasets: [{
			  backgroundColor: barColors,
			  data: yValues
			}]
		  },
		  options: {
			legend: {display: false},
			title: {
			  display: true,
			  text: "Database-wise CPU Utilization"
			},
		scales: {
				yAxes: [{
					display: true,
					ticks: {
						beginAtZero: true   
					}
				}]
			}
		  }
		});
		</script>'
			);

		--second


    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	-----65--
    IF(@Parm8_15=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Drive level latency information'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Drive',	'Volume Mount Point',	'Read Latency(ms)',	'Write Latency(ms)',	'Overall Latency(ms)',	'Avg Bytes/Read',	'Avg Bytes/Write',	'Avg Bytes/Transfer')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'R',	'R',	'R',	'R',	'R',	'R')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', tab.[Drive], 
	tab.volume_mount_point AS [Volume Mount Point], 		CASE WHEN num_of_reads = 0 THEN 0 	ELSE (io_stall_read_ms/num_of_reads) 	END AS [Read Latency],		CASE WHEN num_of_writes = 0 THEN 0 ELSE  io_stall_write_ms/num_of_writes END  AS [Write Latency],	 	CASE 		WHEN (num_of_reads = 0 AND num_of_writes = 0) THEN 0 		ELSE (io_stall/(num_of_reads + num_of_writes))  END AS [Overall Latency],		CASE 		WHEN num_of_reads = 0 THEN 0 		ELSE (num_of_bytes_read/num_of_reads) 	END AS [Avg Bytes/Read],	CASE 		WHEN num_of_writes = 0 THEN 0 		ELSE (num_of_bytes_written/num_of_writes) 	END AS [Avg Bytes/Write],	CASE 		WHEN (num_of_reads = 0 AND num_of_writes = 0) THEN 0 		ELSE ((num_of_bytes_read + num_of_bytes_written)/(num_of_reads + num_of_writes)) 	END AS [Avg Bytes/Transfer] FROM (SELECT LEFT(UPPER(mf.physical_name), 2) AS Drive, SUM(num_of_reads) AS num_of_reads,	    SUM(io_stall_read_ms) AS io_stall_read_ms, SUM(num_of_writes) AS num_of_writes,	    SUM(io_stall_write_ms) AS io_stall_write_ms, SUM(num_of_bytes_read) AS num_of_bytes_read,	    SUM(num_of_bytes_written) AS num_of_bytes_written, SUM(io_stall) AS io_stall, vs.volume_mount_point  FROM sys.dm_io_virtual_file_stats(NULL, NULL) AS vfs INNER JOIN sys.master_files AS mf WITH (NOLOCK) ON vfs.database_id = mf.database_id AND vfs.file_id = mf.file_id	  CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.[file_id]) AS vs  GROUP BY LEFT(UPPER(mf.physical_name), 2), vs.volume_mount_point) AS tab  ORDER BY [Overall Latency] OPTION (RECOMPILE);

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		UPDATE ##RunOnce SET C7='<span class=critical>'+C7+'</span>' WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
		AND TRY_CAST(C7 AS INT)>20

		UPDATE ##RunOnce SET C7='<span class=safe>'+C7+'</span>' WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
		AND TRY_CAST(C7 AS INT)<5

		UPDATE ##RunOnce SET C8='<span class=critical>'+C8+'</span>' WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
		AND TRY_CAST(C8 AS INT)>20

		UPDATE ##RunOnce SET C8='<span class=safe>'+C8+'</span>' WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
		AND TRY_CAST(C8 AS INT)<5

		UPDATE ##RunOnce SET C9='<span class=critical>'+C9+'</span>' WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
		AND TRY_CAST(C9 AS INT)>20

		UPDATE ##RunOnce SET C9='<span class=safe>'+C9+'</span>' WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
		AND TRY_CAST(C9 AS INT)<5

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		---Logice goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END

	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)>2 
	BEGIN
	----Chart setting
	/*
	--populate @DataForChart(id, x, y1, y2, y3)
	bring database name to xName and any three parameters to y1, y2, y3
	set the name of y1, y2 and y3 into variable @xParms like 'Read,write,overall'
	*/
	delete from @DataForChart
	INSERT INTO @DataForChart 
	SELECT top 5 
	tab.[Drive],  
	CASE WHEN num_of_reads = 0 THEN 0 	ELSE (io_stall_read_ms/num_of_reads) END AS [Read Latency],		
	CASE WHEN num_of_writes = 0 THEN 0 ELSE  io_stall_write_ms/num_of_writes END  AS [Write Latency],	 	
	CASE WHEN (num_of_reads = 0 AND num_of_writes = 0) THEN 0 ELSE (io_stall/(num_of_reads + num_of_writes))  END AS [Overall Latency]
	FROM (SELECT LEFT(UPPER(mf.physical_name), 2) AS Drive, SUM(num_of_reads) AS num_of_reads, SUM(io_stall_read_ms) AS io_stall_read_ms, SUM(num_of_writes) AS num_of_writes,	    SUM(io_stall_write_ms) AS io_stall_write_ms, SUM(num_of_bytes_read) AS num_of_bytes_read,	    SUM(num_of_bytes_written) AS num_of_bytes_written, SUM(io_stall) AS io_stall, vs.volume_mount_point  FROM sys.dm_io_virtual_file_stats(NULL, NULL) AS vfs INNER JOIN sys.master_files AS mf WITH (NOLOCK) ON vfs.database_id = mf.database_id AND vfs.file_id = mf.file_id	  
	CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.[file_id]) AS vs  
	GROUP BY LEFT(UPPER(mf.physical_name), 2), vs.volume_mount_point) AS tab  
	ORDER BY [Overall Latency] OPTION (RECOMPILE);

	set @xParms ='Read,write,overall'
	set @ChartTitle ='Drives latencies'

		SELECT @xValues = STRING_AGG('"' + REPLACE(value, ' ', '') + '"', ', ')
		FROM STRING_SPLIT(@xParms, ',');
		SELECT @Labels = STRING_AGG('row' + cast(id as varchar)+ '', ', ')
		FROM @DataForChart;
		-- Generate JavaScript output
		SET @outputScript=N'<div id=Div_'+cast(@QSectionNo as varchar)+cast(@QSectionSubNo as varchar)+ cast(@QSectionSplitNo as varchar)+'></div>'
		SET @outputScript =@outputScript+ '<script>'
		Select @outputScript =@outputScript +'var row'+cast(id as varchar) +'= {x: [' + @xValues + '], y: [' +cast(y1 as varchar)+ ','+cast(y2 as varchar)+','+cast(y3 as varchar)+'], name: "'+xName+'", type: "bar"};'
		from @DataForChart order by xName
		SET @outputScript =@outputScript +'var data = ['+@Labels+'];
		var layout = { title: "'+@ChartTitle+'", barmode: "group" };
		Plotly.newPlot("Div_'+cast(@QSectionNo as varchar)+cast(@QSectionSubNo as varchar)+ cast(@QSectionSplitNo as varchar)+'", data, layout);
		</script>';
		set @QSectionSplitNo=@QSectionSplitNo+1;
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5)
		SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 33, ' D', @outputScript 

	------ END OF CHART

	END

    END TRY
    BEGIN CATCH
	DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----66---
    IF(@Parm8_16=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Find missing index warnings for cached plans in each database (Missing Index Warnings)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset14 TABLE([Database]  VARCHAR(255),ObjectName  VARCHAR(255),objtype  VARCHAR(120),usecounts INT,size_in_bytes BIGINT /*,query_plan  VARCHAR(max)*/)
	INSERT INTO @Dataset14 
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) SELECT TOP(5) DB_NAME(DB_ID()) AS [Database], OBJECT_NAME(objectid) AS [ObjectName],       cp.objtype, cp.usecounts, cp.size_in_bytes /*, cast(qp.query_plan AS  VARCHAR(max))*/  FROM sys.dm_exec_cached_plans AS cp WITH (NOLOCK) CROSS APPLY sys.dm_exec_query_plan(cp.plan_handle) AS qp WHERE CAST(query_plan AS NVARCHAR(MAX)) LIKE N''%MissingIndex%'' AND dbid = DB_ID()  ORDER BY cp.usecounts DESC OPTION (RECOMPILE);'
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'ObjectName',	'objtype',	'usecounts',	'size_in_bytes')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', * FROM @Dataset14 
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		---Logice goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----67---
    IF(@Parm8_17=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='I/O utilization by database (IO Usage By Database)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	--huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Rank',	'Database',	'Total I/O (MB)',	'Total I/O %',	'Read I/O (MB)',	'Read I/O %',	'Write I/O (MB)',	'Write I/O %', @ExcelSheetNo);
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'R',	'L',	'R',	'R',	'R',	'R',	'R',	'R', @ExcelSheetNo);
	;WITH Aggregate_IO_Statistics AS ( 
		SELECT DB_NAME(database_id) AS [Database], CAST(SUM(num_of_bytes_read + num_of_bytes_written) / 1048576 AS DECIMAL(12, 2)) AS [ioTotalMB],     CAST(SUM(num_of_bytes_read ) / 1048576 AS DECIMAL(12, 2)) AS [ioReadMB],     CAST(SUM(num_of_bytes_written) / 1048576 AS DECIMAL(12, 2)) AS [ioWriteMB]     FROM sys.dm_io_virtual_file_stats(NULL, NULL) AS [DM_IO_STATS]     GROUP BY database_id) 
	INSERT INTO ##DataForSheet(R1, R2, S1,QuerySort,  R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, ROW_NUMBER() OVER ( ORDER BY ioTotalMB DESC),  32, 'D', ROW_NUMBER() OVER ( ORDER BY ioTotalMB DESC) AS [I/O Rank],    [Database], ioTotalMB AS [Total I/O (MB)],    CAST(ioTotalMB / SUM(ioTotalMB) OVER () * 100.0 AS DECIMAL(5, 2)) AS [Total I/O %],    ioReadMB AS [Read I/O (MB)],  		CAST(ioReadMB / SUM(ioReadMB) OVER () * 100.0 AS DECIMAL(5, 2)) AS [Read I/O %],    ioWriteMB AS [Write I/O (MB)],  		CAST(ioWriteMB / SUM(ioWriteMB) OVER () * 100.0 AS DECIMAL(5, 2)) AS [Write I/O %], @ExcelSheetNo FROM Aggregate_IO_Statistics where [Database] not  in ('master','msdb','tempdb','distribution','DWQueue' ,'DWDiagnostics','model')   ORDER BY [I/O Rank] OPTION (RECOMPILE);

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0  
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1,  R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12) --all
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, QuerySort, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, QuerySort, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32)
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		IF(@HugeDataCounter<=@AllowFewToReport)
			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		ELSE
		BEGIN
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5,C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
			Set @ExcelSheetNo=@ExcelSheetNo+1;
		END	
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END

	----Chart setting
/*
--populate @DataForChart(id, x, y1, y2, y3)
bring database name to xName and any three parameters to y1, y2, y3
set the name of y1, y2 and y3 into variable @xParms like 'Read,write,overall'
*/
	delete from @DataForChart
	INSERT INTO @DataForChart 
		SELECT TOP 10 C6, C10, C12, C8 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32) order by C8 desc

	set @xParms ='Read I/O %,Write I/O %,Total I/O %'
	set @ChartTitle ='I/O utilization by database (IO Usage By Database)'

SELECT @xValues = STRING_AGG('"' + REPLACE(value, ' ', '') + '"', ', ')
FROM STRING_SPLIT(@xParms, ',');
SELECT @Labels = STRING_AGG('row' + cast(id as varchar)+ '', ', ')
FROM @DataForChart;
-- Generate JavaScript output
SET @outputScript=N'<div id=Div_'+cast(@QSectionNo as varchar)+cast(@QSectionSubNo as varchar)+ cast(@QSectionSplitNo as varchar)+'></div>'
SET @outputScript =@outputScript+ '<script>'
Select @outputScript =@outputScript +'var row'+cast(id as varchar) +'= {x: [' + @xValues + '], y: [' +cast(y1 as varchar)+ ','+cast(y2 as varchar)+','+cast(y3 as varchar)+'], name: "'+xName+'", type: "bar"};'
from @DataForChart order by xName
SET @outputScript =@outputScript +'var data = ['+@Labels+'];
var layout = { title: "'+@ChartTitle+'", barmode: "group" };
Plotly.newPlot("Div_'+cast(@QSectionNo as varchar)+cast(@QSectionSubNo as varchar)+ cast(@QSectionSplitNo as varchar)+'", data, layout);
</script>';
set @QSectionSplitNo=@QSectionSplitNo+1;
INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5)
SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, ' D', @outputScript 

------ END OF CHART


    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
END -- END OF IF(@Database_Performance=0)

IF(@Index_optimization=0)
BEGIN
	SET @QSectionNo=@QSectionNo+1
	SET @QSectionSubNo=1
	SET @QSectionSplitNo=1
	SET @QTotalNo=@QTotalNo+1
    Set @QHeadingOne='Index statistics, missing index, fragmented index and optimizations'
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, -1, 'H1', CAST(@QSectionNo AS VARCHAR)+'. '+@QHeadingOne) 

	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0,'INTRO', '
	<div class=intro><font size=5><p>
	<b>Before reading ahead, Lets know about the importance of Index optimization</b></p></font>
	<div class=intro><font size=4>
	<b>Logical Read:</b> A logical read occurs every time the Database Engine requests a page from the buffer cache,	
	<br><b>Logical Write:</b> A logical write occurs when data is modified in a page in the buffer cache.
	<br><p><b><I>Resolve High Logical Read</B></I></p>
		<OL>
		<LI>Either reduce reads or increase physical memory.
		<LI>First being query design, 
		<LI>Secondly being indexing. If your query is pulling a large number of records, that could be filtered by getting a smaller set prior to pulling that data then you can always cut down on reads that way. With improved indexing, specifically with covering indexes,  you can reduce the number of pages that are being read as well.
		</OL>
		<p><U><I>By the way logical read is not a big issue compare to physical read that has significant impact on performance issue. Physical reads is the one takes significantly longer time.</U></I></P>
		<p><b>Two solutions (either improve your IO system or decrease writes to disk)</b></p>
		<UL>
			<LI>If we can''t increase IO system, decrease your write to disk by looking bad Non-clustered index and drop.
			<LI>Any un-necessary index which confronting write more than read need to be drop.
		</UL>
		<p><B>Physical Read:</B> Physical read indicates total number of data pages that are read from disk. </P>
		<P><B>Physical write:</B> It occurs when the page is written from the buffer cache to disk. </P>
		<P><B>Worker Time:</B> The CPU time taken by query execution.</P>
			<i><u>Common methods to resolve long-running, CPU-bound queries</U></I>
			<UL>
				<LI>Examine the query plan of the query.
				<LI>Update Statistics.
				<LI>Identify and apply Missing Indexes. ...
				<LI>Redesign or rewrite the queries.
			</UL>
		<P><B>Index statistics:</B> Statistics for query optimization are binary large objects (BLOBs), that contain statistical information about the distribution of values in one or more columns of a table 
		or indexed view. The Query Optimizer uses these statistics to estimate the cardinality, or number of rows, 
		in the query result.</P>
		<p><B>Fragmented Index:</B> when indexes have pages in which the logical ordering within the index, 
		based on the key values of the index, does not match the physical ordering of index pages.</P>
		<UL>
			<LI> When number of pages greater than 1000 than... 
			<OL>
				<LI>If avg. fragmentation is between 5% to 30%, than need index re-organize -- ALTER INDEX REORGANIZE
				<LI>If avg. fragmentation is greater than 30% than need index re-build -- ALTER INDEX REBUILD
			</OL>
		</UL>
	</p>
			<P><b>Troubleshooting</b></P>
		<p><b><I>Problem</I></B></P>
		<p>A query, especially those that perform high reads, can suddenly degrade or improve for no apparent reason. 
		This short technote provides some guidelines on how to determine if this is due to buffer pool hit rate, 
		or buffer pool threshing.</p>
		<P><b><I>Symptom</b></I></P>
		<p>The same query can range a few seconds, or become a few minutes with no changes to DB , DBM configuration 
		or application. Runstats does not improve the performance.</P>
		<P><b><I>Cause</b></I></P>
		<p>If most of the data/index pages required is already in the buffer pool, 
		then the query will complete quickly. (logical read)
		<br>Otherwise , depending on the amount of pages to be read in from disk, 
		the query can take much longer. (Physical read)</P>
		<p><b><I>Environment</b></I></P>
		<p>Any platform and any version</P>
		<br>
		</DIV>
	</font></div>'
	---68---
    IF(@Parm9_1=0)
    BEGIN
    BEGIN TRY
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QHeadingTw0='Index fragmentation information'
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')

	--splite data
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Schema',	'Object',	'Index Name',	'index_id',	'index_type_desc',	'avg_fragmentation(%)',	'fragment_count',	'page_count',	'fill_factor',	'has_filter',	'filter_definition',	'allow_page_locks', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'L',	'L',	'L',	'R',	'R',	'R',	'R',	'R',	'L',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM ##FragmentedIndexData

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 --AND 
	BEGIN
	/*
	SELECT * FROM (
SELECT ROW_NUMBER() OVER (Partition by ([Database]) Order by avg_fragmentation_in_percent desc) AS TOP_Frag,
* FROM ##FragmentedIndexData WHERE index_id>2
) TOP5 WHERE TOP_Frag <=10
ORDER BY avg_fragmentation_in_percent desc

	*/
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C8, C11, C12, C13 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31,  231)
		INSERT INTO ##RunOnce(R1, R2, S1, QuerySort, R3,  R4,	C5,	C6,	C7,	C8,	C9) 
		SELECT TOP(@AllowFewToReport) @QSectionNo, @QSectionSubNo, @QSectionSplitNo, ROW_NUMBER() OVER(Order by avg_fragmentation_in_percent desc), 32,'D', [Database], [Index Name], avg_fragmentation_in_percent, fragment_count, page_count  FROM (
			SELECT ROW_NUMBER() OVER (Partition by ([Database]) Order by avg_fragmentation_in_percent desc) AS TOP_Frag,
			* FROM ##FragmentedIndexData WHERE index_id>2
			) TOP5 WHERE TOP_Frag <=1
		ORDER BY avg_fragmentation_in_percent desc

--		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C8, CASE WHEN R3=32 THEN LEFT(C11, (CHARINDEX('.', C11, 1) + 2)) ELSE C11 END C11, C12, C13 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32--, 231)
	--	ORDER by C11 desc
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		IF(@HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	BEGIN TRY --fragmentation summary
		SET @QSectionSubNo=@QSectionSubNo+1;
		INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 11, 'H3', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'.'+CAST(@QSectionSplitNo AS VARCHAR)+'. Index fragmentation summary for reorganize/rebuild (Databases having fragmentation greater than 5) ') 
		IF(SELECT COUNT(*) FROM ##FragmentedIndexData )>0
		BEGIN
			IF (SELECT COUNT(*) FROM ##FragmentedIndexData FID WHERE page_count>1000 AND avg_fragmentation_in_percent>10 and [Index Name] NOT LIKE 'PXML%' AND [Index Name]  NOT LIKE 'PK%' )>0
			BEGIN
				INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
				INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
				INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Total Indexes',	'Healthy',	'Need Reorganize',	'Need Rebuild')
				INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'R',				'R',		'R',				'R')
				INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) 
				SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', [Database], 
					 (SELECT COUNT(*) from ##FragmentedIndexData WHERE [Database]=FID.[Database]) AS [Total Indexes],
					 (SELECT COUNT(*) from ##FragmentedIndexData WHERE [Database]=FID.[Database] AND avg_fragmentation_in_percent < 10) AS [Healthy Indexes],
					 (SELECT COUNT(*) from ##FragmentedIndexData WHERE [Database]=FID.[Database] AND (avg_fragmentation_in_percent BETWEEN 10 AND 30) AND page_count > 1000) AS [Need Reoranize],
					 (SELECT COUNT(*) from ##FragmentedIndexData WHERE [Database]=FID.[Database] AND avg_fragmentation_in_percent > 30 AND page_count > 1000) AS [Need Rebuild]
				FROM ##FragmentedIndexData FID WHERE avg_fragmentation_in_percent>0 GROUP BY [Database] ORDER BY [Need Rebuild] DESC, [Need Reoranize] DESC, [Total indexes] DESC
				INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')

				--Generate re-build and re-organize statements
					INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,C15, C16, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Schema',	'Object',	'Index Name',	'index_id',	'index_type_desc',	'avg_fragmentation(%)',	'fragment_count',	'page_count',	'Re-organize Statement',	'Re-build statement','All statment in one', @ExcelSheetNo)
					INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, ExclSheet)
					SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', 
					[database], [Schema], [Object], [Index Name], [index_id], [index_type_desc], [avg_fragmentation_in_percent], [fragment_count], [page_count], 
					'USE ['+[database]+'] ALTER INDEX ['+[Index Name]+'] ON ['+[Schema]+'].['+[Object]+'] REORGANIZE  WITH ( LOB_COMPACTION = ON )'
					, @ExcelSheetNo FROM ##FragmentedIndexData WHERE  [Index Name] NOT LIKE 'PXML%' AND [Index Name]  NOT LIKE 'PK%' AND page_count > 1000 AND (avg_fragmentation_in_percent BETWEEN 5 AND 30)

					INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C15, ExclSheet)
					SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', 
					[database], [Schema], [Object], [Index Name], [index_id], [index_type_desc], [avg_fragmentation_in_percent], [fragment_count], [page_count], 
					'USE ['+[database]+'] ALTER INDEX ['+[Index Name]+'] ON ['+[Schema]+'].['+[Object]+'] REBUILD PARTITION = ALL WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)', 
					@ExcelSheetNo FROM ##FragmentedIndexData WHERE [Index Name] NOT LIKE 'PXML%' AND [Index Name]  NOT LIKE 'PK%' AND page_count > 1000  AND avg_fragmentation_in_percent > 30 

					UPDATE ##DataForSheet SET C16=ISNULL(C14,'')+ISNULL(C15,'') WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32

					INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'Re-organize & re-build statments where required'
				Set @ExcelSheetNo=@ExcelSheetNo+1;
			END
			ELSE
			BEGIN 
				INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 41, '', '<span class=summary>Indexes are healthy </span>'
			END
		END
		ELSE
			INSERT INTO ##RunOnce (R1,R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0

	--don't need this table any more.
	IF OBJECT_ID(N'tempdb..##FragmentedIndexData') IS NOT NULL
		--SELECT * FROM ##FragmentedIndexData
		DROP TABLE ##FragmentedIndexData 
	END TRY
	BEGIN CATCH
	    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
	END CATCH
   END -- @Parm

	----69---
    IF(@Parm9_2=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='In-memory OLTP index usage (XTP Index Usage)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset23 TABLE(C5  VARCHAR(MAX),	C6  VARCHAR(MAX),	C7  VARCHAR(MAX),	C8  VARCHAR(MAX),	C9  VARCHAR(MAX),	C10  VARCHAR(MAX),	C11  VARCHAR(MAX),	C12  VARCHAR(MAX),	C13  VARCHAR(MAX))
	INSERT INTO @Dataset23
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) SELECT  db_name() AS dbname,  OBJECT_NAME(i.[object_id]) AS [Object], i.index_id, i.[name] AS [Index Name],  i.[type_desc], xis.scans_started, xis.scans_retries,  xis.rows_touched, xis.rows_returned FROM sys.dm_db_xtp_index_stats AS xis WITH (NOLOCK) INNER JOIN sys.indexes AS i WITH (NOLOCK) ON i.[object_id] = xis.[object_id]  AND i.index_id = xis.index_id   ORDER BY OBJECT_NAME(i.[object_id]) OPTION (RECOMPILE);'
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'dbname',	'Object',	'index_id',	'Index Name',	'type_desc',	'scans_started',	'scans_retries',	'rows_touched',	'rows_returned')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'L',	'L',	'R',	'R',	'R',	'R')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', * FROM @Dataset23
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		---Logice goes here	
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	---70----
    IF(@Parm9_3=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Lock waits for each database'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
--- 
-- dbname table_name index_name	index_id partition_number total_row_lock_waits	total_row_lock_wait_in_ms total_page_lock_waits	total_page_lock_wait_in_ms	total_lock_wait_in_ms

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset24 TABLE(dbname NVARCHAR(255),table_name NVARCHAR(max),index_name NVARCHAR(max),index_id INT,partition_number INT, total_row_lock_waits BIGINT, total_row_lock_wait_in_ms BIGINT, total_page_lock_waits BIGINT, total_page_lock_wait_in_ms BIGINT, total_lock_wait_in_ms bigint)
	INSERT INTO @Dataset24 
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) SELECT  db_name() AS dbname,  o.name AS [table_name], i.name AS [index_name], ios.index_id, ios.partition_number, 		SUM(ios.row_lock_wait_count) AS [total_row_lock_waits],  		SUM(ios.row_lock_wait_in_ms) AS [total_row_lock_wait_in_ms], 		SUM(ios.page_lock_wait_count) AS [total_page_lock_waits], 		SUM(ios.page_lock_wait_in_ms) AS [total_page_lock_wait_in_ms], 		SUM(ios.page_lock_wait_in_ms)+ SUM(row_lock_wait_in_ms) AS [total_lock_wait_in_ms] FROM sys.dm_db_index_operational_stats(DB_ID(), NULL, NULL, NULL) AS ios INNER JOIN sys.objects AS o WITH (NOLOCK) ON ios.[object_id] = o.[object_id] INNER JOIN sys.indexes AS i WITH (NOLOCK) ON ios.[object_id] = i.[object_id]  AND ios.index_id = i.index_id WHERE o.[object_id] > 100 GROUP BY o.name, i.name, ios.index_id, ios.partition_number HAVING SUM(ios.page_lock_wait_in_ms)+ SUM(row_lock_wait_in_ms) > 0  ORDER BY total_lock_wait_in_ms DESC OPTION (RECOMPILE);'		
	--keep as huge data and reduce fields on report
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'dbname',	'table_name',	'index_name',	'index_id',	'partition_number',	'total_row_lock_waits',	'total_row_lock_wait_in_ms',	'total_page_lock_waits',	'total_page_lock_wait_in_ms',	'total_lock_wait_in_ms', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'L',	'L',	'R',	'R',	'R',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset24 
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')

	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9) 
		SELECT R1, R2, S1, R3,  R4,	C5, C7, C11, C13, C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5, C7, C11, C13, C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (32)
		if(@HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END

    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----71----
    IF(@Parm9_4=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Indexes that are being maintained but not used (High Write/Zero Read) '
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	--Huge data handling
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'statement',	'index_name',	'user_reads',	'user_writes',	'total_rows',	'drop_command', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'R',	'R',	'R',	'L', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', '[' + DB_NAME() + '].[' + su.[name] + '].[' + o.[name] + ']'   AS [statement] ,    i.[name] AS [index_name] ,    ddius.[user_seeks] + ddius.[user_scans] + ddius.[user_lookups]    AS [user_reads] ,    ddius.[user_updates] AS [user_writes] ,    SUM(SP.rows) AS [total_rows],    'DROP INDEX [' + i.[name] + '] ON [' + su.[name] + '].[' + o.[name]    + '] WITH ( ONLINE = OFF )' AS [drop_command], @ExcelSheetNo FROM    sys.dm_db_index_usage_stats ddius    INNER JOIN sys.indexes i ON ddius.[object_id] = i.[object_id]        AND i.[index_id] = ddius.[index_id]    INNER JOIN sys.partitions SP ON ddius.[object_id] = SP.[object_id]           AND SP.[index_id] = ddius.[index_id]    INNER JOIN sys.objects o ON ddius.[object_id] = o.[object_id]    INNER JOIN sys.sysusers su ON o.[schema_id] = su.[UID] WHERE   OBJECTPROPERTY(ddius.[object_id], 'IsUserTable') = 1    AND ddius.[index_id] > 0 GROUP BY su.[name] ,    o.[name] ,    i.[name] ,    ddius.[user_seeks] + ddius.[user_scans] + ddius.[user_lookups] , ddius.[user_updates]  HAVING  ddius.[user_seeks] + ddius.[user_scans] + ddius.[user_lookups] = 0   ORDER BY ddius.[user_updates] DESC ,     su.[name] ,     o.[name] ,     i.[name] 

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')

	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 --AND  
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32)
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		IF(@HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5,C5, C6) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf', 'Note: Commands for index drop are on sheet'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5,C5, C6) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr', 'Note: Commands for index drop are on sheet'
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----72---
    IF(@Parm9_5=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Detailed activity information for indexes not used for user reads'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	--huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15, ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'statement',	'index_name',	'user_reads',	'user_writes',	'leaf_INSERT_count',	'leaf_delete_count',	'leaf_update_count',	'nonleaf_INSERT_count',	'nonleaf_delete_count',	'nonleaf_update_count',	'drop_command', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15, ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'L', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', '[' + DB_NAME() + '].[' + su.[name] + '].[' + o.[name] + ']'    AS [statement] ,    i.[name] AS [index_name] ,    ddius.[user_seeks] + ddius.[user_scans] + ddius.[user_lookups]   AS [user_reads] ,    ddius.[user_updates] AS [user_writes] ,   ddios.[leaf_INSERT_count] ,    ddios.[leaf_delete_count] ,    ddios.[leaf_update_count] ,    ddios.[nonleaf_INSERT_count] ,    ddios.[nonleaf_delete_count] ,    ddios.[nonleaf_update_count],    'DROP INDEX [' + i.[name] + '] ON [' + su.[name] + '].[' + o.[name]   + '] WITH ( ONLINE = OFF )' AS [drop_command], @ExcelSheetNo FROM    sys.dm_db_index_usage_stats ddius    INNER JOIN sys.indexes i ON ddius.[object_id] = i.[object_id]       AND i.[index_id] = ddius.[index_id]    INNER JOIN sys.partitions SP ON ddius.[object_id] = SP.[object_id]           AND SP.[index_id] = ddius.[index_id]    INNER JOIN sys.objects o ON ddius.[object_id] = o.[object_id]    INNER JOIN sys.sysusers su ON o.[schema_id] = su.[UID]    INNER JOIN sys.[dm_db_index_operational_stats](DB_ID(), NULL, NULL,           NULL)    AS ddios        ON ddius.[index_id] = ddios.[index_id]      AND ddius.[object_id] = ddios.[object_id]      AND SP.[partition_number] = ddios.[partition_number]      AND ddius.[database_id] = ddios.[database_id] WHERE OBJECTPROPERTY(ddius.[object_id], 'IsUserTable') = 1  AND ddius.[index_id] > 0  AND ddius.[user_seeks] + ddius.[user_scans] + ddius.[user_lookups] = 0  ORDER BY ddius.[user_updates] DESC ,    su.[name] ,    o.[name] ,    i.[name] 
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
		
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5, C6, C7, C8, C9) 
		SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H', 'statement', 'index_name', 'user [reads-writes]', 'leaf count[insert - delete - update]', 'non leaf count[insert - delete - update]'
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5, C6, C7, C8, C9) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7+' - '+C8,	C9+' - '+C10+' - '+C11,	C12+' - '+C13+' - '+C14+' - '+C15 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32)
		IF(@HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	-----73---
    IF(@Parm9_6=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Pathways for performance improvement'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	
	--huge data
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12, C13, C14, ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'index_advantage',	'last_user_seek',	'Database.Schema.Table',	'equality_columns',	'inequality_columns',	'included_columns',	'unique_compiles',	'user_seeks',	'avg_total_user_cost',	'avg_user_impact', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12, C13, C14, ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'R',	'R',	'L',	'L',	'L',	'L',	'R',	'R',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12, C13, C14, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', user_seeks * avg_total_user_cost * ( avg_user_impact * 0.01 ) AS [index_advantage], CONVERT(VARCHAR, dbmigs.last_user_seek, 20), dbmid.[statement] AS [Database.Schema.Table], dbmid.equality_columns, dbmid.inequality_columns, dbmid.included_columns, dbmigs.unique_compiles, dbmigs.user_seeks, dbmigs.avg_total_user_cost, dbmigs.avg_user_impact, @ExcelSheetNo FROM    sys.dm_db_missing_index_group_stats AS dbmigs WITH ( NOLOCK ) INNER JOIN sys.dm_db_missing_index_groups AS dbmig WITH ( NOLOCK ) ON dbmigs.group_handle = dbmig.index_group_handle    INNER JOIN sys.dm_db_missing_index_details AS dbmid WITH ( NOLOCK ) ON dbmig.index_handle = dbmid.index_handle WHERE   dbmid.[database_id] = DB_ID()  ORDER BY index_advantage DESC ; 

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32	
	IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, C9, C10) 
		SELECT  R1, R2, S1, R3,  R4,	C5, C7, C6, C10, C13, C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31,  231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, C9, C10) 
		SELECT TOP (@AllowFewToReport) R1, R2, S1, R3,  R4,	C5, C7, C6, C10, C13, C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32 
		IF(@HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE 
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;	
		END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----74---
    IF(@Parm9_7=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Index Read/Write stats, ordered by Reads. (Overall Index Usage - Reads)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset21 TABLE(C5  VARCHAR(MAX),	C6  VARCHAR(MAX),	C7  VARCHAR(MAX),	C8  VARCHAR(MAX),	C9  VARCHAR(MAX),	C10  VARCHAR(MAX),	C11  VARCHAR(MAX),	C12  decimal(15,2),	C13  VARCHAR(MAX),	C14  VARCHAR(MAX),	C15  VARCHAR(MAX),	C16  VARCHAR(MAX),	C17  VARCHAR(MAX),	C18  VARCHAR(MAX),	C19  VARCHAR(MAX),	C20  VARCHAR(MAX))
	INSERT INTO @Dataset21
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) SELECT  db_name() AS dbname, OBJECT_NAME(i.[object_id]) AS [ObjectName], i.[name] AS [IndexName], i.index_id,    s.user_seeks, s.user_scans, s.user_lookups, 	   s.user_seeks + s.user_scans + s.user_lookups AS [Total Reads],  	   s.user_updates AS [Writes],   	   i.[type_desc] AS [Index Type], i.fill_factor AS [Fill Factor], i.has_filter, i.filter_definition,  	   s.last_user_scan, CONVERT(VARCHAR, s.last_user_lookup, 20), CONVERT(VARCHAR, s.last_user_seek, 20)  FROM sys.indexes AS i WITH (NOLOCK) LEFT OUTER JOIN sys.dm_db_index_usage_stats AS s WITH (NOLOCK) ON i.[object_id] = s.[object_id] AND i.index_id = s.index_id AND s.database_id = DB_ID() WHERE OBJECTPROPERTY(i.[object_id],''IsUserTable'') = 1  ORDER BY s.user_seeks + s.user_scans + s.user_lookups DESC OPTION (RECOMPILE); '
	--split data done

	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'dbname',	'ObjectName',	'IndexName',	'index_id',	'user_seeks',	'user_scans',	'user_lookups',	'Total Reads',	'Writes',	'Index Type',	'Fill Factor',	'has_filter',	'filter_definition',	'last_user_scan',	'last_user_lookup',	'last_user_seek', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'L',	'R',	'R',	'R',	'R',	'R',	'L',	'R',	'R',	'L',	'R',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20, RR2, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, ds.C12, @ExcelSheetNo FROM @Dataset21 ds order by ds.C12 desc
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
 	IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, C9,	C10) 
		SELECT  R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C12, C13, C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, C9,	C10) 
		SELECT TOP(@AllowFewToReport ) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C12, C13, C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32) ORDER BY RR2 DESC
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		IF(@HugeDataCounter<=@AllowFewToReport )
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE 
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;

	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----75---
    IF(@Parm9_8=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Index Read/Write stats, ordered by Writes. (Overall Index Usage - Writes)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset22 TABLE(C5  VARCHAR(MAX),	C6  VARCHAR(MAX),	C7  VARCHAR(MAX),	C8  VARCHAR(MAX),	C9  decimal(15,2),	C10  VARCHAR(MAX),	C11  VARCHAR(MAX),	C12  VARCHAR(MAX),	C13  VARCHAR(MAX),	C14  VARCHAR(MAX),	C15  VARCHAR(MAX),	C16  VARCHAR(MAX))
	INSERT INTO @Dataset22
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) SELECT db_name() AS dbname,  OBJECT_NAME(i.[object_id]) AS [ObjectName], i.[name] AS [IndexName], i.index_id, 	   s.user_updates AS [Writes], s.user_seeks + s.user_scans + s.user_lookups AS [Total Reads],  	   i.[type_desc] AS [Index Type], i.fill_factor AS [Fill Factor], i.has_filter, i.filter_definition, 	   s.last_system_update, s.last_user_update FROM sys.indexes AS i WITH (NOLOCK) LEFT OUTER JOIN sys.dm_db_index_usage_stats AS s WITH (NOLOCK) ON i.[object_id] = s.[object_id] AND i.index_id = s.index_id AND s.database_id = DB_ID() WHERE OBJECTPROPERTY(i.[object_id],''IsUserTable'') = 1  ORDER BY s.user_updates DESC OPTION (RECOMPILE);'
	--huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'dbname',	'ObjectName',	'IndexName',	'index_id',	'Writes',	'Total Reads',	'Index Type',	'Fill Factor',	'has_filter',	'filter_definition',	'last_system_update',	'last_user_update', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'L',	'R',	'R',	'L',	'R',	'R',	'R',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset22 order by C9 desc

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12) 
		SELECT  R1, R2, S1, R3,  R4,	C5,	C7,	C9,	C10, C11,	C12, C13,	C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31,  231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12) 
		SELECT TOP (@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C7,	C9,	C10, C11,	C12, C13,	C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32 
		ORDER BY C9 desc
		IF(@HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE 
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;	
		END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	---76--
    IF(@Parm9_10=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Missing Indexes for all databases by Index Advantage'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'DatabaseName',	'index_advantage',	'last_user_seek',	'Database.Schema.Table',	'missing_indexes_for_table',	'similar_missing_indexes_for_table',	'equality_columns',	'inequality_columns',	'included_columns',	'user_seeks',	'avg_total_user_cost',	'avg_user_impact', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'R',	'R',	'L',	'L',	'L',	'L',	'L',	'L',	'R',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, QuerySort, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, ROW_NUMBER() OVER(ORDER BY CONVERT(decimal(18,2), migs.user_seeks * migs.avg_total_user_cost * (migs.avg_user_impact * 0.01)) DESC), 32, 'D', DatabaseName=DB_NAME(mid.database_id), CONVERT(decimal(18,2), migs.user_seeks * migs.avg_total_user_cost * (migs.avg_user_impact * 0.01)) AS [index_advantage], FORMAT(migs.last_user_seek, 'yyyy-MM-dd HH:mm:ss') AS [last_user_seek],  mid.[statement] AS [Database.Schema.Table], COUNT(1) OVER(PARTITION BY mid.[statement]) AS [missing_indexes_for_table], COUNT(1) OVER(PARTITION BY mid.[statement], equality_columns) AS [similar_missing_indexes_for_table], mid.equality_columns, mid.inequality_columns, mid.included_columns, migs.user_seeks,  CONVERT(decimal(18,2), migs.avg_total_user_cost) AS [avg_total_user_cost], migs.avg_user_impact, @ExcelSheetNo  FROM sys.dm_db_missing_index_group_stats AS migs WITH (NOLOCK) INNER JOIN sys.dm_db_missing_index_groups AS mig WITH (NOLOCK) ON migs.group_handle = mig.index_group_handle INNER JOIN sys.dm_db_missing_index_details AS mid WITH (NOLOCK) ON mig.index_handle = mid.index_handle ORDER BY index_advantage DESC OPTION (RECOMPILE);
	--splite data
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9)
		SELECT R1, R2, S1, R3, R4,	C7,	 C6, C9, C10, C16 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, QuerySort, R3, R4,	C5,	C6,	C7,	C8,	C9)
		SELECT TOP (@AllowFewToReport) R1, R2, S1, QuerySort, R3, R4,	C7,	 C6, C9, C10, C16 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32

		IF(@HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE 
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'

		Set @ExcelSheetNo=@ExcelSheetNo+1;	
		END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	-----77---
    IF(@Parm9_9=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Look at most frequently modified indexes and statistics'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset19 TABLE(R4  VARCHAR(MAX), C5  VARCHAR(MAX),	C6  VARCHAR(MAX),	C7  VARCHAR(MAX),	C8  VARCHAR(MAX),	C9  VARCHAR(MAX),	C10  VARCHAR(MAX),	C11  VARCHAR(MAX),	C12  VARCHAR(MAX),	C13  VARCHAR(MAX),	C14  VARCHAR(MAX),	C15  VARCHAR(MAX),	C16  VARCHAR(MAX),	C17  VARCHAR(MAX))
	INSERT INTO @Dataset19
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) SELECT DB_NAME(DB_ID()) [Database], o.[name] AS [Object], o.[object_id], o.[type_desc], s.[name] AS [Statistics Name],    s.stats_id, s.no_recompute, s.auto_created, s.is_incremental, s.is_temporary, 	   sp.modification_counter, sp.[rows], sp.rows_sampled, sp.last_updated FROM sys.objects AS o WITH (NOLOCK) INNER JOIN sys.stats AS s WITH (NOLOCK) ON s.object_id = o.object_id CROSS APPLY sys.dm_db_stats_properties(s.object_id, s.stats_id) AS sp WHERE o.[type_desc] NOT IN (N''SYSTEM_TABLE'', N''INTERNAL_TABLE'') AND sp.modification_counter > 0  ORDER BY sp.modification_counter DESC, o.name OPTION (RECOMPILE);'
	--Huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Object',	'object_id',	'type_desc',	'Statistics Name',	'stats_id',	'no_recompute',	'auto_created',	'is_incremental',	'is_temporary',	'modification_counter',	'rows',	'rows_sampled',	'last_updated', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'L',	'L',	'L',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset19
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8) 
		SELECT  R1, R2, S1, R3,  R4,	C5,	C9,	C15, C18 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)		
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C9,	C15, C18 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32)		
		IF(@HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'

		Set @ExcelSheetNo=@ExcelSheetNo+1;

	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----78---
/*    IF(@Parm9_11=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Missing Indexes for each database by Index Advantage (Missing Indexes)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset13 TABLE(dbase_name  VARCHAR(255),index_advantage BIGINT,last_user_seek VARCHAR(max) ,[Database.Schema.Table]  VARCHAR(255),equality_columns  VARCHAR(255),inequality_columns  VARCHAR(255),included_columns  VARCHAR(max),user_seeks INT,avg_total_user_cost INT,avg_user_impact INT,[Table]  VARCHAR(255),[Table Rows] bigint)
	INSERT INTO @Dataset13 
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) SELECT DISTINCT db_name() AS dbase_name, CONVERT(decimal(18,2), migs.user_seeks * migs.avg_total_user_cost * (migs.avg_user_impact * 0.01)) AS [index_advantage],  CONVERT(VARCHAR, migs.last_user_seek, 20), mid.[statement] AS [Database.Schema.Table], mid.equality_columns, mid.inequality_columns, mid.included_columns, migs.user_seeks, migs.avg_total_user_cost, migs.avg_user_impact, OBJECT_NAME(mid.[object_id]) AS [Table], p.rows AS [Table Rows] FROM sys.dm_db_missing_index_group_stats AS migs WITH (NOLOCK) INNER JOIN sys.dm_db_missing_index_groups AS mig WITH (NOLOCK) ON migs.group_handle = mig.index_group_handle INNER JOIN sys.dm_db_missing_index_details AS mid WITH (NOLOCK) ON mig.index_handle = mid.index_handle INNER JOIN sys.partitions AS p WITH (NOLOCK) ON p.[object_id] = mid.[object_id] WHERE mid.database_id = DB_ID() AND p.index_id < 2   ORDER BY index_advantage DESC OPTION (RECOMPILE)';
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'dbase_name',	'index_advantage',	'last_user_seek',	'Database.Schema.Table',	'equality_columns',	'inequality_columns',	'included_columns',	'user_seeks',	'avg_total_user_cost',	'avg_user_impact',	'Table',	'Table Rows')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', * FROM @Dataset13 
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		---Logice goes here	
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH

END -- @Parm */
	
	-----79---
    IF(@Parm9_12=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Bad non-clustered Indexes where writes are greater than reads'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset12 TABLE([Database]  VARCHAR(128),[Schema]  VARCHAR(80),[Table]  VARCHAR(255),[Index Name]  VARCHAR(255),index_id INT, is_disabled tinyint, is_hypothetical tinyint, has_filter tinyint, fill_factor tinyint, [Total Writes] BIGINT,[Total Reads] BIGINT,Difference bigint)
	INSERT INTO @Dataset12 
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model''))  SELECT DB_NAME(DB_ID()) AS [Database], SCHEMA_NAME(o.[schema_id]) AS [Schema],  OBJECT_NAME(s.[object_id]) AS [Table], i.name AS [Index Name], i.index_id,  i.is_disabled, i.is_hypothetical, i.has_filter, i.fill_factor, s.user_updates AS [Total Writes], s.user_seeks + s.user_scans + s.user_lookups AS [Total Reads], s.user_updates - (s.user_seeks + s.user_scans + s.user_lookups) AS [Difference] FROM sys.dm_db_index_usage_stats AS s WITH (NOLOCK) INNER JOIN sys.indexes AS i WITH (NOLOCK) ON s.[object_id] = i.[object_id] AND i.index_id = s.index_id INNER JOIN sys.objects AS o WITH (NOLOCK) ON i.[object_id] = o.[object_id] WHERE OBJECTPROPERTY(s.[object_id],''IsUserTable'') = 1 AND s.database_id = DB_ID() AND s.user_updates > (s.user_seeks + s.user_scans + s.user_lookups) AND i.index_id > 1 AND i.[type_desc] = N''NONCLUSTERED'' AND i.is_primary_key = 0 AND i.is_unique_constraint = 0 AND i.is_unique = 0  ORDER BY [Difference] DESC, [Total Writes] DESC, [Total Reads] ASC OPTION (RECOMPILE)';
	--huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Schema',	'Table',	'Index Name',	'index_id',	'is_disabled',	'is_hypothetical',	'has_filter',	'fill_factor',	'Total Writes',	'Total Reads',	'Difference', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',		'L',		'L',			'L',		'R',			'R',				'R',			'R',			'R',			'R',			'R',			@ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset12 
	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4, C5, C6, C7,	C8,	C9) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C8, C14, C15, C16 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4, C5, C6, C7,	C8,	C9) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C8, C14, C15, C16 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32)
		IF( @HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5,C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'		
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm	
   	BEGIN TRY ---BAD NC SUMMARY
	IF (SELECT COUNT(*) FROM @Dataset12)>0
		IF (SELECT COUNT(*) FROM @Dataset12 WHERE [Difference]>0)>0
		BEGIN
			SET @QSectionSplitNo=@QSectionSplitNo+1;
			INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 11, 'H3', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'.'+CAST(@QSectionSplitNo AS VARCHAR)+'. Summarized bad non-clustered indexes') 
			INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
			INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 25,'FS','Total Wr -g Rd: Total number of indexes where write greater than read')
			INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 25,'FS','Zero Read: Total number of indexes hit for write and zero read')
			INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 25,'FS','Wr-Rd -g 5xRd: (Write-Read)>Five times read')
			INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 25,'FS','Wr-Rd -g Rd: (Write-Read)>read')
			INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
			INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
			INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8, C9) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Total Wr -g Rd',	'Zero Read',	'Wr-Rd -g 5xRd',	'Wr-Rd -g Rd')
			INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8, C9) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'R',	 'R',	'R',		'R')
			INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9) 
			SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', [Database],
 				(SELECT COUNT([index_id]) from @Dataset12 where [Database]= DS.[Database] AND [Total Writes]>[Total Reads]) AS TWGR,
 				(SELECT COUNT([index_id]) from @Dataset12 where [Database]= DS.[Database] AND [Total Writes]>0 AND [Total Reads]=0) AS TWZR,
				(SELECT COUNT([index_id]) from @Dataset12 where [Database]= DS.[Database] AND [Difference]>(5*[Total Reads])) AS WDRG5R,
				(SELECT COUNT([index_id]) from @Dataset12 where [Database]= DS.[Database] AND [Difference]>[Total Reads]) AS WDRGR
			FROM @Dataset12 DS group by [Database]
	END
	ELSE
		BEGIN 
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 41, '', '<span class=summary>Bad NC indexes are not really bad! </span>'
		END
	END TRY
	BEGIN CATCH
	    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
	END CATCH
	----80---
    IF(@Parm9_13=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Statistics last updated?  (Statistics Update)'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset18 TABLE([Database]  VARCHAR(255), [Object]  VARCHAR(255), [Object Type]  VARCHAR(255), [Index Name]  VARCHAR(255), [Statistics Date] datetime, [auto_created] tinyint, no_recompute tinyint, user_created INT, is_incremental INT, is_temporary INT,row_count BIGINT, used_page_count int)
	INSERT INTO @Dataset18
	EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) SELECT  DB_NAME(DB_ID()) [Database], DB_NAME(DB_ID()) + N''.'' + SCHEMA_NAME(o.Schema_ID) + N''.'' + o.[NAME] AS [Object], o.[type_desc] AS [Object Type],  i.[name] AS [Index Name], 	 STATS_DATE(i.[object_id], i.index_id)   AS [Statistics Date],   	s.auto_created, s.no_recompute, s.user_created, s.is_incremental, s.is_temporary, 	  st.row_count, st.used_page_count FROM sys.objects AS o WITH (NOLOCK) INNER JOIN sys.indexes AS i WITH (NOLOCK) ON o.[object_id] = i.[object_id] INNER JOIN sys.stats AS s WITH (NOLOCK) ON i.[object_id] = s.[object_id]  AND i.index_id = s.stats_id INNER JOIN sys.dm_db_partition_stats AS st WITH (NOLOCK) ON o.[object_id] = st.[object_id] AND i.[index_id] = st.[index_id] 	WHERE o.[type] IN (''U'', ''V'') AND st.row_count > 0  ORDER BY STATS_DATE(i.[object_id], i.index_id) DESC OPTION (RECOMPILE);'

	--huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'Object',	'Object Type',	'Index Name',	'Statistics Date',	'auto_created',	'no_recompute',	'user_created',	'is_incremental',	'is_temporary',	'row_count',	'used_page_count', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'L',	'L',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset18
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
				
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0  
	BEGIN
	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7) 
		SELECT R1, R2, S1, R3,  R4, C6, C8, C9	FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
	UPDATE ##DataForSheet SET C9= 
		REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(C9,'Jan ','01/'),'Feb ','02/'),'Mar ','03/'),'Apr ', '04/'),'May ', '05/'),'Jun ', '06/') ,'Jul ', '07/') ,'Aug ', '08/') ,'Sep ', '09/') ,'Oct ', '10/') ,'Nov ', '11/') ,'Dec ', '12/')  
	WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
	UPDATE ##DataForSheet SET C9= 
		REPLACE(REPLACE(C9,' 20','/20'),' 19','/19')
	WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32

	UPDATE ##DataForSheet SET C9= 
		REPLACE(REPLACE(C9,'AM',' AM'),'PM',' PM')
	WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32


	INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4, C6, C8, ISNULL(C9,'No Updates')	FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32

		IF(@HugeDataCounter<=@AllowFewToReport )
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
END		-- END OF IF(@Index_optimization=0)

IF(@Queries_and_Stored_Procedures=0)
BEGIN
	SET @QSectionNo=@QSectionNo+1
	SET @QSectionSubNo=1
	SET @QSectionSplitNo=1
	SET @QTotalNo=@QTotalNo+1
    Set @QHeadingOne='Most frequently executing and resource utilization''s queries and stored procedures'
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, -1, 'H1', CAST(@QSectionNo AS VARCHAR)+'. '+@QHeadingOne) 

    -----81---
	IF(@Parm10_1=0)
	BEGIN
    BEGIN TRY 
    SET @QHeadingTw0='<span id='+CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'></span>'
	SET @QHeadingTw0=@QHeadingTw0+'Top average elapsed time queries for entire instance (Top Avg Elapsed Time Queries)'
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	--huge data done
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, C15, C16, ExclSheet)  VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H', 'Database',	 'Avg Elapsed Time',	'min_elapsed_time',	'max_elapsed_time',	'last_elapsed_time',	'Exec count',	'Avg Logical Reads',	'Avg Physical Reads',	'Avg Worker Time',	'Missing Index?',	'Creation Time','Link (query/plan)', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, C15, C16, ExclSheet)  VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S', 'L',	 'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R','C',  @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C39,	C40, H2, C38, RR2, ExclSheet)
    SELECT TOP(25) @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', 
		DB_NAME(t.[dbid]) AS [Database], qs.total_elapsed_time/qs.execution_count AS [Avg Elapsed Time], qs.min_elapsed_time, qs.max_elapsed_time, qs.last_elapsed_time, qs.execution_count AS [Exec count],   qs.total_logical_reads/qs.execution_count AS [Avg Logical Reads],  qs.total_physical_reads/qs.execution_count AS [Avg Physical Reads],  qs.total_worker_time/qs.execution_count AS [Avg Worker Time], 
		CASE WHEN CONVERT(NVARCHAR(max), qp.query_plan) LIKE N'%<MissingIndexes>%' THEN 1 ELSE 0 END AS [Missing Index?], 
		CONVERT(VARCHAR, qs.creation_time, 20)  AS [Creation Time],  t.[text] AS [Query Text], CONVERT(NVARCHAR(max), qp.query_plan) AS [Query Plan],'QueryPlan','Elapsed Time', ROW_NUMBER() over(order by qs.total_elapsed_time/qs.execution_count DESC), @ExcelSheetNo 
		FROM sys.dm_exec_query_stats AS qs WITH (NOLOCK) CROSS APPLY sys.dm_exec_sql_text(plan_handle) AS t  CROSS APPLY sys.dm_exec_query_plan(plan_handle) AS qp 
		where DB_NAME(t.[dbid]) not  in ('master','msdb','tempdb','distribution','DWQueue' ,'DWDiagnostics','model')   and t.[text] not like '(@%'
		ORDER BY qs.total_elapsed_time/qs.execution_count DESC OPTION (RECOMPILE);
	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')


	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
				
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, C6, C7, C8, C9, C10, C11, C12)
		SELECT R1, R2, S1, R3, R4, C5, C6, C7, C8, C9, C10, C11, C16 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, C6, C7, C8, C9, C10, C11, C12)
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3, R4, C5, C6, C7,C8,C9,C10, C15, '<a href=#'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'>'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'</a>'  FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32) order by RR2
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		--DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo
		IF(@HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;			
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	-----91--
    IF(@Parm10_11=0)
	BEGIN
    BEGIN TRY 
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
    SET @QHeadingTw0='<span id='+CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'></span>'
	SET @QHeadingTw0=@QHeadingTw0+'Top logical read time queries for entire instance (Pressure on physical memory)'
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	--huge data done
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18, C19, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',		'Total Logical Reads',	'Min Logical Reads',	'Avg Logical Reads',	'Max Logical Reads',	'Min Worker Time',	'Avg Worker Time',	'Max Worker Time',	'Min Elapsed Time',	'Avg Elapsed Time',	'Max Elapsed Time',	'Exec count',	'Missing Index?',	'Creation Time', 'Link(query/plan)',@ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18, C19, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',		'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R','C', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C39, C40, H2, C38, RR2, ExclSheet, C19)
    SELECT TOP(25) @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', 
	DB_NAME(t.[dbid]) AS [Database],  
	  qs.total_logical_reads AS [Total Logical Reads], qs.min_logical_reads AS [Min Logical Reads], qs.total_logical_reads/qs.execution_count AS [Avg Logical Reads], qs.max_logical_reads AS [Max Logical Reads],    qs.min_worker_time AS [Min Worker Time], qs.total_worker_time/qs.execution_count AS [Avg Worker Time],  qs.max_worker_time AS [Max Worker Time],  qs.min_elapsed_time AS [Min Elapsed Time],  qs.total_elapsed_time/qs.execution_count AS [Avg Elapsed Time],  qs.max_elapsed_time AS [Max Elapsed Time], qs.execution_count AS [Exec count],  CASE WHEN CONVERT(NVARCHAR(max), qp.query_plan) LIKE N'%<MissingIndexes>%' THEN 1 ELSE 0 END AS [Missing Index?], CONVERT(VARCHAR, qs.creation_time, 20)  AS [Creation Time], t.[text]  AS [Query Text], convert(varchar(max),qp.query_plan), 'QueryPlan','Memory', ROW_NUMBER() over(order by qs.total_logical_reads/qs.execution_count desc), @ExcelSheetNo
	  , '<a href=#'+cast(ROW_NUMBER() over(order by t.[dbid]) as varchar)+'>'+cast(ROW_NUMBER() over(order by t.[dbid]) as varchar)+'</a>'  
	  FROM sys.dm_exec_query_stats AS qs WITH (NOLOCK) CROSS APPLY sys.dm_exec_sql_text(plan_handle) AS t  CROSS APPLY sys.dm_exec_query_plan(plan_handle) AS qp 
	  WHERE DB_NAME(t.[dbid]) not  in ('master','msdb','tempdb','distribution','DWQueue' ,'DWDiagnostics','model') and t.[text] not like '(@%'  ORDER BY qs.total_logical_reads/qs.execution_count  DESC OPTION (RECOMPILE);

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5, H2)	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D','Queries text on excel sheet','1'			
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32

	IF @HugeDataCounter>0 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11) 
		SELECT  R1, R2, @QSectionSplitNo, R3, R4, C5,	C6,	C7,	C8,	C9, C18, C19 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11) 
		SELECT TOP(@AllowFewToReport) R1, R2, @QSectionSplitNo, R3, R4, C5,	C6,	C7,	C8,	C9, C18, '<a href=#'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'>'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'</a>'  FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32) order by rr2 

		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		IF(@HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END 
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
    
	----92---
    IF(@Parm10_12=0)
	BEGIN
    BEGIN TRY 
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
    SET @QHeadingTw0='<span id='+CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'></span>'
	SET @QHeadingTw0=@QHeadingTw0+'Top total worker time queries for entire instance (Pressure on CPU)'
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	--huge data done
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12, C13, C14, C15,	C16,	C17,	C18, C19,  ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',		'Total Worker Time',	'Min Worker Time',	'Avg Worker Time',	'Max Worker Time',	'Min Elapsed Time',	'Avg Elapsed Time',	'Max Elapsed Time',	'Min Logical Reads',	'Avg Logical Reads',	'Max Logical Reads',	'Exec count',	'Missing Index?',	'Creation Time','Link(query/plan)', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,  C19, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',		'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R', 'C', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C39, H2, C38, RR2, ExclSheet)
    SELECT TOP(25) @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', 
		DB_NAME(t.[dbid]) AS [Database],    qs.total_worker_time AS [Total Worker Time], qs.min_worker_time AS [Min Worker Time], qs.total_worker_time/qs.execution_count AS [Avg Worker Time],  qs.max_worker_time AS [Max Worker Time],  qs.min_elapsed_time AS [Min Elapsed Time],  qs.total_elapsed_time/qs.execution_count AS [Avg Elapsed Time],  qs.max_elapsed_time AS [Max Elapsed Time], qs.min_logical_reads AS [Min Logical Reads], qs.total_logical_reads/qs.execution_count AS [Avg Logical Reads], qs.max_logical_reads AS [Max Logical Reads],  qs.execution_count AS [Exec count], CASE WHEN CONVERT(NVARCHAR(max), qp.query_plan) LIKE N'%<MissingIndexes>%' THEN 1 ELSE 0 END AS [Missing Index?],  CONVERT(VARCHAR, qs.creation_time, 20)  AS [Creation Time], t.[text]  AS [Query Text],'Query','CPU High',ROW_NUMBER() over(order by qs.total_worker_time/qs.execution_count desc), @ExcelSheetNo   
		FROM sys.dm_exec_query_stats AS qs WITH (NOLOCK) CROSS APPLY sys.dm_exec_sql_text(plan_handle) AS t  CROSS APPLY sys.dm_exec_query_plan(plan_handle) AS qp 
		where DB_NAME(t.[dbid]) not  in ('master','msdb','tempdb','distribution','DWQueue' ,'DWDiagnostics','model') and t.[text] not like  '(@%'
		ORDER BY qs.total_worker_time/qs.execution_count  DESC OPTION (RECOMPILE);

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5, H2)	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 132, 'D','Visit excel sheet for query text','1'			
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0  
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12, C13, C14) 
		SELECT  R1, R2, @QSectionSplitNo, R3, R4, C5, C6, C7, C8, C9,  C15, C16, C17, C18, C19 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12, C13, C14) 
		SELECT TOP(@AllowFewToReport) R1, R2, @QSectionSplitNo, R3, R4, C5, C6, C7, C8, C9,  C15, C16, C17, C18, '<a href=#'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'>'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'</a>' FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32) order by RR2

		IF @HugeDataCounter<=@AllowFewToReport
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm

	----83---
	IF(@Parm10_3=0)
	BEGIN
    BEGIN TRY 
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
    SET @QHeadingTw0='<span id='+CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'></span>'
	SET @QHeadingTw0=@QHeadingTw0+'Most frequently executed queries for each database'
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset4 TABLE([Database]  VARCHAR(255),[Query Text]  VARCHAR(MAX), [Query Plan] varchar(max), [Exec count] BIGINT,[Total Logical Reads] BIGINT,[Avg Logical Reads] BIGINT,[Total Worker Time] BIGINT,[Avg Worker Time] BIGINT,[Total Elapsed Time] BIGINT,[Avg Elapsed Time] BIGINT,[Missing Index?] TINYINT,[Creation Time] DATETIME)
    set @cmd='USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model''))  
			SELECT TOP(10) DB_NAME(t.dbid) AS [Database],  t.[text]  AS [Query Text], CONVERT(NVARCHAR(max), qp.query_plan) as [Query Plan], qs.execution_count AS [Exec count], qs.total_logical_reads AS [Total Logical Reads], qs.total_logical_reads/qs.execution_count AS [Avg Logical Reads], qs.total_worker_time AS [Total Worker Time], qs.total_worker_time/qs.execution_count AS [Avg Worker Time],  qs.total_elapsed_time AS [Total Elapsed Time], qs.total_elapsed_time/qs.execution_count AS [Avg Elapsed Time], CASE WHEN CONVERT(NVARCHAR(max), qp.query_plan) LIKE N''%<MissingIndexes>%'' THEN 1 ELSE 0 END AS [Missing Index?], CONVERT(VARCHAR, qs.creation_time, 20)  AS [Creation Time] 
					FROM sys.dm_exec_query_stats AS qs WITH (NOLOCK) CROSS APPLY sys.dm_exec_sql_text(plan_handle) AS t  CROSS APPLY sys.dm_exec_query_plan(plan_handle) AS qp  
					WHERE t.dbid = DB_ID() and t.[text] not like ''(@%%'' ORDER BY qs.execution_count DESC OPTION (RECOMPILE)'
    INSERT INTO @Dataset4 
    EXEC sp_MSforeachdb @cmd ;
	
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, C15,  ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',		'Exec count',	'Total Logical Reads',	'Avg Logical Reads',	'Total Worker Time',	'Avg Worker Time',	'Total Elapsed Time',	'Avg Elapsed Time',	'Missing Index?',	'Creation Time','Link(query/plan)', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, C15,  ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R', 'C', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C39, C40, H2, C38, RR2, ExclSheet)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', [Database],	 [Exec count], [Total Logical Reads], [Avg Logical Reads], [Total Worker Time], [Avg Worker Time], [Total Elapsed Time], [Avg Elapsed Time], [Missing Index?], [Creation Time], [Query Text], [query plan], 'Query', 'High count execution', ROW_NUMBER() over(order by [Exec count] desc), @ExcelSheetNo  FROM @Dataset4
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
				
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0  
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C8,	C10, C12, C13, C14, C15) 
		SELECT R1, R2, S1, R3,  R4,		C5,	C6,	C8,	C10, C12, C13, C14, C15 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C8,	C10, C12, C13, C14, C15) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C8,	C10, C12, C13, C14, '<a href=#'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'>'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'</a>'   FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (32) order by rr2
		IF(@HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm

  ----84---
	IF(@Parm10_4=0)
	BEGIN
    BEGIN TRY 
    SET @QSectionSubNo=@QSectionSubNo+1 
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
    SET @QHeadingTw0='<span id='+CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'></span>'
	SET @QHeadingTw0=@QHeadingTw0+'Find single-use, ad-hoc and prepared queries that are bloating the plan cache '
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
 	--huge data done
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8, C9, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',		'Object Type',	'Cache Object Type',	'Plan Size in KB', 'Link(query/plan)', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8, C9, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',		'L',	'L',	'R','C', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C39, H2, C38, RR2, ExclSheet)
	SELECT TOP(25) @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', 
		DB_NAME(t.[dbid]) AS [Database], cp.objtype AS [Object Type], cp.cacheobjtype AS [Cache Object Type],   cp.size_in_bytes/1024 AS [Plan Size in KB], t.[text] AS [Query Text],  'Query', 'Ad-hoc bloating plan cashe', ROW_NUMBER() over(order by cp.size_in_bytes desc), @ExcelSheetNo 
		FROM sys.dm_exec_cached_plans AS cp WITH (NOLOCK) CROSS APPLY sys.dm_exec_sql_text(plan_handle) AS t WHERE DB_NAME(t.[dbid]) not in('master','msdb','tempdb', 'model') and t.[text] not like '(@%' AND  cp.cacheobjtype = N'Compiled Plan'  AND cp.objtype IN (N'Adhoc', N'Prepared')  AND cp.usecounts = 1 ORDER BY cp.size_in_bytes DESC, DB_NAME(t.[dbid]) OPTION (RECOMPILE);

	--select 'D', DB_NAME(t.[dbid]) AS [Database], t.[text] AS [Query Text],  cp.objtype AS [Object Type], cp.cacheobjtype AS [Cache Object Type],   cp.size_in_bytes/1024 AS [Plan Size in KB] FROM sys.dm_exec_cached_plans AS cp WITH (NOLOCK) CROSS APPLY sys.dm_exec_sql_text(plan_handle) AS t WHERE DB_NAME(t.[dbid]) not in('master') AND  cp.cacheobjtype = N'Compiled Plan'  AND cp.objtype IN (N'Adhoc', N'Prepared')  AND cp.usecounts = 1 ORDER BY cp.size_in_bytes DESC, DB_NAME(t.[dbid]) OPTION (RECOMPILE);
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
				
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, C9) 
		SELECT R1, R2, S1, R3,  R4,	C5 , C6, C7, C8, C9 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, C9) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5, C6, C7,	C8,	'<a href=#'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'>'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'</a>'  FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (32) order by RR2
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		--query text espcial case (can't show whole query on report
		IF @HugeDataCounter>@AllowFewToReport
		begin 
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
			Set @ExcelSheetNo=@ExcelSheetNo+1;
		end
--		else
--			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	----82---
	IF(@Parm10_2=0)
	BEGIN
    BEGIN TRY	
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
    SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
	SET @QHeadingTw0='<span id='+CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'></span>'
	SET @QHeadingTw0=@QHeadingTw0+'Top stored procedures by average input/output (IO) usage for each database (Pressure on storage disk)'
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    DECLARE  @Dataset11 TABLE([Database]  VARCHAR(255),[SP Name]  VARCHAR(255),[Avg IO] INT,[Exec count] INT,[Query Text]  VARCHAR(max))
    set @cmd='USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model''))   
				SELECT TOP(10) DB_NAME(DB_ID()) AS [Database],  OBJECT_NAME(qt.objectid, dbid) AS [SP Name], (qs.total_logical_reads + qs.total_logical_writes) /qs.execution_count AS [Avg IO], qs.execution_count AS [Exec count],   qt.[text] AS [Query Text]	 FROM sys.dm_exec_query_stats AS qs WITH (NOLOCK) CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) AS qt WHERE qt.[dbid] = DB_ID()  ORDER BY [Avg IO] DESC OPTION (RECOMPILE)'
	INSERT INTO @Dataset11 
    EXEC sp_MSforeachdb @cmd ;
	--huge data done
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8, C9,	 ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'SP Name',	'Avg IO',	'Exec count',	'Link', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8, C9,	 ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'R',	'R',	'C', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C39, H2, C38, RR2, ExclSheet)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, 'Query', 'IO', ROW_NUMBER() over(order by [Database]), @ExcelSheetNo FROM @Dataset11 order by [Avg IO]  desc
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
				
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0  
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, C9) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, C9 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, C9) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, '<a href=#'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'>'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'</a>'
		FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN (32) order by RR2
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
		IF @HugeDataCounter>@AllowFewToReport
			BEGIN
				INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
				Set @ExcelSheetNo=@ExcelSheetNo+1;
			END
		ELSE
			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
	
	

	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
    

	----85----
    IF(@Parm10_5=0)
	BEGIN
    BEGIN TRY 
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
	SET @QHeadingTw0='<span id='+CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'></span>'
	SET @QHeadingTw0=@QHeadingTw0+'Top cached stored procedures by avg elapsed time'
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    DECLARE  @Dataset6 TABLE([Database]  VARCHAR(255),[SP Name]  VARCHAR(255),[min_elapsed_time] BIGINT,[avg_elapsed_time] BIGINT,[max_elapsed_time] BIGINT,[last_elapsed_time] BIGINT,[total_elapsed_time] BIGINT,[execution_count] BIGINT, [Calls/Minute] BIGINT, [AvgWorkerTime] BIGINT, [TotalWorkerTime] BIGINT, [Missing Index?] TINYINT,[Last Exection] VARCHAR(MAX),[query text] varchar(max), [Plan Cached] VARCHAR(MAX) )
    INSERT INTO @Dataset6 --,''master'',''msdb'',''model''
    EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model''))  
						SELECT TOP(25) DB_NAME(DB_ID()) AS [Database], p.name AS [SP Name], qs.min_elapsed_time, qs.total_elapsed_time/qs.execution_count AS [avg_elapsed_time],  qs.max_elapsed_time, qs.last_elapsed_time, qs.total_elapsed_time, qs.execution_count,  ISNULL(qs.execution_count/DATEDIFF(Minute, qs.cached_time, GETDATE()), 0) AS [Calls/Minute],  qs.total_worker_time/qs.execution_count AS [AvgWorkerTime],  qs.total_worker_time AS [TotalWorkerTime], CASE WHEN CONVERT(NVARCHAR(max), qp.query_plan) LIKE N''%<MissingIndexes>%'' THEN 1 ELSE 0 END AS [Missing Index?], FORMAT(qs.last_execution_time, ''yyyy-MM-dd HH:mm:ss'', ''en-US'') AS [Last Exection],  FORMAT(qs.cached_time, ''yyyy-MM-dd HH:mm:ss'', ''en-US'') AS [Plan Cached], qt.[text] 
							FROM sys.procedures AS p WITH (NOLOCK) 
							INNER JOIN sys.dm_exec_procedure_stats AS qs WITH (NOLOCK) ON p.[object_id] = qs.[object_id] 
							CROSS APPLY sys.dm_exec_query_plan(qs.plan_handle) AS qp 
							CROSS APPLY sys.dm_exec_sql_text(qs.plan_handle) AS qt
							WHERE qs.database_id = DB_ID() AND DATEDIFF(Minute, qs.cached_time, GETDATE()) > 0  ORDER BY qs.total_elapsed_time/qs.execution_count DESC OPTION (RECOMPILE)';
    --huge data done
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17, C18, C19,	 ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'SP Name',	'min elapsed(ms)',	'avg_elapsed(ms)',	'max_elapsed(ms)',	'last_elapsed(ms)',	'total_elapsed(ms)',	'exec_count',	'Calls/Minute',	'Avg Worker(ms)',	'Total Worker(ms)',	'Missing Index?',	'Last Execution', 'Plan chached',	'Link(query/plan', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17, C18,C19,	 ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R','R', 'C',  @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17, C18, C39,  H2, C38, RR2, ExclSheet)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *,'Query', 'IO', ROW_NUMBER() over(order by avg_elapsed_time desc), @ExcelSheetNo FROM @Dataset6
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
				
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12) 
		SELECT  R1, R2, S1, R3,  R4, C5, C6, C8, C11, C16, C17, C18, C19  FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31,  231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12) 
		SELECT TOP (@AllowFewToReport) R1, R2, S1, R3,  R4,C5, C6, C8, C11, C16, C17, C18, '<a href=#'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'>'+cast(R1 as varchar)+'.'+cast(R2 as varchar)+'.'+cast(RR2 as varchar)+'</a>' FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32 
		IF(@HugeDataCounter<=@AllowFewToReport)
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rf'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;

	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm

    ----86----
    IF(@Parm10_6=0)
	BEGIN
    BEGIN TRY 
    SET @QHeadingTw0='Top cached stored procedures by execution count'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset5 TABLE([Database]  VARCHAR(255),[SP Name]  VARCHAR(255),[Exec count] INT,[Calls/Minute] INT,[Avg Elapsed Time] INT,[Avg Worker Time] INT,[Avg Logical Reads] INT,[Missing Index?] TINYINT,[Last Exection] VARCHAR(MAX),[Plan Cached] VARCHAR(MAX))
    INSERT INTO @Dataset5
    EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model''))  
				SELECT TOP(100) DB_NAME(DB_ID()) AS [Database], p.name AS [SP Name], qs.execution_count AS [Exec count], ISNULL(qs.execution_count/DATEDIFF(Minute, qs.cached_time, GETDATE()), 0) AS [Calls/Minute], qs.total_elapsed_time/qs.execution_count AS [Avg Elapsed Time], qs.total_worker_time/qs.execution_count AS [Avg Worker Time],     qs.total_logical_reads/qs.execution_count AS [Avg Logical Reads], CASE WHEN CONVERT(NVARCHAR(max), qp.query_plan) LIKE N''%<MissingIndexes>%'' THEN 1 ELSE 0 END AS [Missing Index?], FORMAT(qs.last_execution_time, ''yyyy-MM-dd HH:mm:ss'', ''en-US'') AS [Last Exection],  FORMAT(qs.cached_time, ''yyyy-MM-dd HH:mm:ss'', ''en-US'') AS [Plan Cached] 
				FROM sys.procedures AS p WITH (NOLOCK) 
				INNER JOIN sys.dm_exec_procedure_stats AS qs WITH (NOLOCK) ON p.[object_id] = qs.[object_id] 
				CROSS APPLY sys.dm_exec_query_plan(qs.plan_handle) AS qp 
				WHERE qs.database_id = DB_ID() AND DATEDIFF(Minute, qs.cached_time, GETDATE()) > 0  
				ORDER BY qs.execution_count DESC OPTION (RECOMPILE)';
	--huge data -done
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'SP Name',	'Exec count',	'Calls/Minute',	'Avg Elapsed(ms)',	'Avg Worker(ms)',	'Avg Logical(ms)',	'Has Missing?',	'Last Execution',	'Plan Cached', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, ExclSheet)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset5
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
				
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, C12, C13, C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10, C11) 
		SELECT TOP (@AllowFewToReport)  R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8, C12, C13, C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
		IF @HugeDataCounter<=@AllowFewToReport 
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
		ELSE
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rfr'
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	ELSE IF @HugeDataCounter=0
	BEGIN ----st
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm

    -----87---
    IF(@Parm10_7=0)
	BEGIN
    BEGIN TRY 
    SET @QHeadingTw0='Top Cached stored procedures by Total Logical Reads. Logical reads relate to memory pressure '
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    DECLARE  @Dataset8 TABLE([Database]  VARCHAR(255),[SP Name]  VARCHAR(255),[Min] BIGINT, [Max] BIGINT, [Average] BIGINT, [Total] BIGINT, execution_count BIGINT, [Missing Index?] tinyint,[Last Exection] VARCHAR(MAX),[Plan Cached] VARCHAR(MAX))
    INSERT INTO @Dataset8 --------
    EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model''))  SELECT TOP(25) DB_NAME(DB_ID()) AS [Database], p.name AS [SP Name],qs.min_logical_reads [Min], qs.max_logical_reads [Max], qs.total_logical_reads/qs.execution_count AS [Average], qs.total_logical_reads AS [Total], qs.execution_count [Exec Count], CASE WHEN CONVERT(NVARCHAR(max), qp.query_plan) LIKE N''%<MissingIndexes>%'' THEN 1 ELSE 0 END AS [Missing Index?],  FORMAT(qs.last_execution_time,  ''yyyy-MM-dd HH:mm:ss'', ''en-US'') AS [Last Execution],  FORMAT(qs.cached_time, ''yyyy-MM-dd HH:mm:ss'', ''en-US'') AS [Plan Cached] FROM sys.procedures AS p WITH (NOLOCK) INNER JOIN sys.dm_exec_procedure_stats AS qs WITH (NOLOCK) ON p.[object_id] = qs.[object_id] CROSS APPLY sys.dm_exec_query_plan(qs.plan_handle) AS qp WHERE qs.database_id = DB_ID() AND DATEDIFF(Minute, qs.cached_time, GETDATE()) > 0  ORDER BY qs.total_logical_reads DESC OPTION (RECOMPILE)';
    --huge data done
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13 , C14, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'SP Name',	'Min',	'Max',	'Average',	'Total',	'Exec Count',	'Index Missing',	'Last Execution',	'Plan Cached', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13 , C14, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'R',	'R',	'R',	'R',	'R',	'R',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13, C14, ExclSheet)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset8

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
				
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 --AND 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13, C14) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13, C14) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32)
		IF @HugeDataCounter<=@AllowFewToReport 
			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		ELSE
		BEGIN
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
			Set @ExcelSheetNo=@ExcelSheetNo+1;
		END
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm

    ----88---
    IF(@Parm10_8=0)
	BEGIN
    BEGIN TRY 
    SET @QHeadingTw0='Top Cached stored procedures by Total Logical Writes'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    DECLARE  @Dataset10 TABLE([Database]  VARCHAR(255),[SP Name]  VARCHAR(255),[Min] BIGINT, [Max] BIGINT, [Average] BIGINT, [Total] BIGINT,execution_count INT, [Missing Index?] tinyint, [Last Exection] VARCHAR(MAX),[Plan Cached] VARCHAR(MAX))
    INSERT INTO @Dataset10 --
    EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model''))  SELECT TOP(25) DB_NAME(DB_ID()) AS [Database], p.name AS [SP Name], qs.min_logical_writes [Min], qs.max_logical_writes [Max], qs.total_logical_writes/qs.execution_count AS [Average], qs.total_logical_writes AS [Total], qs.execution_count [Exec Count], CASE WHEN CONVERT(NVARCHAR(max), qp.query_plan) LIKE N''%<MissingIndexes>%'' THEN 1 ELSE 0 END AS [Missing Index?],  FORMAT(qs.last_execution_time, ''yyyy-MM-dd HH:mm:ss'', ''en-US'') AS [Last Exection],  FORMAT(qs.cached_time, ''yyyy-MM-dd HH:mm:ss'', ''en-US'') AS [Plan Cached] FROM sys.procedures AS p WITH (NOLOCK) INNER JOIN sys.dm_exec_procedure_stats AS qs WITH (NOLOCK) ON p.[object_id] = qs.[object_id] CROSS APPLY sys.dm_exec_query_plan(qs.plan_handle) AS qp WHERE qs.database_id = DB_ID() AND qs.total_logical_writes > 0 AND DATEDIFF(Minute, qs.cached_time, GETDATE()) > 0  ORDER BY qs.total_logical_writes DESC OPTION (RECOMPILE)';
    --huge data done
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	 ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database', 'SP Name', 'Min', 'Max', 'Average', 'Total', 'Exec Count', 'Missing Index?', 'Last Execution', 'Plan Cached', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	 ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L', 'L', 'R', 'R', 'R', 'R', 'R', 'R', 'R', 'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	 ExclSheet)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset10
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
				
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0  
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14  FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14  FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32)
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		IF @HugeDataCounter<=@AllowFewToReport
			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		ELSE
		BEGIN
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
			Set @ExcelSheetNo=@ExcelSheetNo+1;
		END
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
    
	----89----
    IF(@Parm10_9=0)
	BEGIN
    BEGIN TRY 
    SET @QHeadingTw0='Top Cached stored procedures by Total Physical Reads. Physical reads relate to disk read I/O pressure'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    DECLARE  @Dataset9 TABLE([Database]  VARCHAR(255),[SP Name]  VARCHAR(255), [Min] BIGINT, [Max] BIGINT, [Average] BIGINT, Total BIGINT, execution_count BIGINT, [Missing Index?] tinyint, [Last Exection] VARCHAR(MAX),[Plan Cached] VARCHAR(MAX))
    INSERT INTO @Dataset9 
    EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model''))  
							SELECT TOP(25) DB_NAME(DB_ID()) AS [Database], p.name AS [SP Name], qs.min_physical_reads, qs.max_physical_reads, qs.total_physical_reads/qs.execution_count AS [AvgPhysicalReads], qs.total_physical_reads, qs.execution_count, CASE WHEN CONVERT(NVARCHAR(max), qp.query_plan) LIKE N''%<MissingIndexes>%'' THEN 1 ELSE 0 END AS [Missing Index?],FORMAT(qs.last_execution_time, ''yyyy-MM-dd HH:mm:ss'', ''en-US'') AS [Last Exection], FORMAT(qs.cached_time, ''yyyy-MM-dd HH:mm:ss'', ''en-US'')  AS [Plan Cached] 
									FROM sys.procedures AS p WITH (NOLOCK)INNER JOIN sys.dm_exec_procedure_stats AS qs WITH (NOLOCK)ON p.[object_id] = qs.[object_id]CROSS APPLY sys.dm_exec_query_plan(qs.plan_handle) AS qp 
									WHERE qs.database_id = DB_ID() AND qs.total_physical_reads > 0 
									ORDER BY qs.total_physical_reads DESC, qs.total_logical_reads DESC OPTION (RECOMPILE)';
    --huge data done
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database', 'SP Name', 'Min', 'max', 'Average', 'Total', 'Exec Count', 'Missing Index?', 'Last Execution', 'Plan Cached', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L', 'L', 'R', 'R', 'R', 'R', 'R', 'R', 'R', 'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14, ExclSheet)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset9 

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
				
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13,	C14 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32)
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		IF @HugeDataCounter<=@AllowFewToReport 
			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		ELSE
		BEGIN
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
			Set @ExcelSheetNo=@ExcelSheetNo+1;
		END
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END

    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
    
	----90----
    IF(@Parm10_10=0)
	BEGIN
	BEGIN TRY 
	SET @QHeadingTw0='Top Cached stored procedures by Total Worker time. Worker time relates to CPU cost'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
    DECLARE  @Dataset7 TABLE([Database]  VARCHAR(255),[SP Name]  VARCHAR(255), [Min] BIGINT, [Max] BIGINT, [Average] BIGINT, [Total] BIGINT,  [execution_count] BIGINT, [Missing Index?] TINYINT, [Last Exection] VARCHAR(MAX), [Plan Cached] VARCHAR(MAX))
    INSERT INTO @Dataset7
    EXEC sp_MSforeachdb 'USE [?] if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model''))  
						SELECT TOP(25) DB_NAME(DB_ID()) AS [Database], p.name AS [SP Name],qs.min_worker_time, qs.max_worker_time,   qs.total_worker_time/qs.execution_count AS [AvgWorkerTime], qs.total_worker_time, qs.execution_count,  CASE WHEN CONVERT(NVARCHAR(max), qp.query_plan) LIKE N''%<MissingIndexes>%'' THEN 1 ELSE 0 END AS [Missing Index?], FORMAT(qs.last_execution_time, ''yyyy-MM-dd HH:mm:ss'', ''en-US'') AS [Last Exection],  FORMAT(qs.cached_time, ''yyyy-MM-dd HH:mm:ss'', ''en-US'') AS [Plan Cached] 
								FROM sys.procedures AS p WITH (NOLOCK) INNER JOIN sys.dm_exec_procedure_stats AS qs WITH (NOLOCK) ON p.[object_id] = qs.[object_id] CROSS APPLY sys.dm_exec_query_plan(qs.plan_handle) AS qp 
								WHERE qs.database_id = DB_ID() AND DATEDIFF(Minute, qs.cached_time, GETDATE()) > 0  
								ORDER BY qs.total_worker_time DESC OPTION (RECOMPILE)';
    --huge data done
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12, C13, C14, ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Database',	'SP Name',	'Min',	'Max', 'Average', 'Total',	'Exec Count', 'Missing Index?',	'Last Exection',	'Plan Cached', @ExcelSheetNo)
    INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12, C13, C14, ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'R',	'R', 'R', 'R',	'R', 'R',	'R',	'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10, C11, C12, C13, C14, ExclSheet)
    SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',*, @ExcelSheetNo FROM @Dataset7 order by [Total] desc
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
				
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0  
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6, C9,	C10,	C11,	C12, C13,	C14,	C15 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12, C13) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C9,	C10,	C11,	C12, C13,	C14,	C15 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (32)
		/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
		IF @HugeDataCounter<=@AllowFewToReport 
			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		ELSE
		BEGIN
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
			Set @ExcelSheetNo=@ExcelSheetNo+1;
		END
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	
	END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
END -- END OF IF(@Queries_and_Stored_Procedures=0)

IF(@Database_login_users_roles_and_permissions=0)
BEGIN
	SET @QSectionNo=@QSectionNo+1
	SET @QSectionSubNo=1
	SET @QSectionSplitNo=1
	SET @QTotalNo=@QTotalNo+1
    Set @QHeadingOne='Database credentials and securties check'
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, -1, 'H1', CAST(@QSectionNo AS VARCHAR)+'. '+@QHeadingOne) 

	----93---
    IF(@Parm11_1=0)
    BEGIN
    BEGIN TRY
    SET @QHeadingTw0='List of logins in SQL Server'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, C5, C6, C7, C8, C9, C10) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'Login Name',	'Account Type',	'create_date',	'modify_date',	'default_database_name',	'default_language_name')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9,	C10)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', name AS [Login Name], type_desc AS [Account Type], CONVERT(VARCHAR, create_date, 20), CONVERT(VARCHAR, modify_date, 20), default_database_name, default_language_name
	FROM sys.server_principals WHERE TYPE IN ('U', 'S', 'G') --and name not like '%##%' --  ORDER BY name, type_desc
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
	INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		---Logice goes here	
					INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	----94----
    IF(@Parm11_2=0)
    BEGIN
    BEGIN TRY
    SET @QHeadingTw0='List of users in SQL Server'
    SET @QSectionSubNo=@QSectionSubNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')

	DECLARE  @Dataset30 TABLE(C5  VARCHAR(MAX),	C6  VARCHAR(MAX),	C7  VARCHAR(MAX),	C8  VARCHAR(MAX),	C9  VARCHAR(MAX),	C10  VARCHAR(MAX))
	INSERT INTO @Dataset30
	EXEC sp_MSforeachdb 'USE [?] SELECT db_name() AS [database], name AS username, SID, CONVERT(VARCHAR, create_date, 20),  CONVERT(VARCHAR, modify_date, 20), type_desc AS type FROM sys.database_principals WHERE type not in (''A'', ''G'', ''R'', ''X'') and sid is not null and name != ''guest'''
	--huge data done
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9, C10, ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'database',	'username','SID',	'create_date',	'modify_date',	'type', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9, C10, ExclSheet) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L','C',	'R',	'R',	'L', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8,	C9, C10, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset30	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
				
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
 IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9,	C10 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
	IF @HugeDataCounter>@AllowFewToReport
		BEGIN
			INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
			Set @ExcelSheetNo=@ExcelSheetNo+1;
		END
	ELSE
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm
	
	----95---
    IF(@Parm11_3=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='List of roles in SQL Server'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=' Query.  '+cast(@QSectionNo AS  VARCHAR)+' List of users in SQL Server, Timestamp. '+CONVERT( VARCHAR, GETDATE(),20) 
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	set @Sql_exec ='USE [?] SELECT '+Cast(@QSectionNo AS  VARCHAR) +','+Cast(@QSectionSubNo AS  VARCHAR) +','+Cast(@QSectionSplitNo AS  VARCHAR) +', 32, ''D'', db_name() AS [database],			[name], CONVERT(VARCHAR, create_date, 20), CONVERT(VARCHAR, modify_date, 20), '+cast(@ExcelSheetNo as varchar)+' FROM sys.database_principals WHERE type = ''R''  ORDER BY [name]'
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'database',	'name',	'create_date',						'modify_date', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'L',	'R',						'R', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7,	C8, ExclSheet)
		EXEC sp_MSforeachdb @Sql_exec
	--huge data done
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
				
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
 IF @HugeDataCounter>0 --AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8) 
		SELECT TOP(@AllowFewToReport) R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
		IF @HugeDataCounter>@AllowFewToReport
			BEGIN
				INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
				Set @ExcelSheetNo=@ExcelSheetNo+1;
			END
		ELSE
			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END

    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	-----96-------
    IF(@Parm11_4=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Fix SQL Server orphaned users'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset32 TABLE(UserName  VARCHAR(MAX),	UserSID  VARCHAR(MAX))
	INSERT INTO @Dataset32
	EXEC sp_MSforeachdb 'USE [?] EXEC sp_change_users_login report'
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'UserName',	'UserSID')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5,	C6)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', * FROM @Dataset32
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	/*INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 60 ,'AltenateDiv','</div>' )*/
	if(select Count(R1) from ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo and R3=32)=0 
	BEGIN
	    DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
		---Logice goes here	
					INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	END
    END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
	END -- @Parm
	
	-----97--
    IF(@Parm11_5=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Determine how many open connections exist to the specific database'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE  @Dataset31 TABLE([Database_name]  VARCHAR(MAX),	[No of Connections]  VARCHAR(MAX),	[Login Name]  VARCHAR(MAX))
	INSERT INTO @Dataset31
	EXEC sp_MSforeachdb 'USE [?] SELECT DB_NAME(dbid) AS [DB Name],	   COUNT(dbid) AS [Number Of Connections],	   loginame AS [Login Name]	FROM sys.sysprocesses	GROUP BY dbid, loginame'
	--huge data

	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'DB Name',	'Number Of Connections',	'Login Name', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'C',	'L', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @Dataset31
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')

	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7) 
		SELECT TOP(@AllowFewToReport)  R1, R2, S1, R3,  R4,	C5,	C6,	C7 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
		IF @HugeDataCounter>@AllowFewToReport
			BEGIN
				INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
				Set @ExcelSheetNo=@ExcelSheetNo+1;
			END
		ELSE
			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm


	-----98--
    IF(@Parm11_6=0)
    BEGIN
    BEGIN TRY
	SET @QHeadingTw0='Determine User/Role permission on each database.'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	SET @QTotalNo=@QTotalNo+1
	SET @QRunTime=CONVERT( VARCHAR, GETDATE(),20)
/*	if(@QTotalNo%2=0)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate1','<div class="Alternate1">' )
	ELSE
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0 ,'Altenate2', '<div class="Alternate2">' )*/
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0) 
 	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 20 ,'QNo',CAST(@QTotalNo AS VARCHAR) )
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 21,'ET',@QRunTime)
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
	DECLARE @User_Database as TABLE(DatabaseName varchar(max), [User/Role] varchar(max), Permission varchar(max))
	INSERT INTO @User_Database
	EXEC sp_MSforeachdb  'USE [?] --if(db_name() not in (''master'',''msdb'',''tempdb'',''distribution'',''DWQueue'' ,''DWDiagnostics'',''model'')) 
	SELECT  DB_NAME() AS [Database], p.name AS [User/Role], dperm.permission_name AS [Permission] FROM sys.database_permissions AS dperm INNER JOIN sys.database_principals AS p ON dperm.grantee_principal_id = p.principal_id WHERE p.type IN (''S'', ''U'', ''G'') /* S: SQL user, U: Windows user, G: Windows group*/ ORDER BY[Database], [User/Role], [Permission];'
	--huge data
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 31, 'H',	'DB Name',	'User/Role',	'Permissions', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7, ExclSheet) VALUES (	@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 231, 'S',	'L',	'C',	'L', @ExcelSheetNo)
	INSERT INTO ##DataForSheet(R1, R2, S1, R3, R4,	C5,	C6,	C7, ExclSheet)
	SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D', *, @ExcelSheetNo FROM @User_Database
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 131,'<DivObsUL>',  '<Div><UL  class=Observations>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 130, 'H',	'Observations')
	---Logic goes here	
	
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 139,'</DIVObsUL>','</UL></DIV>')
	INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
	SELECT @HugeDataCounter=count(*) FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3=32
	IF @HugeDataCounter>0 AND @HugeDataCounter<=@AllowFewToReport 
	BEGIN
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7) 
		SELECT R1, R2, S1, R3,  R4,	C5,	C6,	C7 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 IN  (31, 231)
		INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7) 
		SELECT TOP(@AllowFewToReport)  R1, R2, S1, R3,  R4,	C5,	C6,	C7 FROM ##DataForSheet WHERE  R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 =32
		IF @HugeDataCounter>@AllowFewToReport
			BEGIN
				INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5, C5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo, 'rr'
				Set @ExcelSheetNo=@ExcelSheetNo+1;
			END
		ELSE
			DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
	END
	ELSE IF @HugeDataCounter=0
	BEGIN
		DELETE FROM ##DataForSheet WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo 
		DELETE FROM ##RunOnce WHERE R1=@QSectionNo AND R2=@QSectionSubNo AND S1=@QSectionSplitNo AND R3 NOT IN(-1,0,10,20,21,60)
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 71, 'NODATA', 0
	END
	ELSE
	BEGIN
		INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 23, 'DATAONSHEET', @ExcelSheetNo
		Set @ExcelSheetNo=@ExcelSheetNo+1;
	END
	END TRY
    BEGIN CATCH
    INSERT INTO ##RunOnce (R1, R2, S1, R3, R4, R5) SELECT @QSectionNo, @QSectionSubNo, @QSectionSplitNo, 72, 'E', 'Error. '+CAST(ERROR_NUMBER() AS VARCHAR)+' Source. '+ERROR_MESSAGE() AS msg
    END CATCH
   END -- @Parm







END --END OF IF(@Database_login_users_roles_and_permissions=0)
--At the end some queries analysis
IF OBJECT_ID(N'tempdb..#queries') IS NOT NULL drop table #queries;
IF OBJECT_ID(N'tempdb..#queries3') IS NOT NULL drop table #queries3;
select cast(R1 as varchar)+'.'+cast(R2 as varchar) +'.'+cast(RR2 as varchar) AS Qid, C5 as [database],C39 as Query_text, C40 As [QueryPlan], C38 as PoorOnResources  INTO #queries 
	FROM ##DataForSheet where h2 like 'query%'
Update #queries set qid='<a href=#'+LEFT(qid, LEN(qid) - CHARINDEX('.',REVERSE (qid))) +'><span id='+qid+'>'+qid+'</span></a>'

;WITH UniqueQueriesQid AS ( 
		SELECT DISTINCT Qid, Query_text FROM #queries  
), 
AggregatedDataQid AS ( 
		SELECT Query_text, STRING_AGG(Qid, ',  ') AS QidList FROM UniqueQueriesQid GROUP BY Query_text 
), 
UniqueQueriesWithPOR AS ( 
		SELECT DISTINCT PoorOnResources, Query_text FROM #queries  
), 
AggregatedDataPOR AS ( 
		SELECT Query_text, STRING_AGG(PoorOnResources, ',  ') AS PORList FROM UniqueQueriesWithPOR GROUP BY Query_text 
)
SELECT cast('' as sysname) DbName, qid.QidList, por.query_text,cast('' as varchar(max)) query_plan, por.PORList into #Queries3 FROM AggregatedDataQid qid INNER JOIN AggregatedDataPOR por on por.query_text=qid.query_text
update #Queries3 set  DbName=[database], query_plan=QueryPlan  from #queries inner join #Queries3 on #Queries3.Query_text=#queries.Query_text
delete from #Queries3 where DbName  in ('master','msdb','tempdb','distribution','DWQueue' ,'DWDiagnostics','model') 

IF OBJECT_ID(N'tempdb..#queries') IS NOT NULL drop table #queries;
/*
select distinct r1, r2, rr2, C38, C39, C40 from ##DataForSheet where
r1=10
and r2 in (1, 2, 3)
and rr2 in (18, 19, 22)
*/
--update ##DataForSheet set C38='', C39='', C40=''
--QueriesForTuning
if(@Queries_and_Stored_Procedures=0)
BEGIN
	SET @QHeadingTw0='Queries for tuning and analysis'
	SET @QSectionSubNo=@QSectionSubNo+1
	SET @QSectionSplitNo=@QSectionSplitNo+1
	INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5,H2) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 10, 'H2', CAST(@QSectionNo AS VARCHAR)+'.'+CAST(@QSectionSubNo AS VARCHAR)+'. '+@QHeadingTw0,'Don not tablarize') 
	Declare @Params varchar(max), @NoOfQueries int, @DbName varchar(max), @Qid varchar(max), @QText varchar(max), @QPlan varchar(max);
	DECLARE QueryPoorOn CURSOR FOR
	select PORList, count(query_text) from #Queries3 group by PORList order by count(query_text) desc
	OPEN QueryPoorOn
	FETCH NEXT FROM QueryPoorOn INTO @Params, @NoOfQueries
	while(@@FETCH_STATUS=0)
	BEGIN
		SET @QSectionNo=@QSectionNo+1
		SET @QSectionSubNo=1
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',	'<span class=PoorQueryHeadings>' +@Params+'</span> <span> No of queries	'+cast(@NoOfQueries as varchar)+'</span>')
		SET @QSectionSubNo=@QSectionSubNo+1
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 30, '<TABLE>','<TABLE>')
		DECLARE QueryText CURSOR FOR
		select  DbName, Query_text,query_plan, QidList from #queries3 where PORList=@Params
		Open QueryText
		FETCH NEXT FROM QueryText INTO @DbName, @QText, @QPlan, @Qid
		while(@@FETCH_STATUS=0)
		BEGIN
				SET @QSectionSubNo=@QSectionSubNo+1
				INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',	'Database : '+@DbName+@Qid)
				if(len(@QPlan)>10)
				BEGIN
					SET @QSectionSubNo=@QSectionSubNo+1
					INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',	'<textarea>'+@QText+'</textarea>')
				END
--				if(len(@QPlan)>10)
	--			BEGIN
					--SET @QSectionSubNo=@QSectionSubNo+1
					--INSERT INTO ##RunOnce(R1, R2, S1, R3, R4,	C5) VALUES (@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 32, 'D',	'<textarea>'+@QPlan+'</textarea>')
		--		END
		FETCH NEXT FROM QueryText INTO @DbName, @QText, @QPlan, @Qid
		END	
		Close QueryText
		deallocate QueryText
		SET @QSectionSubNo=@QSectionSubNo+1
		INSERT INTO ##RunOnce(R1, R2, S1, R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 39, '</TABLE>', '</TABLE>')
		FETCH NEXT FROM QueryPoorOn INTO @Params, @NoOfQueries
	END
	Close QueryPoorOn
	DEALLOCATE QueryPoorOn
END

if(@includefooter=0)
begin
SET @QSectionNo=@QSectionNo+1
INSERT INTO ##RunOnce(R1,  R3, R4, R5) VALUES(@QSectionNo,1000, 'script', 
'<footer>
	<div style="background-color: #f2ebeb;padding:15px;align-items: center;text-align: center;">
	<img src="https://drive.google.com/uc?export=view&id=1HpKA67nBwB5iTLscsmCpk3cNkF61IqOG" width=915px />
  <span>
  <p><b>More help required about sql server scripts? </b></p>
  <a href="mailto:niazdawar@yahoo.com"><img height=40px style="height:40px; padding:10px" src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSou3ESdl3F7WbLOki-qTuP8w04Ku9Vyf9E-g&usqp=CAU" alt="Email for more assistance" /></a></span> 
  <span> <a href="https://wa.me/+923358443565"> <img height=40px style="height:40px; padding:10px" src="https://cdn-icons-png.flaticon.com/512/124/124034.png?w=360" alt="Contact on whatsapp" /></a></span>
	</div>
</footer>
')
end
SET @QSectionNo=@QSectionNo+1
INSERT INTO ##RunOnce(R1, R2, S1,R3, R4, R5) VALUES(@QSectionNo, @QSectionSubNo, @QSectionSplitNo, 0,'script', 
'<script>
function ShowHidTables(Divid,toggleElement) {    
var x = document.getElementById(Divid);    	if (x.style.display === "none") { x.style.display = "block";    } else { x.style.display = "none";    }    
var y = document.getElementById(toggleElement);  if (y.innerHTML === "show") { y.innerHTML = "hide";    } else { y.innerHTML = "show";    }  
}    
</script>
</body>
</html>')
-----------------------------updates

--UPDATE ##RunOnce SET TABLE_TAG='<div id="'+H1+'"><table border=1 cellspacing=0>' WHERE H2 !='OBSRV' AND TABLE_TAG LIKE '<TABLE%' --and R1='3'

INSERT INTO ##DataForSheet(R1, R2,S1, R3, R4, H2, ExclSheet) 
SELECT distinct ro.R1, ro.R2, ro.S1, ro.R3, ro.R4, ro.R5, ds.ExclSheet from ##RunOnce ro inner join ##DataForSheet ds on ro.r1=ds.r1 and ro.r2=ds.r2  where ro.R3=10 

UPDATE ##RunOnce SET TRTH_TAG='<THEAD><TR>', TD='<TH Style="Text-align:left;vertical-align:text-bottom;">' WHERE R3=31  
UPDATE ##RunOnce SET TRTH_TAG='<THEAD><TR>', TDR='<TH Style="Text-align:Right;vertical-align:text-bottom;">' WHERE R3=31  
UPDATE ##RunOnce SET TRTH_TAG='<THEAD><TR>', TDC='<TH Style="Text-align:Center">' WHERE R3=31  
UPDATE ##RunOnce SET TRTH_TAG='<TR>', TD='<TD Style="Text-align:left;vertical-align:text-bottom;">' WHERE R3=32  
UPDATE ##RunOnce SET TRTH_TAG='<TR>', TDR='<TD Style="Text-align:Right;vertical-align:text-bottom;">' WHERE R3=32  
UPDATE ##RunOnce SET TRTH_TAG='<TR>', TDC='<TD Style="Text-align:center;vertical-align:text-bottom;">' WHERE R3=32  



--UPDATE ##RunOnce SET R5='<span class=ShowHideButton id=btn'+cast(R1 as varchar) + '_' + cast(R2 as varchar)+' onclick="ShowHidTables('''+cast(R1 as varchar) + '_' + cast(R2 as varchar)+''',''btn'+cast(R1 as varchar) + '_' + cast(R2 as varchar)+''')">hide</span><div id="div'+cast(R1 as varchar) + '_' + cast(R2 as varchar)++'"><table border=1 cellspacing=0>' WHERE R3=30
UPDATE ##RunOnce SET R5='<div id="div'+cast(R1 as varchar) + '_' + cast(R2 as varchar)++'"><table border=1 cellspacing=0>' WHERE R3=30
UPDATE ##RunOnce SET R5='</table></div>' WHERE R3=39
UPDATE ##RunOnce SET R5='<P class="fields">'+R5+'</p>' WHERE R3=25
UPDATE ##RunOnce SET R5='<P class="error">'+R5+'</p>' WHERE R3=72
UPDATE ##RunOnce SET R5='<H1 class=headingOne>'+ R5+'</H1>' WHERE R3=-1
UPDATE ##RunOnce SET R5='<H2 class=headingTwo>'+ R5+'</H2>' WHERE R3=10
UPDATE ##RunOnce SET R5='<H3 class=headingThree>'+ R5+'</H3>' WHERE R3=11
UPDATE ##RunOnce SET R5='<span class=timeStamp> | Timestamp: '+ R5+' | <a href="#" >Top</a></span>' WHERE R3=21
UPDATE ##RunOnce SET R5='<span class=numbering>Query no. '+ R5+'</span>' WHERE R3=20
UPDATE ##RunOnce SET R5='<span class=data_on_sheet style="background-color:yellow;"> | Sheet. '+ R5+'</span>' WHERE R3=23 -- data on sheet
UPDATE ##RunOnce SET C5= '<span class=lessDataOnReport>'+CASE C5 WHEN 'rf' THEN ' {less data by fields only}'
								 WHEN 'rr' THEN ' {less data by rows only}'  
								 WHEN 'rfr' THEN ' {less data by both fields and rows}' ELSE C5 END +'</span>' WHERE R3=23  --less data in report
UPDATE ##RunOnce SET R5='<p class="no_data">No data retrieved/found</p>' WHERE R3=71 --No data found
UPDATE ##RunOnce SET C5='<li class="general_observations">'+C5+'</p>' WHERE R3=132 AND H2='0'  --Observations normal
UPDATE ##RunOnce SET C5='<li class="positive_observations">'+C5+'</p>' WHERE R3=132 AND H2='1'  --Observations positive
UPDATE ##RunOnce SET C5='<li class="negative_observations">'+C5+'</p>' WHERE R3=132 AND H2='2'  --Observations negative
--DELETE a FROM ##RunOnce a  INNER JOIN ##RunOnce b on a.R1=b.R1 and b.h2='OBSRV' and a.r3>7

UPDATE ##RunOnce SET C5=REPLACE(UPPER(LEFT(C5,1))+LOWER(SUBSTRING(C5,2,LEN(C5))),'_',' ') WHERE R3=31 AND C5 IS NOT NULL
UPDATE ##RunOnce SET C6=REPLACE(UPPER(LEFT(C6,1))+LOWER(SUBSTRING(C6,2,LEN(C6))),'_',' ') WHERE R3=31 AND C6 IS NOT NULL
UPDATE ##RunOnce SET C7=REPLACE(UPPER(LEFT(C7,1))+LOWER(SUBSTRING(C7,2,LEN(C7))),'_',' ') WHERE R3=31 AND C7 IS NOT NULL
UPDATE ##RunOnce SET C8=REPLACE(UPPER(LEFT(C8,1))+LOWER(SUBSTRING(C8,2,LEN(C8))),'_',' ') WHERE R3=31 AND C8 IS NOT NULL
UPDATE ##RunOnce SET C9=REPLACE(UPPER(LEFT(C9,1))+LOWER(SUBSTRING(C9,2,LEN(C9))),'_',' ') WHERE R3=31 AND C9 IS NOT NULL
UPDATE ##RunOnce SET C10=REPLACE(UPPER(LEFT(C10,1))+LOWER(SUBSTRING(C10,2,LEN(C10))),'_',' ') WHERE R3=31 AND C10 IS NOT NULL
UPDATE ##RunOnce SET C11=REPLACE(UPPER(LEFT(C11,1))+LOWER(SUBSTRING(C11,2,LEN(C11))),'_',' ') WHERE R3=31 AND C11 IS NOT NULL
UPDATE ##RunOnce SET C12=REPLACE(UPPER(LEFT(C12,1))+LOWER(SUBSTRING(C12,2,LEN(C12))),'_',' ') WHERE R3=31 AND C12 IS NOT NULL
UPDATE ##RunOnce SET C13=REPLACE(UPPER(LEFT(C13,1))+LOWER(SUBSTRING(C13,2,LEN(C13))),'_',' ') WHERE R3=31 AND C13 IS NOT NULL
UPDATE ##RunOnce SET C14=REPLACE(UPPER(LEFT(C14,1))+LOWER(SUBSTRING(C14,2,LEN(C14))),'_',' ') WHERE R3=31 AND C14 IS NOT NULL
UPDATE ##RunOnce SET C15=REPLACE(UPPER(LEFT(C15,1))+LOWER(SUBSTRING(C15,2,LEN(C15))),'_',' ') WHERE R3=31 AND C15 IS NOT NULL
UPDATE ##RunOnce SET C16=REPLACE(UPPER(LEFT(C16,1))+LOWER(SUBSTRING(C16,2,LEN(C16))),'_',' ') WHERE R3=31 AND C16 IS NOT NULL
UPDATE ##RunOnce SET C17=REPLACE(UPPER(LEFT(C17,1))+LOWER(SUBSTRING(C17,2,LEN(C17))),'_',' ') WHERE R3=31 AND C17 IS NOT NULL
UPDATE ##RunOnce SET C18=REPLACE(UPPER(LEFT(C18,1))+LOWER(SUBSTRING(C18,2,LEN(C18))),'_',' ') WHERE R3=31 AND C18 IS NOT NULL
UPDATE ##RunOnce SET C19=REPLACE(UPPER(LEFT(C19,1))+LOWER(SUBSTRING(C19,2,LEN(C19))),'_',' ') WHERE R3=31 AND C19 IS NOT NULL
UPDATE ##RunOnce SET C20=REPLACE(UPPER(LEFT(C20,1))+LOWER(SUBSTRING(C20,2,LEN(C20))),'_',' ') WHERE R3=31 AND C20 IS NOT NULL
UPDATE ##RunOnce SET C21=REPLACE(UPPER(LEFT(C21,1))+LOWER(SUBSTRING(C21,2,LEN(C21))),'_',' ') WHERE R3=31 AND C21 IS NOT NULL
UPDATE ##RunOnce SET C22=REPLACE(UPPER(LEFT(C22,1))+LOWER(SUBSTRING(C22,2,LEN(C22))),'_',' ') WHERE R3=31 AND C22 IS NOT NULL
UPDATE ##RunOnce SET C23=REPLACE(UPPER(LEFT(C23,1))+LOWER(SUBSTRING(C23,2,LEN(C23))),'_',' ') WHERE R3=31 AND C23 IS NOT NULL
UPDATE ##RunOnce SET C24=REPLACE(UPPER(LEFT(C24,1))+LOWER(SUBSTRING(C24,2,LEN(C24))),'_',' ') WHERE R3=31 AND C24 IS NOT NULL
UPDATE ##RunOnce SET C25=REPLACE(UPPER(LEFT(C25,1))+LOWER(SUBSTRING(C25,2,LEN(C25))),'_',' ') WHERE R3=31 AND C25 IS NOT NULL
UPDATE ##RunOnce SET C26=REPLACE(UPPER(LEFT(C26,1))+LOWER(SUBSTRING(C26,2,LEN(C26))),'_',' ') WHERE R3=31 AND C26 IS NOT NULL
UPDATE ##RunOnce SET C27=REPLACE(UPPER(LEFT(C27,1))+LOWER(SUBSTRING(C27,2,LEN(C27))),'_',' ') WHERE R3=31 AND C27 IS NOT NULL
UPDATE ##RunOnce SET C28=REPLACE(UPPER(LEFT(C28,1))+LOWER(SUBSTRING(C28,2,LEN(C28))),'_',' ') WHERE R3=31 AND C28 IS NOT NULL
UPDATE ##RunOnce SET C29=REPLACE(UPPER(LEFT(C29,1))+LOWER(SUBSTRING(C29,2,LEN(C29))),'_',' ') WHERE R3=31 AND C29 IS NOT NULL
UPDATE ##RunOnce SET C30=REPLACE(UPPER(LEFT(C30,1))+LOWER(SUBSTRING(C30,2,LEN(C30))),'_',' ') WHERE R3=31 AND C30 IS NOT NULL
UPDATE ##RunOnce SET C31=REPLACE(UPPER(LEFT(C31,1))+LOWER(SUBSTRING(C31,2,LEN(C31))),'_',' ') WHERE R3=31 AND C31 IS NOT NULL
UPDATE ##RunOnce SET C32=REPLACE(UPPER(LEFT(C32,1))+LOWER(SUBSTRING(C32,2,LEN(C32))),'_',' ') WHERE R3=31 AND C32 IS NOT NULL
UPDATE ##RunOnce SET C33=REPLACE(UPPER(LEFT(C33,1))+LOWER(SUBSTRING(C33,2,LEN(C33))),'_',' ') WHERE R3=31 AND C33 IS NOT NULL
UPDATE ##RunOnce SET C34=REPLACE(UPPER(LEFT(C34,1))+LOWER(SUBSTRING(C34,2,LEN(C34))),'_',' ') WHERE R3=31 AND C34 IS NOT NULL
UPDATE ##RunOnce SET C35=REPLACE(UPPER(LEFT(C35,1))+LOWER(SUBSTRING(C35,2,LEN(C35))),'_',' ') WHERE R3=31 AND C35 IS NOT NULL
UPDATE ##RunOnce SET C36=REPLACE(UPPER(LEFT(C36,1))+LOWER(SUBSTRING(C36,2,LEN(C36))),'_',' ') WHERE R3=31 AND C36 IS NOT NULL
UPDATE ##RunOnce SET C37=REPLACE(UPPER(LEFT(C37,1))+LOWER(SUBSTRING(C37,2,LEN(C37))),'_',' ') WHERE R3=31 AND C37 IS NOT NULL
UPDATE ##RunOnce SET C38=REPLACE(UPPER(LEFT(C38,1))+LOWER(SUBSTRING(C38,2,LEN(C38))),'_',' ') WHERE R3=31 AND C38 IS NOT NULL
UPDATE ##RunOnce SET C39=REPLACE(UPPER(LEFT(C39,1))+LOWER(SUBSTRING(C39,2,LEN(C39))),'_',' ') WHERE R3=31 AND C39 IS NOT NULL
UPDATE ##RunOnce SET C40=REPLACE(UPPER(LEFT(C40,1))+LOWER(SUBSTRING(C40,2,LEN(C40))),'_',' ') WHERE R3=31 AND C40 IS NOT NULL

--SELECT * FROM ##RunOnce where R4='H1'
update ##RunOnce SET C6=Replace(C6,'<', '< '),	C7=Replace(C7,'<', '< '),	C8=Replace(C8,'<', '< '),	C9=Replace(C9,'<', '< '),	C10=Replace(C10,'<', '< '),	C11=Replace(C11,'<', '< '),	C12=Replace(C12,'<', '< '),	C13=Replace(C13,'<', '< '),	C14=Replace(C14,'<', '< '),	C15=Replace(C15,'<', '< '),	C16=Replace(C16,'<', '< '),	C17=Replace(C17,'<', '< '),	C18=Replace(C18,'<', '< '),	C19=Replace(C19,'<', '< '),	C20=Replace(C20,'<', '< ') 
WHERE H2='HTML_TAG_IN_CODE' --- like Query Plan inside query output

--IF(SELECT COUNT(H1) FROM ##RunOnce WHERE H1=@Owner)=0 update ##RunOnce Set R5=Replace(R5,'<', '< ')
--SELECT R1, R2,S1, R3, Count(R1)FROM ##RunOnce
--GROUP BY R1, R2,S1,  R3
--HAVING COUNT(R1)>1
--SELECT * FROM ##RunOnce WHERE R1=4 AND R2=1 AND R3 IN (10,31, 231)
--

;WITH CTE AS(
SELECT ROW_NUMBER( ) OVER (PARTITION BY R1,R2, R3  ORDER BY R3)  AS QueryRecords, R1,   
	R2, S1, R3, QuerySort, R4,	R5, TRTH_TAG, TD, TDR, TDC, C5,	C6,	C7,	C8,	C9,	C10,	C11,	C12,	C13,	C14,	C15,	C16,	C17,	C18,	C19,	C20,	C21,	C22,	C23,	C24,	C25,	C26,	C27,	C28,	C29,	C30,	C31,	C32,	C33,	C34,	C35,	C36,	C37,	C38,	C39,	C40
FROM ##RunOnce 
), CTE2 AS(
SELECT 
	R1, R2, S1,  R3, QuerySort, R4,
	REPORT=ISNULL(R5,'')+ 
	ISNULL(TRTH_TAG,'')+ TD+ ISNULL(C5,'')+
	CASE WHEN (SELECT ISNULL(C6,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C6,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C6,'') +	CASE WHEN (SELECT ISNULL(C7,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C7,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C7,'') +	CASE WHEN (SELECT ISNULL(C8,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C8,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C8,'') +	CASE WHEN (SELECT ISNULL(C9,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C9,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C9,'') +	CASE WHEN (SELECT ISNULL(C10,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C10,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C10,'') +	CASE WHEN (SELECT ISNULL(C11,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C11,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C11,'') +	CASE WHEN (SELECT ISNULL(C12,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C12,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C12,'') +	CASE WHEN (SELECT ISNULL(C13,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C13,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C13,'') +	CASE WHEN (SELECT ISNULL(C14,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C14,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C14,'') +	CASE WHEN (SELECT ISNULL(C15,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C15,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C15,'') +	CASE WHEN (SELECT ISNULL(C16,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C16,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C16,'') +	CASE WHEN (SELECT ISNULL(C17,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C17,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C17,'') +	CASE WHEN (SELECT ISNULL(C18,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C18,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C18,'') +	CASE WHEN (SELECT ISNULL(C19,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C19,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C19,'') +	CASE WHEN (SELECT ISNULL(C20,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C20,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C20,'') +	CASE WHEN (SELECT ISNULL(C21,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C21,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C21,'') +	CASE WHEN (SELECT ISNULL(C22,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C22,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C22,'') +	CASE WHEN (SELECT ISNULL(C23,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C23,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C23,'') +	CASE WHEN (SELECT ISNULL(C24,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C24,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C24,'') +	CASE WHEN (SELECT ISNULL(C25,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C25,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C25,'') +	CASE WHEN (SELECT ISNULL(C26,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C26,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C26,'') +	CASE WHEN (SELECT ISNULL(C27,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C27,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C27,'') +	CASE WHEN (SELECT ISNULL(C28,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C28,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C28,'') +	CASE WHEN (SELECT ISNULL(C29,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C29,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C29,'') +	CASE WHEN (SELECT ISNULL(C30,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C30,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C30,'') +	CASE WHEN (SELECT ISNULL(C31,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C31,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C31,'') +	CASE WHEN (SELECT ISNULL(C32,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C32,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C32,'') +	CASE WHEN (SELECT ISNULL(C33,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C33,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C33,'') +	CASE WHEN (SELECT ISNULL(C34,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C34,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C34,'') +	CASE WHEN (SELECT ISNULL(C35,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C35,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C35,'') +	CASE WHEN (SELECT ISNULL(C36,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C36,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C36,'') +	CASE WHEN (SELECT ISNULL(C37,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C37,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C37,'') +	CASE WHEN (SELECT ISNULL(C38,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C38,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C38,'') +	CASE WHEN (SELECT ISNULL(C39,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C39,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C39,'') +	CASE WHEN (SELECT ISNULL(C40,'')  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT ISNULL(C40,'')   FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + ISNULL(C40,'') 
	--CASE WHEN (SELECT C6 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C6  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C6 + 	 CASE WHEN (SELECT C7 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C7  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C7 + 	 CASE WHEN (SELECT C8 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C8  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C8 + 	 CASE WHEN (SELECT C9 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C9  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C9 + 	 CASE WHEN (SELECT C10 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C10  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C10 + 	 CASE WHEN (SELECT C11 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C11  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C11 + 	 CASE WHEN (SELECT C12 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C12  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C12 + 	 CASE WHEN (SELECT C13 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C13  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C13 + 	 CASE WHEN (SELECT C14 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C14  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C14 + 	 CASE WHEN (SELECT C15 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C15  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C15 + 	 CASE WHEN (SELECT C16 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C16  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C16 + 	 CASE WHEN (SELECT C17 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C17  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C17 + 	 CASE WHEN (SELECT C18 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C18  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C18 + 	 CASE WHEN (SELECT C19 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C19  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C19 + 	 CASE WHEN (SELECT C20 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C20  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C20 + 	 CASE WHEN (SELECT C21 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C21  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C21 + 	 CASE WHEN (SELECT C22 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C22  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C22 + 	 CASE WHEN (SELECT C23 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C23  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C23 + 	 CASE WHEN (SELECT C24 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C24  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C24 + 	 CASE WHEN (SELECT C25 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C25  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C25 + 	 CASE WHEN (SELECT C26 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C26  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C26 + 	 CASE WHEN (SELECT C27 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C27  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C27 + 	 CASE WHEN (SELECT C28 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C28  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C28 + 	 CASE WHEN (SELECT C29 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C29  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C29 + 	 CASE WHEN (SELECT C30 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C30  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C30 + 	 CASE WHEN (SELECT C31 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C31  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C31 + 	 CASE WHEN (SELECT C32 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C32  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C32 + 	 CASE WHEN (SELECT C33 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C33  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C33 + 	 CASE WHEN (SELECT C34 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C34  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C34 + 	 CASE WHEN (SELECT C35 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C35  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C35 + 	 CASE WHEN (SELECT C36 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C36  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C36 + 	 CASE WHEN (SELECT C37 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C37  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C37 + 	 CASE WHEN (SELECT C38 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C38  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C38 + 	 CASE WHEN (SELECT C39 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C39  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C39 + 	 CASE WHEN (SELECT C40 FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =31 ) != '' THEN CASE  (SELECT  C40  FROM ##RunOnce WHERE R1= c.R1 AND R2=c.R2 AND S1=c.S1 AND R3 =231 ) WHEN 'R' THEN TDR WHEN 'C' THEN TDC ELSE TD END  ELSE '' END + C40  
FROM CTE c WHERE  R3 != 231 
) 
select 	R1, R2, S1,  R3, QuerySort, R4, REPORT=REPLACE(REPORT,'<','< ') from CTE2 ORDER BY R1, R2,S1, R3, QuerySort

GO
--INSERT INTO ##RunOnce(R1, R2, S1, R3,  R4,	C5,	C6,	C7,	C8,	C9)
	--VALUES (@QSectionNo,	@QSectionSubNo, @QSectionSplitNo,  231, 'S',	'LEFT', 'RIGHT',	'LEFT',	'RIGHT',	'LEFT')	

--SELECT * FROM ##DataForSheet WHERE R1=1 AND R2=4 ORDER BY R1, R2,S1, R3
IF(SELECT count(*) FROM ##DataForSheet)>0
BEGIN
	SELECT ExclSheet, R1, 	R2, 	R3, 	R4,'', H2,	C5,  	CASE WHEN LEN(C6)>30000 AND R1=10 THEN SUBSTRING(C6,0, 30000) ELSE C6 END AS C6, 	C7, 	C8, 	CASE WHEN LEN(C9)>30000 AND R1=10 THEN SUBSTRING(C9,0, 30000) ELSE C9 END AS C9, 	C10, 	C11, 	C12, 	C13, 	C14, 	C15, 	C16, 	CASE WHEN LEN(C17)>30000 AND R1=10 THEN SUBSTRING(C17,0, 30000) ELSE C17 END AS C17, 	C18, 	C19, 	C20, 	C21, 	C22, 	C23, 	C24, 	C25, 	C26, 	C27, 	C28, 	C29, 	C30, 	C31, 	C32, 	C33, 	C34, 	C35, 	C36, 	C37, 	C38, 	C39, 	C40, ExclSheet
	FROM ##DataForSheet WHERE R3 != 231  ORDER BY R1, R2, S1, R3 
END

