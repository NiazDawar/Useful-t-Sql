/*--===========Inactive Session Cleanup Script======================================================
Owner:				Niaz Dawar
Created On:			2024-11-14 22:46:25 -- SELECT GETDATE()
Last Modified On:	2024-11-14 22:46:25
Purpose:			The script is designed to identify and kill user sessions that have been in a 'sleeping' state 
					for a specified duration of inactivity (more than @Greater_Then_MinutesSinceLastActivity minutes), 
					with no active transactions. It is particularly useful for cleaning up orphaned or idle sessions that might 
					be consuming system resources unnecessarily. 
					The script also temp logs the details of any killed sessions or errors in a table for audit purposes.

					This script is useful for database administrators who need to identify and clean up idle sessions in SQL Server 
					that may be affecting server performance by holding on to resources unnecessarily.
*/
Declare @Greater_Then_MinutesSinceLastActivity int=1 --sleeping session where last_request_end_time old by 30 minutes
DECLARE @user_spid INT, @MinutesSinceLastActivity int;
DECLARE @SessionKillLog AS TABLE(
    LogID INT IDENTITY(1,1) PRIMARY KEY,
    SPID INT,
    KillTime DATETIME,
    Comment NVARCHAR(255),
	is_killed bit
)
WHILE (SELECT COUNT(session_id)
       FROM sys.dm_exec_sessions
       WHERE status = N'sleeping'
         AND open_transaction_count = 0
         AND is_user_process = 1
         AND DATEDIFF(MINUTE, last_request_end_time, GETDATE()) >= @Greater_Then_MinutesSinceLastActivity
         AND session_id <> @@SPID) > 0
BEGIN
    BEGIN TRY
        SELECT top 1 @user_spid=session_id, @MinutesSinceLastActivity=DATEDIFF(MINUTE, last_request_end_time, GETDATE())
        FROM sys.dm_exec_sessions
        WHERE status = N'sleeping'
          AND open_transaction_count = 0
          AND is_user_process = 1
          AND DATEDIFF(MINUTE, last_request_end_time, GETDATE()) >= @Greater_Then_MinutesSinceLastActivity
          AND session_id <> @@SPID;
            BEGIN TRY
                PRINT 'Killing session: ' + CONVERT(VARCHAR, @user_spid);
                EXEC('KILL ' + @user_spid);
                INSERT INTO @SessionKillLog (SPID, KillTime, Comment, is_killed)
                VALUES (@user_spid, GETDATE(), 'Session killed due to inactivite from last '+CONVERT(varchar, @MinutesSinceLastActivity)+' minutes', 1);
            END TRY
            BEGIN CATCH
                INSERT INTO @SessionKillLog (SPID, KillTime, Comment, is_killed)
                VALUES (@user_spid, GETDATE(), 'Error occurred while killing session: ' + ERROR_MESSAGE(), 0);
            END CATCH;
    END TRY
    BEGIN CATCH
        PRINT 'Error occurred in the main block: ' + ERROR_MESSAGE();
END CATCH;
END

IF EXISTS (SELECT * FROM @SessionKillLog) 
BEGIN
IF EXISTS (SELECT * FROM @SessionKillLog where is_killed=1) 
    SELECT SPID, KillTime, Comment, 'Yes' as [Is Session Killed?] FROM @SessionKillLog where is_killed=1;
IF EXISTS (SELECT * FROM @SessionKillLog where is_killed=0) 
    SELECT SPID, KillTime, Comment, 'No' as [Is Session Killed?] FROM @SessionKillLog where is_killed=0;
END
ELSE
	PRINT 'No sessions matching the specified criteria were found for termination.'
GO