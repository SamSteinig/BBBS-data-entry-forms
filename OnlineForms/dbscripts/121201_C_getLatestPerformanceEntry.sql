--Put this at the top of your SQL file
USE BBBSA
GO

/* Procedure to generate drop down list date range */
IF NOT OBJECT_ID('sp_getLatestPerformanceEntry') is NULL
DROP PROC sp_getLatestPerformanceEntry
GO
PRINT 'Procedure sp_getLatestPerformanceEntry dropped'
GO

CREATE PROC sp_getLatestPerformanceEntry (
		@vcAgencyID		VARCHAR(4)
	)
AS
BEGIN
	
	DECLARE @MaxYear INTEGER
	SELECT @MaxYear=Max(Year) FROM tbl_frmPerformance WHERE AgencyID=@vcAgencyID

	DECLARE @MaxMonth INTEGER
	SELECT @MaxMonth=Max(Month) FROM tbl_frmPerformance WHERE AgencyID=@vcAgencyID AND Year=@MaxYear

	DECLARE @MaxDate VARCHAR(30)
	SELECT @MaxDate = CAST(@MaxYear AS varchar(4)) + '-' + CAST(@MaxMonth AS varchar(2)) + '-01'

	SELECT PerformanceID, AgencyID, Year, Month ,@MaxDate as LatestDate,
	CONVERT(datetime, CONVERT(VARCHAR, '2001-01-01', 101)) AS StartDate,
	DATEADD(month, -12, CONVERT(datetime, CONVERT(VARCHAR, @MaxDate, 101))) AS yearStartMaxDate,
	DATEADD(month, -12, CONVERT(datetime, GetDate())) AS yearStartTodayDate,
	DATEADD(month, 2, CONVERT(datetime, CONVERT(VARCHAR, @MaxDate, 101))) AS StopDate
	FROM tbl_frmPerformance WHERE AgencyID=@vcAgencyID AND Year=@MaxYear AND Month=@MaxMonth

END

GO

GRANT EXECUTE ON sp_getLatestPerformanceEntry TO B3SAWeb
GO
PRINT 'Procedure sp_getLatestPerformanceEntry created'
GO
