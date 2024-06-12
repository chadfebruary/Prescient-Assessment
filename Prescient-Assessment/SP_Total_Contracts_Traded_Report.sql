CREATE PROCEDURE SP_Total_Contracts_Traded_Report
	@DateFrom DATE,
	@DateTo DATE
AS
BEGIN
	DECLARE @DailyTotals TABLE(FileDate DATE, Total FLOAT)
	INSERT INTO @DailyTotals
	SELECT
		FileDate,
		SUM(ContractsTraded)
	FROM dbo.DailyMtm
	WHERE FileDate between @DateFrom and @DateTo and ContractsTraded > 0
	GROUP BY filedate

	SELECT
		FORMAT(mtm.FileDate, 'MM/dd/yyyy') as [File Date],
		mtm.[Contract] as [Contract],
		SUM(mtm.ContractsTraded) as [Contracts Traded],
		CAST((SUM(mtm.ContractsTraded) / MAX(dt.Total)) * 100 AS DECIMAL(10,8)) AS [% Of Total Contracts Traded]
	FROM dbo.DailyMTM mtm
	INNER JOIN @DailyTotals dt ON dt.FileDate = mtm.FileDate
	WHERE mtm.FileDate between @DateFrom and @DateTo and mtm.ContractsTraded > 0
	GROUP BY mtm.FileDate, mtm.[Contract]
	ORDER BY mtm.FileDate, mtm.[Contract]
END