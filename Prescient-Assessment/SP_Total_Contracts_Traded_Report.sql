create proc SP_Total_Contracts_Traded_Report
	@DateFrom date,
	@DateTo date
as
begin
	select
		FileDate,
		[Contract],
		ContractsTraded,
		'' as [%OfTotalContractsTraded]
	from dbo.DailyMTM
	where filedate between @DateFrom and @DateTo
end