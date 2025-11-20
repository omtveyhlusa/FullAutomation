--SELECT Fab.[ECN No_], Fab.[Part No_], Fab.[Work Center], CONVERT(decimal(10,5),Fab.[Run Time]) AS 'Run Time (Min)', FabC.[Work Center] AS 'Calculated Work Center', CONVERT(Decimal(10,5),Fabc.[Run Time]) AS 'Calculated Run Time (min)', ABS(CONVERT(Decimal(10,5),Fab.[Run Time]) - CONVERT(Decimal(10,5),Fabc.[Run Time])) AS 'Difference'
--FROM [ECN ME Fabrication Calculated] AS FabC INNER JOIN
--[ECN ME Fabrication] AS Fab ON FabC.[ECN No_] = Fab.[ECN No_] AND Fabc.[Part No_] = Fab.[Part No_] AND Fabc.[Operation No_] = Fab.[Operation No_]
--WHERE Fab.[Work Center] NOT LIKE '091%' AND CONVERT(decimal(10,5),Fab.[Run Time]) <> CONVERT(Decimal(10,5),Fabc.[Run Time])
--ORDER BY [ECN No_]

DECLARE @PN Varchar(50)
SET @PN = '%'

SELECT [ECN No_], [Part No_], [Work Center], [Run Time (Min)] * 60 AS '[Run Time (Sec)]', [Calculated Work Center], [Calculated Run Time (min)] * 60 AS '[Calculated Run Time (Sec)]', ABS(100*(Difference / [Calculated Run Time (min)])) AS '% Different'
FROM(SELECT Fab.[ECN No_], Fab.[Part No_], Fab.[Work Center], CONVERT(decimal(10,5),Fab.[Run Time]) AS 'Run Time (Min)', FabC.[Work Center] AS 'Calculated Work Center', CONVERT(Decimal(10,5),Fabc.[Run Time]) AS 'Calculated Run Time (min)', ABS(CONVERT(Decimal(10,5),Fab.[Run Time]) - CONVERT(Decimal(10,5),Fabc.[Run Time])) AS 'Difference'
FROM [ECN ME Fabrication Calculated] AS FabC INNER JOIN
[ECN ME Fabrication] AS Fab ON FabC.[ECN No_] = Fab.[ECN No_] AND Fabc.[Part No_] = Fab.[Part No_] AND Fabc.[Operation No_] = Fab.[Operation No_]
WHERE Fab.[Work Center] NOT LIKE '091%' AND FabC.[Run Time] <> 0 AND CONVERT(decimal(10,5),Fab.[Run Time]) <> CONVERT(Decimal(10,5),Fabc.[Run Time])
GROUP BY Fab.[ECN No_], Fab.[Part No_], Fab.[Work Center], CONVERT(decimal(10,5),Fab.[Run Time]), FabC.[Work Center], CONVERT(Decimal(10,5),Fabc.[Run Time]), ABS(CONVERT(Decimal(10,5),Fab.[Run Time]) - CONVERT(Decimal(10,5),Fabc.[Run Time])))
AS OGtable
WHERE ABS(100*(Difference / [Calculated Run Time (min)])) > 5 AND [Work Center] NOT LIKE '%30%' AND [Work Center] NOT LIKE '20%' AND [Work Center] NOT LIKE '40%' AND [Part No_] LIKE @PN
--WHERE [Work Center] LIKE '10%'
ORDER BY [ECN No_] DESC
