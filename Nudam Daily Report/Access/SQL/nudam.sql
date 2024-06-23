-- crosstabShiftCounter
TRANSFORM First(ShiftCounter)
SELECT history.ProductionDate, history.ProductionLineNo
FROM ProductionLineStatusHistory AS history
INNER JOIN CustomDatesLines AS lines
ON (history.ProductionDate = lines.ProductionDate) AND (history.ProductionLineNo=lines.ProductionLineNo)
GROUP BY history.ProductionDate, history.ProductionLineNo
PIVOT ShiftNo IN ("Shift1", "Shift2", "Shift3");

-- customFullView
SELECT old.ProductionDate, old.ProductionLineNo, old.ShiftNo,
IIF(new.ShiftCounter IS NULL, 0, new.ShiftCounter) AS ShiftCounter,
IIF(new.NewCycle IS NULL, 0, IIF(new.NewCycle = -1, old.CycleTimeLast, new.NewCycle)) AS CycleTimeLast
FROM ProductionLineStatusHistory AS old
LEFT JOIN unpivotManAdj AS new
ON (old.ShiftNo = new.ShiftNo) AND (old.ProductionLineNo = new.ProductionLineNo) AND (old.ProductionDate = new.ProductionDate)
WHERE old.ProductionDate IN
	(SELECT DISTINCT ProductionDate FROM ManAdjust);

-- minCycleTime
SELECT history.ProductionDate, history.ProductionLineNo, Min(CycleTimeLast) AS CycleTime
FROM ProductionLineStatusHistory AS history
INNER JOIN CustomDatesLines AS lines
ON (history.ProductionLineNo=lines.ProductionLineNo) AND (history.ProductionDate = lines.ProductionDate)
GROUP BY history.ProductionDate, history.ProductionLineNo;

-- shotCycle
SELECT pvTbl.ProductionDate, pvTbl.ProductionLineNo, Shift1, Shift2, Shift3, minCycle.CycleTime AS CycleTime, mCycle.CycleTime AS ManCycle,
IIF(mCycle.CycleTime < minCycle.CycleTime, mCycle.CycleTime, -1) AS NewCycle
FROM (crosstabShiftCounter AS pvTbl
	INNER JOIN minCycleTime AS minCycle
	ON (pvTbl.ProductionDate = minCycle.ProductionDate) AND (pvTbl.ProductionLineNo = minCycle.ProductionLineNo)
	)
INNER JOIN CycleTime AS mCycle ON pvTbl.ProductionLineNo = mCycle.ProductionLineNo;

-- unpivotManAdj
SELECT ProductionDate, ProductionLineNo, 'Shift1' AS ShiftNo, Shift1 AS ShiftCounter, NewCycle
FROM ManAdjust

UNION ALL

SELECT ProductionDate, ProductionLineNo, 'Shift2' AS ShiftNo, Shift2 AS ShiftCounter, NewCycle
FROM ManAdjust

UNION ALL

SELECT ProductionDate, ProductionLineNo, 'Shift3' AS ShiftNo, Shift3 AS ShiftCounter, NewCycle
FROM ManAdjust;


-- fixDownRunning
UPDATE ProductionLineStatusHistory
SET LineStatus = IIF(ShiftCounter <= 0, "Down", "Running")
WHERE ProductionDate > DateSerial(Year(Date()) - 1, 1, 1);


-- updateFromCustomFull
UPDATE ProductionLineStatusHistory AS old
INNER JOIN customFull AS new
ON (old.ShiftNo = new.ShiftNo) AND (old.ProductionLineNo = new.ProductionLineNo) AND (old.ProductionDate = new.ProductionDate)
SET old.ShiftCounter = new.ShiftCounter, old.CycleTimeLast = new.CycleTimeLast;

