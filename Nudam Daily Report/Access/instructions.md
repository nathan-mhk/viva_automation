# NUDAM Access DB Instructions

0. Insert the desired dates and line numbers inside the table `CustomDatesLines`.

1. Run the macro `GenmanAdjTbl`. The table `ManAdjust` should shows up after the code had finished running.
    > Query `shotCycle` uses query `crosstabShiftCounter` && query `minCycleTime`. Which is used by the macro to create (overwrite everytime) table `ManAdjust`

2. User perform any necessary changes on the columns `Shift1`, `Shift2`, and `Shift3` of the table `ManAdjust`.
   - For correct values of ShiftCounters, leave them as is.
   - For cycle time, if the values in column `CycleTime` were correct, replace the value in column `NewCycle` with -1

3. Save the changes made on the table `ManAdjust`, then run the macro `ApplyManualAdjustment`. A message "Manual Adjustments Applied" should pops up after the changes were applied to the DB.
    > Query `UnpivotManAdj` unpivots the table `ManAdjust`, which is used by the query `customFullView`. Query `customFullView` was used to create the table `CustomFull`, which was when used by `updateFromCustomFull` to apply the changes