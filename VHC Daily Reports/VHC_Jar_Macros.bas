Option Explicit

Const PROD_SHEET_NAME As String = "Production"
Const PROD_TABLE_NAME As String = "Prod"
Const PROD_NUM_ROWS As Variant = 7

Const HTL_SHEET_NAME As String = "HTL"
Const HTL_TABLE_NAME As String = "HTL"
Const HTL_NUM_ROWS As Variant = 6

Function ModifyDates(table As ListObject, numRows As Variant) As Range

    Dim numTblRows As Variant
    Dim firstRow As Variant

    Dim sourceRange As Range
    Dim destinRange As Range

    numTblRows = table.Range.Rows.Count
    firstRow = numTblRows - numRows

    Set sourceRange = table.Range(firstRow, 1)
    Set destinRange = sourceRange.Offset(1, 0)
    sourceRange.AutoFill Destination:=Range(sourceRange, destinRange)

    Set sourceRange = destinRange
    Set destinRange = table.Range(numTblRows, 1)
    sourceRange.AutoFill Destination:=Range(sourceRange, destinRange), Type:=xlFillCopy

    Set ModifyDates = sourceRange

End Function

Sub CopyDown(table As ListObject, numRows As Variant)

    Application.ScreenUpdating = False

    Dim firstRowToCopy As Variant
    Dim oldBtmRowNum As Variant
    
    oldBtmRowNum = table.Range.Rows.Count
    firstRowToCopy = oldBtmRowNum - numRows + 1

    table.Range.Rows(firstRowToCopy).Resize(numRows + 1).EntireRow.Hidden = False

    ' Copy the rows
    Dim copyRange As Range
    Set copyRange = table.Range.Rows(firstRowToCopy).Resize(numRows)
    copyRange.Copy
    
    ' ListRows(-1) to keep it in bounds; .Offset(1) to insert the row below
    table.ListRows(oldBtmRowNum - 1).Range.Offset(1).Insert
    Application.CutCopyMode = False

    Dim firstRangeToHide As Range
    Set firstRangeToHide = ModifyDates(table, numRows)
    firstRangeToHide.Resize(numRows).EntireRow.Hidden = True

    ' Modify the dates
    ' Dim sourceCell As Range
    ' Dim destinCell As Range

    ' Set sourceCell = table.Range(oldBtmRowNum, 1)
    ' Set destinCell = sourceCell.Offset(1, 0)
    ' sourceCell.AutoFill Destination:=Range(sourceCell, destinCell)

    ' Set sourceCell = destinCell
    ' Set destinCell = table.Range(table.Range.Rows.Count, 1)
    ' sourceCell.AutoFill Destination:=Range(sourceCell, destinCell), Type:=xlFillCopy

    ' ' ActiveSheet.Rows(sourceCell.Row & ":" & destinCell.Row).EntireRow.Hidden = True
    ' sourceCell.Resize(numRows).EntireRow.Hidden = True
    
    Application.ScreenUpdating = True

End Sub

Sub UpdateFormatting(tblName As String)
    Application.ScreenUpdating = False
    
    Dim table As ListObject
    Dim rng As Range
    
    Set table = ActiveSheet.ListObjects(tblName)
    
    Set rng = table.Range
    
    rng.FormatConditions.Delete
    
    rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=INDIRECT(""" & tblName & "[@Date]"")=MAX(INDIRECT(""" & tblName & "[Date]""))-1").Interior.Color = RGB(255, 255, 0)
    
    Application.ScreenUpdating = True
End Sub

Sub Cpy(wsName As String, tblName As String, numRows As Variant)

    Sheets(wsName).Select
    CopyDown ActiveSheet.ListObjects(tblName), numRows

End Sub

Sub ModifyDatesAndHide(tblName As String, numRows As Variant)

    Application.ScreenUpdating = False

    Dim firstRangeToHide As Range
    Set firstRangeToHide = ModifyDates(ActiveSheet.ListObjects(tblName), numRows)
    firstRangeToHide.Resize(numRows).EntireRow.Hidden = True
    
    Application.ScreenUpdating = True
    
End Sub

Sub ProdCpy()
'
' ProdCpy Macro
'
' Keyboard Shortcut: Ctrl+Shift+X
'
    Cpy PROD_SHEET_NAME, PROD_TABLE_NAME, PROD_NUM_ROWS
    UpdateFormatting PROD_TABLE_NAME
    
End Sub

Sub ProdModifyDatesAndHide()
'
' ProdModifyDatesAndHide Macro
'
' Keyboard Shortcut: Ctrl+Shift+S
'
    ModifyDatesAndHide PROD_TABLE_NAME, PROD_NUM_ROWS
    
End Sub

Sub HTLCpy()
'
' HTLCpy Macro
'
' Keyboard Shortcut: Ctrl+Shift+C

    Cpy HTL_SHEET_NAME, HTL_TABLE_NAME, HTL_NUM_ROWS
    UpdateFormatting HTL_TABLE_NAME
    
End Sub

Sub HTLModifyDatesAndHide()
'
' HTLModifyDatesAndHide Macro
'
' Keyboard Shortcut: Ctrl+Shift+D

    ModifyDatesAndHide HTL_TABLE_NAME, HTL_NUM_ROWS
    
End Sub

Sub ExpandCollapseDate(ptName As String, cbName As String)

    Application.ScreenUpdating = False
    
    Dim expand As Boolean
    expand = ActiveSheet.CheckBoxes(cbName).Value > 0
    
    ActiveSheet.PivotTables(ptName).PivotFields("Date").ShowDetail = expand
    
    Application.ScreenUpdating = True
    
End Sub

Sub Grp(ptName As String, byDay As Boolean)

    Application.ScreenUpdating = False
    
    On Error Resume Next

    Dim dateCell As Range
    Dim pt As PivotTable
    
    Set pt = ActiveSheet.PivotTables(ptName)
    
    pt.RowRange.Cells(2, 1).Select
    
    If byDay Then
        Selection.Ungroup
    Else
        ' Group by 7 days instead as there's no group by week option
        Selection.Group Start:=True, End:=True, By:=7, Periods:=Array(False, _
            False, False, True, False, False, False)
    End If
    
    Application.ScreenUpdating = False
    
End Sub

Sub ProdCB()
    ' ProductionCheckbox
    ExpandCollapseDate "ProdDayWk", "ProdExpand"
End Sub

Sub ProdOpt(byDay As Boolean)
    ' Production Option (Daily, Weekly)
    Grp "ProdDayWk", byDay
    
    ' By default, rows will be expanded after ungrouping.
    ' Need to fold rows depends on the state of the checkbox
    ProdCB
End Sub

Sub HTLCB()
    ' HTLCheckbox
    ExpandCollapseDate "HTLDayWk", "HTLExpand"
End Sub

Sub HTLOpt(byDay As Boolean)
    ' HTL Option (Daily, Weekly)
    Grp "HTLDayWk", byDay
    
    ' By default, rows will be expanded after ungrouping.
    ' Need to fold rows depends on the state of the checkbox
    HTLCB
End Sub

Sub RefreshAllClpsDate()
'
' RefreshAllClpsDate Macro
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    Application.ScreenUpdating = False
    
    ActiveWorkbook.RefreshAll
    
    ' ExpandCollapseDate uses ActiveSheet, so need to manually select the corresponding sheets
    Sheets("Production Summary").Select
    ProdCB
    
    Sheets("HTL Summary").Select
    HTLCB
    
    Application.ScreenUpdating = True
    
End Sub

Sub RefilterDate(pvTable As PivotTable, pvFilterType As XlPivotFilterType)
    
    Dim pvField As PivotField
    
    Set pvField = pvTable.PivotFields("Date")
    pvField.ClearAllFilters
    pvField.PivotFilters.Add Type:=pvFilterType, Value1:=Format(Date - 1, "yyyy-mm-dd")
    
End Sub

Sub RefilterMold(pvTable As PivotTable, moldName As String)

    Dim pvField As PivotField
    
    Set pvField = pvTable.PivotFields("MoldNo")
    pvField.ClearAllFilters
    pvField.PivotFilters.Add Type:=xlCaptionContains, Value1:=moldName
    
End Sub

Sub RefilterGraph()

    ' Keyboard Shortchut: Ctrl+Shift+G
    Application.ScreenUpdating = False
    
    Dim pvTable As PivotTable
    Dim pvField As PivotField
    
    Sheets("Production Summary").Select
    Set pvTable = ActiveSheet.PivotTables("FiftyPT")
    pvTable.RefreshTable
    RefilterDate pvTable, xlSpecificDate
    RefilterMold pvTable, "50"
    
    Sheets("HTL Summary").Select
    Set pvTable = ActiveSheet.PivotTables("SevenPT")
    pvTable.RefreshTable
    RefilterDate pvTable, xlSpecificDate
    RefilterMold pvTable, "7"
    
    Set pvTable = ActiveSheet.PivotTables("FourteenPT")
    pvTable.RefreshTable
    RefilterDate pvTable, xlSpecificDate
    RefilterMold pvTable, "14"
    
    Sheets("Graph Summary").Select
    Set pvTable = ActiveSheet.PivotTables("FnlAssemSum")
    pvTable.RefreshTable
    RefilterDate pvTable, xlBeforeOrEqualTo
    
    Application.ScreenUpdating = True
    
End Sub
