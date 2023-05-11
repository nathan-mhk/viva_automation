Sub ProdCpy()
'
' ProdCpy Macro
'
' Keyboard Shortcut: Ctrl+Shift+X
'
    Application.ScreenUpdating = False
    
    ' Copy all columns
    ActiveCell.Range("A1:AS6").Select
    Selection.Copy
    
    ' Paste all columns
    ActiveCell.Offset(6, 0).Range("A1").Select
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
    
    ' AutoFill dates by +1 day
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A2"), Type:=xlFillDefault
    
    ' AutoFill dates by copying
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A6"), Type:=xlFillCopy
    
    ' The first data cell of yst
    ActiveCell.Offset(-6, 7).Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub

Sub HTLCpy()
'
' HTLCpy Macro
'
' Keyboard Shortcut: Ctrl+Shift+C
'
    Application.ScreenUpdating = False
    
    ' Copy all columns
    ' Change here~~~~~~~~~~v to increase #rows
    ActiveCell.Range("A1:BI5").Select
    Selection.Copy
    
    ' Paste all columns
    ' Change here~~~~~v to increase #rows
    ActiveCell.Offset(5, 0).Range("A1").Select
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
    
    ' AutoFill dates by +1 day
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A2"), Type:=xlFillDefault
    
    ' AutoFill dates by copying
    ActiveCell.Offset(1, 0).Range("A1").Select
    ' Change here~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~v to increase #rows
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A5"), Type:=xlFillCopy
    
    ' The first data cell of yst
    ' Change here~~~~~~v to increase #rows
    ActiveCell.Offset(-5, 6).Range("A1").Select
    
    Application.ScreenUpdating = True
    
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
    pvField.PivotFilters.Add Type:=pvFilterType, Value1:=Format(Date - 2, "yyyy-mm-dd")
    
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


