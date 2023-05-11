Sub ProdCpy()
'
' ProdCpy Macro
'
' Keyboard Shortcut: Ctrl+Shift+X
'
    Application.ScreenUpdating = False

    ' Copy all columns
    ActiveCell.Range("A1:AY14").Select
    Selection.Copy
    
    ' Paste all columns
    ActiveCell.Offset(14, 0).Range("A1").Select
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
    
    ' AutoFill dates by +1 day
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A2"), Type:=xlFillDefault
    
    ' AutoFill dates by copying
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A14"), Type:=xlFillCopy

    ' The first data cell of yst
    ActiveCell.Offset(-14, 3).Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub

Sub AssCpy()
'
' AssCpy Macro
'
' Keyboard Shortcut: Ctrl+Shift+C
'
    Application.ScreenUpdating = False
    
    ' Copy all columns
    ActiveCell.Range("A1:J9").Select
    Selection.Copy
    
    ' Paste all columnsa
    ActiveCell.Offset(9, 0).Range("A1").Select
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
    
    ' AutoFill dates by +1 day
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A2"), Type:=xlFillDefault
    
    ' AutoFill dates by copying
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A9"), Type:=xlFillCopy
    
    ' The first data cell of yst
    ActiveCell.Offset(-9, 2).Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub

Sub RefreshAll()
'
' RefreshAll Macro
' Keyboard Shortcut: Ctrl+Shift+R
    Application.ScreenUpdating = False
    ActiveWorkbook.RefreshAll
    Application.ScreenUpdating = True
    
End Sub

Sub RefilterDate(pvTable As PivotTable, pvFilterType As XlPivotFilterType)
    
    Dim pvField As PivotField
    
    Set pvField = pvTable.PivotFields("Date")
    pvField.ClearAllFilters
    pvField.PivotFilters.Add Type:=pvFilterType, Value1:=Format(Date - 2, "yyyy-mm-dd")
    
End Sub

Sub RefilterGraph()
'
' Keyboard Shortchut: Ctrl+Shift+G
    Application.ScreenUpdating = False
    
    Dim pvTable As PivotTable
    
    Sheets("Graph Summary").Select
    Set pvTable = ActiveSheet.PivotTables("FnlAssemSum")
    pvTable.RefreshTable
    RefilterDate pvTable, xlBeforeOrEqualTo
    
    Application.ScreenUpdating = True
    
End Sub
