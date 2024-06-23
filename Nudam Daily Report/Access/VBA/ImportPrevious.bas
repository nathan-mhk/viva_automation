Option Explicit

'------------------------------------------------------------
' ImportPrevious
'
'------------------------------------------------------------
Sub ImportPrevious(wsName As String, tbName As String)

    DoCmd.SetWarnings False
    
    ' Reference: https://www.mrexcel.com/board/threads/open-and-close-excel-from-access-vba.662190/
    
    Dim fileNameFull As String
    Dim xlApp As Excel.Application
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    Dim table As Excel.ListObject
    
    fileNameFull = "C:\Users\Nudam\Documents\Manual Update.xlsm"

    Debug.Print ("Opening Excel file")
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.ScreenUpdating = False
    ' xlApp.Visible = True
    xlApp.Visible = False
    Set wb = xlApp.workbooks.Open(fileNameFull)
    Debug.Print ("Excel file opened")
    
    Set ws = wb.Worksheets(wsName)
    Set table = ws.ListObjects(tbName)
    
    Debug.Print ("Importing table " & tbName)
    
    Dim dateVal, lineVal, shiftVal, statusVal, shotVal, CycleVal, sqlStr As String
    Dim rowIndex As Integer
    
    ' Starts at 1 instead of 0, so 1 = header row; 2 = first data row
    For rowIndex = 2 To table.Range.Rows.Count
        dateVal = table.Range(rowIndex, 1).Value
        lineVal = table.Range(rowIndex, 2).Value2
        shiftVal = table.Range(rowIndex, 3).Value2
        statusVal = table.Range(rowIndex, 4).Value2
        shotVal = table.Range(rowIndex, 5).Value2
        CycleVal = table.Range(rowIndex, 6).Value2
        
        dateVal = Format(dateVal, "yyyy-mm-dd")
        
        If shotVal < 0 And CycleVal < 0 Then GoTo ContinueForLoop
        
        sqlStr = "UPDATE ProductionLineStatusHistory SET ProductionLineStatusHistory.LineStatus = """ & statusVal & """"
        
        If shotVal >= 0 Then
            sqlStr = sqlStr & ", ProductionLineStatusHistory.ShiftCounter = """ & shotVal & """"
        End If
        
        If CycleVal >= 0 Then
            sqlStr = sqlStr & ", ProductionLineStatusHistory.CycleTimeLast = """ & CycleVal & """"
        End If
        
        ' Date literal: https://stackoverflow.com/a/19810054
        sqlStr = sqlStr & " WHERE ProductionLineStatusHistory.ProductionDate=#" & dateVal & "# "
        sqlStr = sqlStr & "AND ProductionLineStatusHistory.ShiftNo = """ & shiftVal & """ "
        sqlStr = sqlStr & "AND ProductionLineStatusHistory.ProductionLineNo = """ & lineVal & """"

        Debug.Print (sqlStr)
        DoCmd.RunSQL sqlStr
        
ContinueForLoop:
    Next
    
    Debug.Print ("Table imported")
    
    Debug.Print ("Closing Excel")
    wb.Close (False)
    xlApp.Quit
    Debug.Print ("Excel closed")
    
    DoCmd.SetWarnings True

End Sub

Public Function ImportPreviousRun()
    ImportPrevious "Custom Full", "CustomFull"
    
    Dim dateVal, dateStr, sqlStr As String
    Dim diff As Integer
        
    ' Sunday = 1, Saturday = 7
    If Weekday(Date) = 2 Then
        ' Monday
        diff = -3
        dateStr = " >= "
    Else
        ' Tuesday - Friday
        diff = -1
        dateStr = " = "
    End If
    
    dateStr = dateStr & "#" & Format(Date + diff, "yyyy-mm-dd") & "#"
    
    sqlStr = "UPDATE ProductionLineStatusHistory SET ProductionLineStatusHistory.LineStatus = ""Running"" "
    sqlStr = sqlStr & "WHERE (((ProductionLineStatusHistory.ShiftCounter)>0) AND ((ProductionLineStatusHistory.ProductionDate)" & dateStr & "))"
    
    DoCmd.SetWarnings False
    
    Debug.Print (sqlStr)
    DoCmd.RunSQL sqlStr
    
    DoCmd.SetWarnings True
End Function
