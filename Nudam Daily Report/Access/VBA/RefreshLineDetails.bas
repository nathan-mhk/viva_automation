Option Explicit

'------------------------------------------------------------
' RefreshLineDetails
'
'------------------------------------------------------------
Public Function RefreshLineDetails()

    DoCmd.SetWarnings False
    
    ' Reference: https://www.mrexcel.com/board/threads/open-and-close-excel-from-access-vba.662190/
    
    Dim fileNameFull As String
    Dim xlApp As Excel.Application
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    Dim table As Excel.ListObject
    
    fileNameFull = "C:\Users\Nudam\Documents\Finch Plant Daily Production Summary.xlsx"

    Debug.Print ("Opening Excel file")
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.ScreenUpdating = False
    ' xlApp.Visible = True
    xlApp.Visible = False
    Set wb = xlApp.workbooks.Open(fileNameFull)
    Debug.Print ("Excel file opened")
    
    Set ws = wb.Worksheets("LineDetails")
    Set table = ws.ListObjects("LineDetailsParsed")
    
    Debug.Print ("Importing table")
    Dim lineVal, moldVal, colorVal, siloVal, typeVal, CycleVal, rmkVal, sqlStr As String
    Dim rowIndex, lineIndex, moldIndex, colorIndex, siloIndex, typeIndex, rmkIndex As Integer
    
    lineIndex = 1
    moldIndex = 2
    colorIndex = 3
    siloIndex = 4
    rmkIndex = 6
    typeIndex = 7
    
    For rowIndex = 2 To table.Range.Rows.Count
        lineVal = table.Range(rowIndex, lineIndex).Value2
        moldVal = table.Range(rowIndex, moldIndex).Value2
        colorVal = table.Range(rowIndex, colorIndex).Value2
        siloVal = table.Range(rowIndex, siloIndex).Value2
        rmkVal = table.Range(rowIndex, rmkIndex).Value2
        typeVal = table.Range(rowIndex, typeIndex).Value2
        
        ' Escape double quotes
        rmkVal = Replace(rmkVal, Chr(34), """""")
        
        sqlStr = "UPDATE ProductionLineStatus SET ProductionLineStatus.MoldNo = """ & moldVal & """, "
        sqlStr = sqlStr & "ProductionLineStatus.ColorCode = """ & colorVal & """, "
        sqlStr = sqlStr & "ProductionLineStatus.SiloNo = """ & siloVal & """, "
        sqlStr = sqlStr & "ProductionLineStatus.StatusRemarks = """ & rmkVal & """, "
        sqlStr = sqlStr & "ProductionLineStatus.MoldType = """ & typeVal & """ "
        sqlStr = sqlStr & "WHERE ProductionLineStatus.ProductionLineNo=""" & lineVal & """"

        Debug.Print (sqlStr)
        DoCmd.RunSQL sqlStr
    Next
    Debug.Print ("Table imported")
    
    Debug.Print ("Closing Excel")
    wb.Close (False)
    xlApp.Quit
    Debug.Print ("Excel closed")
    
    DoCmd.SetWarnings True

End Function
