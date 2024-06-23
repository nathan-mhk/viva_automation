Option Explicit

Function CurrentShift() As Integer
    CurrentShift = Int(Hour(Now) / 8)
    CurrentShift = IIf(CurrentShift = 0, 3, CurrentShift)
End Function

Public Function InsertMissingRecordsRun()

    Dim tgtDate As String
    Dim tgtShift As String
    Dim strSQL As String
    Dim shiftNo As Integer
    
    tgtDate = ""
    tgtShift = 0
    
    Do
        tgtDate = InputBox("Enter a missing date (YYYY-MM-DD). Empty to exit", "Missing Date", tgtDate)
        If tgtDate = "" Then Exit Do
        
        tgtShift = tgtShift Mod 3 + 1
        tgtShift = InputBox("Enter missing shift (1/2/3). 0 to exit", "Missing Shift", tgtShift)
        If tgtShift = 0 Then Exit Do
        
        strSQL = "UPDATE ProductionLineStatus SET ProductionDate = #" & tgtDate & "#, ShiftNo = ""Shift" & tgtShift & """"
        Debug.Print strSQL
        DoCmd.SetWarnings False
        DoCmd.RunSQL strSQL
        
        strSQL = "INSERT INTO ProductionLineStatusHistory SELECT * FROM ProductionLineStatus"
        Debug.Print strSQL
        DoCmd.SetWarnings False
        DoCmd.RunSQL strSQL
    Loop
    
    shiftNo = CurrentShift
    strSQL = "UPDATE ProductionLineStatus SET ProductionDate = #" & IIf(shiftNo = 3, Date - 1, Date) & "#, ShiftNo = ""Shift" & shiftNo & """"
    Debug.Print strSQL
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL

End Function
