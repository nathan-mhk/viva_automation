Option Explicit

Public Function manualAdjustment()
    Dim strSQL As String
    
    strSQL = "DELETE * FROM CustomFull"
    Debug.Print strSQL
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    
    DoCmd.OpenQuery ("customFullView")
    
    strSQL = "INSERT INTO CustomFull SELECT * FROM customFullView"
    Debug.Print strSQL
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    
    DoCmd.OpenQuery ("updateFromCustomFull")
    DoCmd.OpenQuery ("fixDownRunning")
    MsgBox ("Manual Adjustments Applied")
End Function

