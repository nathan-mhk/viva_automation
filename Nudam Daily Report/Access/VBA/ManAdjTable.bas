Option Explicit

Public Function genManAdjTbl()
    Dim strSQL As String
    
    strSQL = "DELETE * FROM ManAdjust"
    Debug.Print strSQL
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    
    DoCmd.OpenQuery ("shotCycle")
    
    strSQL = "INSERT INTO ManAdjust SELECT * FROM shotCycle"
    Debug.Print strSQL
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
End Function
