Option Explicit
Sub CutInsert(srcStr As String, dstStr As String)
    Range(srcStr).Select
    Selection.Cut
    Range(dstStr).Select
    Selection.Insert Shift:=xlToRight   ' There's no xlToLeft
End Sub

Sub MoveCols()
Attribute MoveCols.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim colNames() As Variant
    colNames = Array("S1Cycle", "S1Hrs", "S1Shot", "S1Cavity", _
    "S2Cycle", "S2Hrs", "S2Shot", "S2Cavity", _
    "S3Cycle", "S3Hrs", "S3Shot", "S3Cavity", _
    "Remarks", "S1Mold", "S2Mold", "S3Mold", "S1Print", "S2Print", "S3Print", _
    "HrStrtCls", "HrMaint", "HrSample")
    
    Application.ScreenUpdating = False
    
    Dim colName As Variant
    For Each colName In colNames
        CutInsert "Prod[[#All],[" & colName & "]]", "Prod[[#All],[Status]]"
    Next colName
    
    Application.ScreenUpdating = True
End Sub

Sub ResetCols()
    Dim dstColNames As Variant
    dstColNames = Array("S3Print", "S1CycleNotZero")
    
    Dim srcColNamesArr As Variant
    srcColNamesArr = Array(Array("Status", _
    "S1Cycle", "S1Hrs", "S1Shot", "S1Cavity", "S1Qty", "S1Dft", _
    "S2Cycle", "S2Hrs", "S2Shot", "S2Cavity", "S2Qty", "S2Dft", _
    "S3Cycle", "S3Hrs", "S3Shot", "S3Cavity", "S3Qty", "S3Dft", _
    "Total", "TotalDft", "DftRte", "Remarks", _
    "S1Mold", "S2Mold", "S3Mold", "S1Print", "S2Print"), _
    Array("HrStrtCls", "HrMaint", "HrSample"))
    
    Application.ScreenUpdating = False
    
    Dim i As Integer
    Dim dstColName As Variant, srcColNames As Variant, srcColName As Variant
    For i = LBound(dstColNames) To UBound(dstColNames)
        dstColName = dstColNames(i)
        srcColNames = srcColNamesArr(i)
        For Each srcColName In srcColNames
            CutInsert "Prod[[#All],[" & srcColName & "]]", "Prod[[#All],[" & dstColName & "]]"
        Next srcColName
    Next i
    
    Application.ScreenUpdating = True
End Sub

Sub Ctrl_M()
Attribute Ctrl_M.VB_ProcData.VB_Invoke_Func = "m\n14"
    Debug.Print ActiveWorkbook.Name & "::MoveCols"
    Application.Run ("'" & ActiveWorkbook.Name & "'!MoveCols")
End Sub

Sub Ctrl_Shift_M()
Attribute Ctrl_Shift_M.VB_ProcData.VB_Invoke_Func = "M\n14"
    Debug.Print ActiveWorkbook.Name & "::ResetCols"
    Application.Run ("'" & ActiveWorkbook.Name & "'!ResetCols")
End Sub
