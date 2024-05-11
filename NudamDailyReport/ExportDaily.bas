Option Explicit

Public Const superFinch = "Path\To\File\supervisor-Finch.xls"
Public Const dailySummary = "Finch Plant Daily Production Summary.xlsx"
Public Const divLine = vbNewLine & "--------------------------------------------------------------------------------------------------------------------------------------------------------"
Public nextTimer As Double

Sub ImportLineDetails()
    
    Dim sourceWB As Workbook, _
        destinWB As Workbook
    
    Dim sourceWS As Worksheet, _
        destinWS As Worksheet
    
    ' Reference: https://stackoverflow.com/a/3389577
    Set sourceWB = Workbooks.Open(fileName:=superFinch, ReadOnly:=True)
    Set destinWB = Excel.Workbooks(dailySummary)
    
    Set sourceWS = sourceWB.Worksheets("Line Status")
    Set destinWS = destinWB.Worksheets("LineDetailsRaw")
    
    Debug.Print (Now & "    Removing existing table")
    destinWS.Range("A1:J80").Delete
    
    Debug.Print (Now & "    Copying designated range")
    sourceWS.Range("A3:J77").Copy
    destinWS.Range("A3").PasteSpecial (xlPasteValues)
    Application.CutCopyMode = False
    
    sourceWB.Close (False)
    
    Debug.Print (Now & "    Copy completed")
    Debug.Print (Now & "    Creating table from the copied range")
    
    ' Create a table from range
    ' Reference: https://stackoverflow.com/a/36874483
    destinWS.Activate
    destinWS.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("A3:J77"), XlListObjectHasHeaders:=xlNo).Name = "LineDetailsRaw"
    Debug.Print (Now & "    Table created")
    
    
    With destinWB
        Debug.Print (Now & "    Refreshing LineDetailsParsed")
        .connections("Query - LineDetailsParsed").Refresh
        Debug.Print (Now & "    LineDetailsParsed refreshed")
        
        Debug.Print (Now & "    Refreshing LineDetails")
        .connections("Query - LineDetails").Refresh
        Debug.Print (Now & "    LineDetails refreshed")
        
        .Save
        Debug.Print (Now & "    File saved")
    End With
    
End Sub

Sub RefreshQueries()
    
    With Excel.Workbooks(dailySummary)
        Dim queries As Variant
        queries = Array( _
                            "Present", _
                            "Day", _
                            "Shift1", _
                            "Shift2", _
                            "Shift3", _
                            "DailySummary", _
                            "ProductionSummary" _
                        )
        
        Dim i As Variant
        For Each i In queries
            Debug.Print (Now & "    Refreshing " & i)
            .connections("Query - " & i).Refresh
            Debug.Print (Now & "    " & i & " refreshed")
        Next i
        
        ' Make sure the dates are correct
        ' For some reasons the value could be #REF!
        For i = 1 To 3
            With .Worksheets("Shift" & i)
                .Activate
                .Range("L1").Value = "=$A$4"
            End With
        Next i
        
        With .Worksheets("Daily Summary")
            .Activate
            
            ' Manually re-insert the formulae lmao
            ' Shift 1
            .Range("J7").Value = "=[@S1Shot]*[@TotalCavity]"                ' Qty
            .Range("K7").Value = "=[@S1Shot]*[@CycleTime]/28800"            ' U-Rate
            .Range("J7:K7").AutoFill Destination:=.Range("J7:K82"), Type:=xlFillValues
            
            ' Shift 2
            .Range("M7").Value = "=[@S2Shot]*[@TotalCavity]"
            .Range("N7").Value = "=[@S2Shot]*[@CycleTime]/28800"
            .Range("M7:N7").AutoFill Destination:=.Range("M7:N82"), Type:=xlFillValues
            
            ' Shift 3
            .Range("P7").Value = "=[@S3Shot]*[@TotalCavity]"
            .Range("Q7").Value = "=[@S3Shot]*[@CycleTime]/28800"
            .Range("P7:Q7").AutoFill Destination:=.Range("P7:Q82"), Type:=xlFillValues
            
            .Range("S7").Value = "=SUM([@S1Qty],[@S2Qty],[@S3Qty])"         ' Total
            .Range("T7").Value = "=AVERAGE([@S1Urt],[@S2Urt],[@S3Urt])"     ' U-Rate
            .Range("S7:T7").AutoFill Destination:=.Range("S7:T82"), Type:=xlFillValues

        End With
        
        .Save
        Debug.Print (Now & "    File saved")
    End With

End Sub

Sub ExportDailyCopy()
    
    ' Strings for file names and paths
    Dim rootPath As String, _
        currentDate As String, _
        folderPath As String, _
        fileNameStr As String, _
        fileNameFull As String, _
        ocFolderPath As String, _
        ocFileName As String, _
        ocFileNameFull As String
    
    ' Workbook objects for referencing
    Dim sourceWB As Workbook, _
        destinWB As Workbook
    
    ' `For Each` loop iterator
    Dim wrksht As Worksheet
    Dim connection As WorkbookConnection
    
    rootPath = "Path\To\File"
    currentDate = Format(Date - 1, "yyyy-mm-dd")
    
    ' Path\To\File\yyyy Prod Report\yyyy-mm-dd Production Line Status - Daily Printed.xlsx
    folderPath = rootPath & "\" & Format(currentDate, "yyyy") & " Prod Report" & "\"
    ocFolderPath = folderPath & "OriginalCopy" & "\"

    ocFileName = currentDate & ".xlsx"
    fileNameStr = currentDate & " Production Line Status - Daily Printed.xlsx"
    
    fileNameFull = folderPath & fileNameStr
    ocFileNameFull = ocFolderPath & ocFileName
    
    ' Need to include `vbDirectory` for the case of empty directory
    ' Reference: https://stackoverflow.com/a/15482073
    ' Reference: https://stackoverflow.com/a/43661302
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
        ' `""""` prints `"`
        Debug.Print (Now & "    mkdir " & """" & folderPath & """")

        MkDir ocFolderPath
        Debug.Print (Now & "    mkdir " & """" & ocFolderPath & """")
    Else
        Debug.Print (Now & "    Directory found: " & """" & folderPath & """")
    End If
    
    ' Need to manually open the workbooks for them to show up in the Project List
    
    Debug.Print (Now & "    Saving daily copies to the directory")
    Set sourceWB = Excel.Workbooks(dailySummary)
    sourceWB.SaveCopyAs (fileNameFull)
    sourceWB.SaveCopyAs (ocFileNameFull)

    Debug.Print (Now & "    Daily copies has been saved")
    Application.Wait (Now + TimeValue("00:00:10"))
    Debug.Print (Now & "    Opening daily copies")
    
    Dim cpyArr As Variant
    Dim cpyFName As Variant
    Dim isFirst As Boolean      ' For static copies only, NOT daily copies
    
    cpyArr = Array(fileNameFull, ocFileNameFull)
    isFirst = True
    
    For Each cpyFName In cpyArr
        Set destinWB = Workbooks.Open(fileName:=cpyFName, ReadOnly:=False)
        
        If isFirst Then
            Debug.Print (Now & "    Converting to static values")
            ' Reference: https://www.extendoffice.com/documents/excel/4140-excel-save-workbook-as-values.html
            For Each wrksht In destinWB.Worksheets
                    wrksht.Cells.Copy
                    wrksht.Cells.PasteSpecial (xlPasteValues)
                Next
                Application.CutCopyMode = False
            Debug.Print (Now & "    Conversion completed")
        End If
        
        With destinWB
            With .Worksheets("Daily Summary")
                .Activate
                .Rows(6).Hidden = True
                .Columns("I").Hidden = True
                .Columns("L").Hidden = True
                .Columns("O").Hidden = True
                .Columns("R").Hidden = True
                
                ' Make the date static
                .Range("M1").Copy
                .Range("M1").PasteSpecial (xlPasteValues)
                Application.CutCopyMode = False
            End With
            
            If isFirst Then
                With .Worksheets("Present Shift")
                    .Activate
                End With

                Debug.Print (Now & "    Deleting connections")
                On Error Resume Next
                For Each connection In .connections
                    'Debug.Print (connection)
                    connection.Delete
                Next
                Debug.Print (Now & "    Connections deleted")
            End If
    
            .Close (True)
        End With
        
        isFirst = False
    Next cpyFName
    
    Debug.Print (Now & "    Daily copies saved and closed")

End Sub

Sub Run()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Debug.Print (Now & "    Run started" & divLine)
    
    Debug.Print (Now & "    Running ImportLineDetails " & divLine)
    ImportLineDetails
    Debug.Print (divLine & vbNewLine & Now & "    ImportLineDetails completed" & divLine)

    Debug.Print (Now & "    Running RefreshQueries " & divLine)
    RefreshQueries
    Debug.Print (divLine & vbNewLine & Now & "    RefreshQueries completed" & divLine)
    
    Debug.Print (Now & "    Running ExportDailyCopy " & divLine)
    ExportDailyCopy
    Debug.Print (divLine & vbNewLine & Now & "    ExportDailyCopy completed" & divLine)
    
    Debug.Print (Now & "    Run completed")
    
    Workbooks("automation.xlsm").Activate
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Call Timer

End Sub

Sub Timer()

    '                            HH:MM:SS
    nextTimer = Now + TimeValue("01:00:00")
    Application.OnTime EarliestTime:=nextTimer, Procedure:="ExportDaily.Run"
    Debug.Print (Now & "    Next scheduled run: " & Format(nextTimer, "yyyy-mm-dd HH:MM:SS") & " (" & nextTimer & ")" & divLine)

End Sub

Sub StopTimer()

    On Error Resume Next
    Application.OnTime EarliestTime:=nextTimer, Procedure:="ExportDaily.Run", Schedule:=False
    
    Debug.Print ("Timer stopped: " & Format(nextTimer, "yyyy-mm-dd HH:MM:SS") & " (" & nextTimer & ")" & divLine)
    MsgBox ("Timer scheduled at " & Format(nextTimer, "HH:MM:SS") & " on " & Format(nextTimer, "yyyy-mm-dd") & " has been stopped")
       
End Sub

Sub ManualStop()
    
    Dim timeVal As Double
    timeVal = 44728.6259606481
    
    On Error Resume Next
    Application.OnTime EarliestTime:=timeVal, Procedure:="ExportDaily.Run", Schedule:=False
    
    MsgBox ("Timer scheduled at " & timeVal & " stopped manually")
End Sub
