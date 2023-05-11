Option Explicit

Const PCS_PER_TRAY_CELL As String = "K5"
Const TRAYS_PER_PALLET_CELL As String = "L5"
Const STR_FORMAT As String = "#,##0"

Const LABEL_HEIGHT As Integer = 7
Const LABEL_RANGE_WIDTH As Integer = 3
Const LABEL_RANGE_COL_L As Integer = 1
Const LABEL_RANGE_COL_R As Integer = 9

Const L_LABEL_COL As Integer = 1
Const R_LABEL_COL As Integer = 6

Const PLT_ROW_OFFSET As Integer = 3
Const PLT_COL_OFFSET As Integer = 1
Const STR_ROW_OFFSET As Integer = 4
Const STR_COL_OFFSET As Integer = 0

Function GetFirstRowNum(pageNum As Variant) As Variant
    GetFirstRowNum =  LABEL_HEIGHT * (pageNum - 1) + 1
End Function

Sub SetPrintArea(pageNum As Variant)

    Dim TLRow As Variant, TLCol As Variant
    Dim BRRow As Variant, BRCol As Variant
    Dim printArea As Range

    ' TLRow = GetFirstRowNum(pageNum)
    ' TLCol = LABEL_RANGE_COL_L
    ' BRRow = TLRow + LABEL_HEIGHT - 1
    ' BRCol = LABEL_RANGE_COL_R

    ' Set printArea = Range( _
    '     Cells(TLRow, TLCol), _
    '     Cells(BRRow, BRCol) _
    ' )

    BRRow = GetFirstRowNum(pageNum) + LABEL_HEIGHT - 1
    BRCol = LABEL_RANGE_COL_R
    Set printArea = Range( _
        Cells(1, 1), _
        Cells(BRRow, BRCol) _
    )

    ActiveSheet.PageSetup.PrintArea = printArea.Address

End Sub

Sub RemovePrevLabels()

    Dim lastRow As Variant
    lastRow = Cells.SpecialCells(xlCellTypeLastCell).Row

    If lastRow > LABEL_HEIGHT Then
        Rows(LABEL_HEIGHT + 1 & ":" & lastRow).Delete
        SetPrintArea 1
    End if

End Sub

Sub ToggleRLabel(pageNum As Variant, hidden As Boolean)

    Dim lrLabelRangeTLRow As Variant, lrLabelRangeBRRow As Variant
    Dim rLabelRangeTLCol As Variant, rLabelRangeBRCol As Variant
    
    Dim rLabelRange As Range

    lrLabelRangeTLRow = GetFirstRowNum(pageNum)
    lrLabelRangeBRRow = lrLabelRangeTLRow + LABEL_HEIGHT - 1

    rLabelRangeTLCol = R_LABEL_COL
    rLabelRangeBRCol = rLabelRangeTLCol + LABEL_RANGE_WIDTH

    Set rLabelRange = Range( _
        Cells(lrLabelRangeTLRow, rLabelRangeTLCol), _
        Cells(lrLabelRangeBRRow, rLabelRangeBRCol) _
    )

    If hidden = True Then
        rLabelRange.Font.Color = vbWhite
        rLabelRange.Borders.ColorIndex = xlNone
    Else
        Dim lLabelRangeTLCol As Variant, lLabelRangeBRCol As Variant
    
        Dim lLabelRange As Range

        lLabelRangeTLCol = L_LABEL_COL
        lLabelRangeBRCol = lLabelRangeTLCol + LABEL_RANGE_WIDTH

        Set lLabelRange = Range( _
            Cells(lrLabelRangeTLRow, lLabelRangeTLCol), _
            Cells(lrLabelRangeBRRow, lLabelRangeBRCol) _
        )

        ' Copy the format of lLabel to rLabel
        lLabelRange.Copy
        rLabelRange.PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If

End Sub

Sub UpdateLabelValues(pageNum As Variant, pltNum As Variant, labelString As Variant, firstColNum As Variant)
    
    Dim firstRowNum As Variant
    Dim pltRowNum As Variant, pltColNum As Variant
    Dim strRowNum As Variant, strColNum As Variant
    
    firstRowNum = GetFirstRowNum(pageNum)
    
    pltRowNum = firstRowNum + PLT_ROW_OFFSET
    pltColNum = firstColNum + PLT_COL_OFFSET

    strRowNum = firstRowNum + STR_ROW_OFFSET
    strColNum = firstColNum + STR_COL_OFFSET
    
    ' PltNo
    Cells(pltRowNum, pltColNum).Value = pltNum
    
    ' labelString (countString)
    Cells(strRowNum, strColNum).Value = labelString
    
End Sub

Sub UpdateLLabel(pageNum As Variant, lPltNum As Variant, lCountString As Variant)
    ' Wrapper for UpdateLabelValues()
    UpdateLabelValues pageNum, lPltNum, lCountString, L_LABEL_COL
End Sub

Sub UpdateRLabel(pageNum As Variant, rPltNum As Variant, rCountString As Variant)
    ' Wrapper for UpdateLabelValues()
    UpdateLabelValues pageNum, rPltNum, rCountString, R_LABEL_COL
End Sub

Sub UpdateLabels(pageNum As Variant, numLabels As Variant, lPltNum As Variant, rPltNum As Variant, countString As String, endCountString As String)

    Dim lCountString As string, rCountString As String

    If numLabels = 1 Then
        lCountString = endCountString
    Else
        lCountString = countString
    End If

    If rPltNum > numLabels Then
        ' Current page has dangling labels (lLabel)
        UpdateLLabel pageNum, lPltNum, lCountString
        ToggleRLabel pageNum, True
        
        Debug.Print "[" & lPltNum & "]"
        Debug.Print lCountString
    Else
        ' Current page has lLabel && rLabel
        If rPltNum < numLabels Then
            ' Intermediate rLabels
            rCountString = countString
        Else
            ' Last rLabel (rPltNum == numLabels)
            rCountString = endCountString
        End If
    
        UpdateLLabel pageNum, lPltNum, lCountString
        UpdateRLabel pageNum, rPltNum, rCountString
        ToggleRLabel pageNum, False
        
        Debug.Print "[" & lPltNum & "]                                       [" & rPltNum & "]"
        Debug.Print lCountString & "        " & rCountString
    End If

End Sub

Sub CreateNextPage(pageNum As Variant)

    Dim currFirstRowNum As Variant, currLastRowNum As Variant, nextFirstRowNum As Variant
    Dim pageBreakRowNum As Variant

    currFirstRowNum = GetFirstRowNum(pageNum)
    nextFirstRowNum = GetFirstRowNum(pageNum + 1)
    currLastRowNum = nextFirstRowNum - 1

    Rows(currFirstRowNum & ":" & currLastRowNum).Copy Destination:=Rows(nextFirstRowNum)

    SetPrintArea (pageNum + 1)

    pageBreakRowNum = GetFirstRowNum(pageNum + 2)
    'Rows(pageBreakRowNum).PageBreak = xlPageBreakManual

End Sub

Sub GeneratePages(numPages As Variant, numLabels As Variant, countString As String, endCountString As String)

    Debug.Print Format(numLabels, STR_FORMAT) & " label(s), " & Format(numPages, STR_FORMAT) & " page(s)."
    
    Dim lPltNum As Variant, rPltNum As Variant
    lPltNum = 1
    rPltNum = numPages + 1
    
    Dim pageNum As Variant
    For pageNum = 1 To numPages
        
        ' Update current page
        UpdateLabels pageNum, numLabels, lPltNum, rPltNum, countString, endCountString
        
        ' Skip everything if current page is last page
        If pageNum >= numPages Then
            Exit For
        End If
        
        ' Create next page
        CreateNextPage pageNum
        
        ' Increment left and right label number
        lPltNum = lPltNum + 1
        rPltNum = rPltNum + 1
    Next pageNum
    
    Debug.Print "END"

End Sub

Sub GeneratePalletLabels()

    Application.ScreenUpdating = False

    RemovePrevLabels

    Dim numPallets As Variant, numTrays As Variant
    
    Dim pcsPerTray As Variant, traysPerPallet As Variant, pcsPerPallet As Variant
    
    Dim numPages As Variant, numLabels As Variant
    
    Dim endTrays As Variant, endPcs As Variant
    
    Dim countString As String, endCountString As String
    
    'Get number of pallets and number of trays from user
    numPallets = Application.InputBox("Please input the number of pallets", "Number of Pallets", Type:=1)
    numTrays = Application.InputBox("Please input the number of trays", "Number of Trays", Type:=1)
    
    'Get pcsPerTray and traysPerPallet info from cells K5 and L5
    pcsPerTray = ActiveSheet.Range(PCS_PER_TRAY_CELL).Value
    traysPerPallet = ActiveSheet.Range(TRAYS_PER_PALLET_CELL).Value
    pcsPerPallet = pcsPerTray * traysPerPallet
    
    ' Calculate number of labels to be generated/printed and the pcs for the last pallet.
    If numTrays > 0 Then
        ' Partial pallet
        numLabels = numPallets + 1
        endTrays = numTrays
        endPcs = pcsPerTray * numTrays
    Else
        ' Full pallet
        numLabels = numPallets
        endTrays = traysPerPallet
        endPcs = pcsPerPallet
    End If
    
    numPages = Application.WorksheetFunction.Ceiling(numLabels / 2, 1)

    ' The count string for full pallets
    countString = Format(pcsPerTray, STR_FORMAT) & " pcs x " & Format(traysPerPallet, STR_FORMAT) & " trays = " & Format(pcsPerPallet, STR_FORMAT) & " pcs"
    
    ' The count string for the last pallet
    endCountString = Format(pcsPerTray, STR_FORMAT) & " pcs x " & Format(endTrays, STR_FORMAT) & " trays = " & Format(endPcs, STR_FORMAT) & " pcs"
    
    Debug.Print "Total " & Format(numPallets, STR_FORMAT) & " pallets " & Format(numTrays, STR_FORMAT) & " trays."
    GeneratePages numPages, numLabels, countString, endCountString

    Application.ScreenUpdating = True

End Sub
