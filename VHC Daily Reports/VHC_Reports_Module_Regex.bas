' Sources
' https://www.ablebits.com/office-addins-blog/regex-match-excel/#regexpmatch-function
' https://www.ablebits.com/office-addins-blog/regex-extract-strings-excel/#vba-function
' https://www.ablebits.com/office-addins-blog/excel-regex-replace/#function

Public Function RegExpMatch(input_range As Range, pattern As String, Optional match_case As Boolean = True) As Variant
  Dim arRes() As Variant 'array to store the results
  Dim iInputCurRow, iInputCurCol, cntInputRows, cntInputCols As Long 'index of the current row in the source range, index of the current column in the source range, count of rows, count of columns

  On Error GoTo ErrHandl

  RegExpMatch = arRes

  Set regex = CreateObject("VBScript.RegExp")
  regex.pattern = pattern
  regex.Global = True
  regex.MultiLine = True
  If True = match_case Then
    regex.ignorecase = False
  Else
    regex.ignorecase = True
  End If

  cntInputRows = input_range.Rows.Count
  cntInputCols = input_range.Columns.Count
  ReDim arRes(1 To cntInputRows, 1 To cntInputCols)

  For iInputCurRow = 1 To cntInputRows
    For iInputCurCol = 1 To cntInputCols
      arRes(iInputCurRow, iInputCurCol) = regex.Test(input_range.Cells(iInputCurRow, iInputCurCol).Value)
    Next
  Next

  RegExpMatch = arRes
  Exit Function
ErrHandl:
    RegExpMatch = CVErr(xlErrValue)
End Function

Public Function RegExpExtract(text As String, pattern As String, Optional instance_num As Integer = 0, Optional match_case As Boolean = True)
  Dim text_matches() As String
  Dim matches_index As Integer

  On Error GoTo ErrHandl
        
  RegExpExtract = ""
        
  Set regex = CreateObject("VBScript.RegExp")
  regex.pattern = pattern
  regex.Global = True
  regex.MultiLine = True
        
  If True = match_case Then
    regex.ignorecase = False
  Else
    regex.ignorecase = True
  End If
        
  Set matches = regex.Execute(text)
        
  If 0 < matches.Count Then
      If (0 = instance_num) Then
        ReDim text_matches(matches.Count - 1, 0)
        For matches_index = 0 To matches.Count - 1
          text_matches(matches_index, 0) = matches.Item(matches_index)
        Next matches_index
        RegExpExtract = text_matches
      Else
        RegExpExtract = matches.Item(instance_num - 1)
      End If
  End If
  Exit Function
        
ErrHandl:
    RegExpExtract = CVErr(xlErrValue)
End Function


Public Function RegExpReplace(text As String, pattern As String, text_replace As String, Optional instance_num As Integer = 0, Optional match_case As Boolean = True) As String
  Dim text_result, text_find As String
  Dim matches_index, pos_start As Integer

  On Error GoTo ErrHandle
  text_result = text
  Set regex = CreateObject("VBScript.RegExp")

  regex.pattern = pattern
  regex.Global = True
  regex.MultiLine = True

  If True = match_case Then
    regex.ignorecase = False
  Else
    regex.ignorecase = True
  End If

  Set matches = regex.Execute(text)

  If 0 < matches.Count Then
    If (0 = instance_num) Then
      text_result = regex.Replace(text, text_replace)
    Else
      If instance_num <= matches.Count Then
        pos_start = 1
        For matches_index = 0 To instance_num - 2
          pos_start = InStr(pos_start, text, matches.Item(matches_index), vbBinaryCompare) + Len(matches.Item(matches_index))
        Next matches_index

        text_find = matches.Item(instance_num - 1)
        text_result = Left(text, pos_start - 1) &amp; Replace(text, text_find, text_replace, pos_start, 1, vbBinaryCompare)
      End If
    End If
  End If

  RegExpReplace = text_result
  Exit Function

ErrHandle:
  RegExpReplace = CVErr(xlErrValue)
End Function


