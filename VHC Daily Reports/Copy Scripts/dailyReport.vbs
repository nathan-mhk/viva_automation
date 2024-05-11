Const destinationFolder = "C:\Users\User\Downloads\DailyReport\"

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim folder
Set folder = fso.GetFolder(destinationFolder)
For Each file In folder.Files
    ' WScript.Echo "del " & file.Path
    fso.DeleteFile file.Path, True
Next

Dim productTypes
productTypes = Array("DEO", "DER", "DSTF", "JAR", "MEO")

Dim tgtDate
Dim tgtYear
Dim tgtMonth

tgtDate = Date - 2
tgtYear = Year(tgtDate)
tgtMonth = Right("00" & Month(tgtDate), 2)

' WScript.Echo tgtYear & "-" & tgtMonth

For Each productType In productTypes
    sourceFile = "Path\To\File\Finch " & productType & " Report\" & tgtYear & "\Finch " & productType & " Report " & tgtYear & "-" & tgtMonth & ".xlsm"
    ' WScript.Echo sourceFile
    fso.CopyFile sourceFile, destinationFolder
Next ' productType

set fso = Nothing
