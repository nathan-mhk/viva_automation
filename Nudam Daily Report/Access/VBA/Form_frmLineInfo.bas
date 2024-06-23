Option Compare Database
Option Explicit

Dim BlnAutoRun As Boolean
Dim blnrun As Boolean
Dim strGreenSign As String
Dim strYellowSign As String
Dim strStopSign As String
Dim BlnReadStatus As Boolean
Dim chars As String
Dim strIdentifier As String
Dim strCommChar As String * 1
Dim strStatus As String

Dim lngStatus As Long

'-------------------------------------------------------------------------------
' Public Constants
'-------------------------------------------------------------------------------

' Output Control Lines (CommSetLine)
Const LINE_BREAK = 1
Const LINE_DTR = 2
Const LINE_RTS = 3

' Input Control Lines  (CommGetLine)
Const LINE_CTS = &H10&
Const LINE_DSR = &H20&
Const LINE_RING = &H40&
Const LINE_RLSD = &H80&
Const LINE_CD = &H80&

Const intPortID = 2
'-------------------------------------------------------------------------------
' System Constants
'-------------------------------------------------------------------------------
Private Const ERROR_IO_INCOMPLETE = 996&
Private Const ERROR_IO_PENDING = 997
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const OPEN_EXISTING = 3

' COMM Functions
Private Const MS_CTS_ON = &H10&
Private Const MS_DSR_ON = &H20&
Private Const MS_RING_ON = &H40&
Private Const MS_RLSD_ON = &H80&
Private Const PURGE_RXABORT = &H2
Private Const PURGE_RXCLEAR = &H8
Private Const PURGE_TXABORT = &H1
Private Const PURGE_TXCLEAR = &H4

' COMM Escape Functions
Private Const CLRBREAK = 9
Private Const CLRDTR = 6
Private Const CLRRTS = 4
Private Const SETBREAK = 8
Private Const SETDTR = 5
Private Const SETRTS = 3

'-------------------------------------------------------------------------------
' System Structures
'-------------------------------------------------------------------------------
Private Type COMSTAT
    fBitFields As Long ' See Comment in Win32API.Txt
    cbInQue As Long
    cbOutQue As Long
End Type

Private Type COMMTIMEOUTS
    ReadIntervalTimeout As Long
    ReadTotalTimeoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long
End Type

'
' The DCB structure defines the control setting for a serial
' communications device.
'
Private Type DCB
    DCBlength As Long
    BaudRate As Long
    fBitFields As Long ' See Comments in Win32API.Txt
    wReserved As Integer
    XonLim As Integer
    XoffLim As Integer
    ByteSize As Byte
    Parity As Byte
    StopBits As Byte
    XonChar As Byte
    XoffChar As Byte
    ErrorChar As Byte
    EofChar As Byte
    EvtChar As Byte
    wReserved1 As Integer 'Reserved; Do Not Use
End Type

Private Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'-------------------------------------------------------------------------------
' System Functions
'-------------------------------------------------------------------------------
'
' Fills a specified DCB structure with values specified in
' a device-control string.
'
Private Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" _
    (ByVal lpDef As String, lpDCB As DCB) As Long
'
' Retrieves information about a communications error and reports
' the current status of a communications device. The function is
' called when a communications error occurs, and it clears the
' device's error flag to enable additional input and output
' (I/O) operations.
'
Private Declare Function ClearCommError Lib "kernel32" _
    (ByVal hFile As Long, lpErrors As Long, lpStat As COMSTAT) As Long
'
' Closes an open communications device or file handle.
'
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'
' Creates or opens a communications resource and returns a handle
' that can be used to access the resource.
'
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
    (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
'
' Directs a specified communications device to perform a function.
'
Private Declare Function EscapeCommFunction Lib "kernel32" _
    (ByVal nCid As Long, ByVal nFunc As Long) As Long
'
' Formats a message string such as an error string returned
' by anoher function.
'
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
    Arguments As Long) As Long
'
' Retrieves modem control-register values.
'
Private Declare Function GetCommModemStatus Lib "kernel32" _
    (ByVal hFile As Long, lpModemStat As Long) As Long
'
' Retrieves the current control settings for a specified
' communications device.
'
Private Declare Function GetCommState Lib "kernel32" _
    (ByVal nCid As Long, lpDCB As DCB) As Long
'
' Retrieves the calling thread's last-error code value.
'
Private Declare Function GetLastError Lib "kernel32" () As Long
'
' Retrieves the results of an overlapped operation on the
' specified file, named pipe, or communications device.
'
Private Declare Function GetOverlappedResult Lib "kernel32" _
    (ByVal hFile As Long, lpOverlapped As OVERLAPPED, _
    lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long
'
' Discards all characters from the output or input buffer of a
' specified communications resource. It can also terminate
' pending read or write operations on the resource.
'
Private Declare Function PurgeComm Lib "kernel32" _
    (ByVal hFile As Long, ByVal dwFlags As Long) As Long
'
' Reads data from a file, starting at the position indicated by the
' file pointer. After the read operation has been completed, the
' file pointer is adjusted by the number of bytes actually read,
' unless the file handle is created with the overlapped attribute.
' If the file handle is created for overlapped input and output
' (I/O), the application must adjust the position of the file pointer
' after the read operation.
'
Private Declare Function ReadFile Lib "kernel32" _
    (ByVal hFile As Long, ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, _
    lpOverlapped As OVERLAPPED) As Long
'
' Configures a communications device according to the specifications
' in a device-control block (a DCB structure). The function
' reinitializes all hardware and control settings, but it does not
' empty output or input queues.
'
Private Declare Function SetCommState Lib "kernel32" _
    (ByVal hCommDev As Long, lpDCB As DCB) As Long
'
' Sets the time-out parameters for all read and write operations on a
' specified communications device.
'
Private Declare Function SetCommTimeouts Lib "kernel32" _
    (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
'
' Initializes the communications parameters for a specified
' communications device.
'
Private Declare Function SetupComm Lib "kernel32" _
    (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
'
' Writes data to a file and is designed for both synchronous and a
' synchronous operation. The function starts writing data to the file
' at the position indicated by the file pointer. After the write
' operation has been completed, the file pointer is adjusted by the
' number of bytes actually written, except when the file is opened with
' FILE_FLAG_OVERLAPPED. If the file handle was created for overlapped
' input and output (I/O), the application must adjust the position of
' the file pointer after the write operation is finished.
'
Private Declare Function WriteFile Lib "kernel32" _
    (ByVal hFile As Long, ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, _
    lpOverlapped As OVERLAPPED) As Long


Private Declare Sub AppSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
'-------------------------------------------------------------------------------
' Program Constants
'-------------------------------------------------------------------------------

Private Const MAX_PORTS = 4

'-------------------------------------------------------------------------------
' Program Structures
'-------------------------------------------------------------------------------

Private Type COMM_ERROR
    lngErrorCode As Long
    strFunction As String
    strErrorMessage As String
End Type

Private Type COMM_PORT
    lngHandle As Long
    blnPortOpen As Boolean
    udtDCB As DCB
End Type

 

'-------------------------------------------------------------------------------
' Program Storage
'-------------------------------------------------------------------------------

Private udtCommOverlap As OVERLAPPED
Private udtCommError As COMM_ERROR
Private udtPorts(1 To MAX_PORTS) As COMM_PORT
'-------------------------------------------------------------------------------
' GetSystemMessage - Gets system error text for the specified error code.
'-------------------------------------------------------------------------------
Public Function GetSystemMessage(lngErrorCode As Long) As String
    Dim intPos As Integer
    Dim strMessage As String, strMsgBuff As String * 256

    Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, lngErrorCode, 0, strMsgBuff, 255, 0)

    intPos = InStr(1, strMsgBuff, vbNullChar)
    strMessage = strMsgBuff
    ' If intPos > 0 Then
    '     strMessage = Trim$(Left$(strMsgBuff, 1, intPos - 1))
    ' Else
    '     strMessage = Trim$(strMsgBuff)
    ' End If
    
    GetSystemMessage = strMessage
    
End Function

Public Function PauseApp(PauseInSeconds As Long)
    
    Call AppSleep(PauseInSeconds * 1000)
    
End Function

'-------------------------------------------------------------------------------
' CommOpen - Opens/Initializes serial port.
'
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'   strPort     - COM port name. (COM1, COM2, COM3, COM4)
'   strSettings - Communication settings.
'                 Example: "baud=9600 parity=E data=8 stop=1"
'
' Returns:
'   Error Code  - 0 = No Error.
'
'-------------------------------------------------------------------------------
Public Function CommOpen(intPortID As Integer, strPort As String, _
    strSettings As String) As Long
    
    Dim lngStatus       As Long
    Dim udtCommTimeOuts As COMMTIMEOUTS

    On Error GoTo Routine_Error
    
    ' See if port already in use.
    If udtPorts(intPortID).blnPortOpen Then
        lngStatus = -1
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommOpen"
            .strErrorMessage = "Port in use."
        End With
        MsgBox "error Open"
        GoTo Routine_Exit
    End If

    ' Open serial port.
    udtPorts(intPortID).lngHandle = CreateFile(strPort, GENERIC_READ Or _
        GENERIC_WRITE, 0, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    If udtPorts(intPortID).lngHandle = -1 Then
        MsgBox "Open Comm error -1 " & strPort
        lngStatus = SetCommError("CommOpen (CreateFile)")
        GoTo Routine_Exit
    End If

    udtPorts(intPortID).blnPortOpen = True

    ' Setup device buffers (1K each).
    lngStatus = SetupComm(udtPorts(intPortID).lngHandle, 1024, 1024)
    
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetupComm)")
        GoTo Routine_Exit
    End If

    ' Purge buffers.
    lngStatus = PurgeComm(udtPorts(intPortID).lngHandle, PURGE_TXABORT Or _
        PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (PurgeComm)")
        GoTo Routine_Exit
    End If

    ' Set serial port timeouts.
    With udtCommTimeOuts
        .ReadIntervalTimeout = -1
        .ReadTotalTimeoutMultiplier = 0
        .ReadTotalTimeoutConstant = 1000
        .WriteTotalTimeoutMultiplier = 0
        .WriteTotalTimeoutMultiplier = 1000
    End With

    lngStatus = SetCommTimeouts(udtPorts(intPortID).lngHandle, udtCommTimeOuts)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetCommTimeouts)")
        GoTo Routine_Exit
    End If

    ' Get the current state (DCB).
    lngStatus = GetCommState(udtPorts(intPortID).lngHandle, _
        udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (GetCommState)")
        GoTo Routine_Exit
    End If

    ' Modify the DCB to reflect the desired settings.
    lngStatus = BuildCommDCB(strSettings, udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (BuildCommDCB)")
        GoTo Routine_Exit
    End If

    ' Set the new state.
    lngStatus = SetCommState(udtPorts(intPortID).lngHandle, _
        udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetCommState)")
        GoTo Routine_Exit
    End If

    lngStatus = 0

Routine_Exit:
        CommOpen = lngStatus
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommOpen"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
End Function

Private Function SetCommError(strFunction As String) As Long
    
    With udtCommError
        .lngErrorCode = Err.LastDllError
        .strFunction = strFunction
        .strErrorMessage = GetSystemMessage(.lngErrorCode)
        SetCommError = .lngErrorCode
    End With
    
End Function

Private Function SetCommErrorEx(strFunction As String, lngHnd As Long) As Long
    Dim lngErrorFlags As Long
    Dim udtCommStat As COMSTAT
    
    With udtCommError
        .lngErrorCode = GetLastError
        .strFunction = strFunction
        .strErrorMessage = GetSystemMessage(.lngErrorCode)
    
        Call ClearCommError(lngHnd, lngErrorFlags, udtCommStat)
    
        .strErrorMessage = .strErrorMessage & "  COMM Error Flags = " & _
            Hex$(lngErrorFlags)
        
        SetCommErrorEx = .lngErrorCode
    End With
    
End Function

'-------------------------------------------------------------------------------
' CommSet - Modifies the serial port settings.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'   strSettings - Communication settings.
'                 Example: "baud=9600 parity=E data=8 stop=1"
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommSet(intPortID As Integer, strSettings As String) As Long
    
    Dim lngStatus As Long
    
    On Error GoTo Routine_Error

    lngStatus = GetCommState(udtPorts(intPortID).lngHandle, _
        udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (GetCommState)")
        GoTo Routine_Exit
    End If

    lngStatus = BuildCommDCB(strSettings, udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (BuildCommDCB)")
        GoTo Routine_Exit
    End If

    lngStatus = SetCommState(udtPorts(intPortID).lngHandle, _
        udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (SetCommState)")
        GoTo Routine_Exit
    End If

    lngStatus = 0

Routine_Exit:
        CommSet = lngStatus
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommSet"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommClose - Close the serial port.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommClose(intPortID As Integer) As Long
    
    Dim lngStatus As Long
    
    On Error GoTo Routine_Error

    If udtPorts(intPortID).blnPortOpen Then
        lngStatus = CloseHandle(udtPorts(intPortID).lngHandle)
    
        If lngStatus = 0 Then
            lngStatus = SetCommError("CommClose (CloseHandle)")
            GoTo Routine_Exit
        End If
    
        udtPorts(intPortID).blnPortOpen = False
    End If

    lngStatus = 0

Routine_Exit:
        CommClose = lngStatus
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommClose"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommFlush - Flush the send and receive serial port buffers.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommFlush(intPortID As Integer) As Long
    
    Dim lngStatus As Long
    
    On Error GoTo Routine_Error

    lngStatus = PurgeComm(udtPorts(intPortID).lngHandle, PURGE_TXABORT Or _
        PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommFlush (PurgeComm)")
        GoTo Routine_Exit
    End If

    lngStatus = 0

Routine_Exit:
        CommFlush = lngStatus
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommFlush"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommRead - Read serial port input buffer.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'   strData     - Data buffer.
'   lngSize     - Maximum number of bytes to be read.
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommRead(intPortID As Integer, strData As String, _
    lngSize As Long) As Long

    Dim lngStatus As Long
    Dim lngRdSize As Long, lngBytesRead As Long
    Dim lngRdStatus As Long, strRdBuffer As String * 1024
    Dim lngErrorFlags As Long, udtCommStat As COMSTAT
    
    On Error GoTo Routine_Error

    strData = ""
    lngBytesRead = 0
    DoEvents
    
    ' Clear any previous errors and get current status.
    lngStatus = ClearCommError(udtPorts(intPortID).lngHandle, lngErrorFlags, _
        udtCommStat)

    If lngStatus = 0 Then
        lngBytesRead = -1
        lngStatus = SetCommError("CommRead (ClearCommError)")
        GoTo Routine_Exit
    End If
        
    If udtCommStat.cbInQue > 0 Then
        If udtCommStat.cbInQue > lngSize Then
            lngRdSize = udtCommStat.cbInQue
        Else
            lngRdSize = lngSize
        End If
    Else
        lngRdSize = 0
    End If

    If lngRdSize Then
        lngRdStatus = ReadFile(udtPorts(intPortID).lngHandle, strRdBuffer, _
            lngRdSize, lngBytesRead, udtCommOverlap)

        If lngRdStatus = 0 Then
            lngStatus = GetLastError
            If lngStatus = ERROR_IO_PENDING Then
                ' Wait for read to complete.
                ' This function will timeout according to the
                ' COMMTIMEOUTS.ReadTotalTimeoutConstant variable.
                ' Every time it times out, check for port errors.

                ' Loop until operation is complete.
                While GetOverlappedResult(udtPorts(intPortID).lngHandle, _
                    udtCommOverlap, lngBytesRead, True) = 0
                                    
                    lngStatus = GetLastError
                                        
                    If lngStatus <> ERROR_IO_INCOMPLETE Then
                        lngBytesRead = -1
                        lngStatus = SetCommErrorEx( _
                            "CommRead (GetOverlappedResult)", _
                            udtPorts(intPortID).lngHandle)
                        GoTo Routine_Exit
                    End If
                Wend
            Else
                ' Some other error occurred.
                lngBytesRead = -1
                lngStatus = SetCommErrorEx("CommRead (ReadFile)", _
                    udtPorts(intPortID).lngHandle)
                GoTo Routine_Exit
            
            End If
        End If
        ' MsgBox Len(strRdBuffer) & "  " & strRdBuffer
        ' strData = Left(strRdBuffer, lngBytesRead)
        strData = strRdBuffer
    End If

Routine_Exit:
        CommRead = lngBytesRead
        Exit Function

Routine_Error:
        lngBytesRead = -1
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommRead"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommWrite - Output data to the serial port.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'   strData     - Data to be transmitted.
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommWrite(intPortID As Integer, strData As String) As Long
    
    Dim i As Integer
    Dim lngStatus As Long, lngSize As Long
    Dim lngWrSize As Long, lngWrStatus As Long
    
    On Error GoTo Routine_Error
    
    ' Get the length of the data.
    lngSize = Len(strData)

    ' Output the data.
    lngWrStatus = WriteFile(udtPorts(intPortID).lngHandle, strData, lngSize, _
        lngWrSize, udtCommOverlap)

    ' Note that normally the following code will not execute because the driver
    ' caches write operations. Small I/O requests (up to several thousand bytes)
    ' will normally be accepted immediately and WriteFile will return true even
    ' though an overlapped operation was specified.
        
    DoEvents
    
    If lngWrStatus = 0 Then
        lngStatus = GetLastError
        If lngStatus = 0 Then
            GoTo Routine_Exit
        ElseIf lngStatus = ERROR_IO_PENDING Then
            ' We should wait for the completion of the write operation so we know
            ' if it worked or not.
            '
            ' This is only one way to do this. It might be beneficial to place the
            ' writing operation in a separate thread so that blocking on completion
            ' will not negatively affect the responsiveness of the UI.
            '
            ' If the write takes long enough to complete, this function will timeout
            ' according to the CommTimeOuts.WriteTotalTimeoutConstant variable.
            ' At that time we can check for errors and then wait some more.

            ' Loop until operation is complete.
            While GetOverlappedResult(udtPorts(intPortID).lngHandle, _
                udtCommOverlap, lngWrSize, True) = 0
                                
                lngStatus = GetLastError
                                    
                If lngStatus <> ERROR_IO_INCOMPLETE Then
                    lngStatus = SetCommErrorEx( _
                        "CommWrite (GetOverlappedResult)", _
                        udtPorts(intPortID).lngHandle)
                    GoTo Routine_Exit
                End If
            Wend
        Else
            ' Some other error occurred.
            lngWrSize = -1
                    
            lngStatus = SetCommErrorEx("CommWrite (WriteFile)", _
                udtPorts(intPortID).lngHandle)
            GoTo Routine_Exit
        
        End If
    End If
    
    For i = 1 To 10
        DoEvents
    Next
    
Routine_Exit:
        CommWrite = lngWrSize
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommWrite"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommGetLine - Get the state of selected serial port control lines.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'   intLine     - Serial port line. CTS, DSR, RING, RLSD (CD)
'   blnState    - Returns state of line (Cleared or Set).
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommGetLine(intPortID As Integer, intLine As Integer, _
    blnState As Boolean) As Long
    
    Dim lngStatus As Long
    Dim lngComStatus As Long, lngModemStatus As Long
    
    On Error GoTo Routine_Error

    lngStatus = GetCommModemStatus(udtPorts(intPortID).lngHandle, lngModemStatus)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommReadCD (GetCommModemStatus)")
        GoTo Routine_Exit
    End If

    If (lngModemStatus And intLine) Then
        blnState = True
    Else
        blnState = False
    End If
        
    lngStatus = 0
        
Routine_Exit:
        CommGetLine = lngStatus
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommReadCD"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommSetLine - Set the state of selected serial port control lines.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'   intLine     - Serial port line. BREAK, DTR, RTS
'                 Note: BREAK actually sets or clears a "break" condition on
'                 the transmit data line.
'   blnState    - Sets the state of line (Cleared or Set).
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommSetLine(intPortID As Integer, intLine As Integer, _
    blnState As Boolean) As Long
   
    Dim lngStatus As Long
    Dim lngNewState As Long
    
    On Error GoTo Routine_Error
    
    If intLine = LINE_BREAK Then
        If blnState Then
            lngNewState = SETBREAK
        Else
            lngNewState = CLRBREAK
        End If
    
    ElseIf intLine = LINE_DTR Then
        If blnState Then
            lngNewState = SETDTR
        Else
            lngNewState = CLRDTR
        End If
    
    ElseIf intLine = LINE_RTS Then
        If blnState Then
            lngNewState = SETRTS
        Else
            lngNewState = CLRRTS
        End If
    End If

    lngStatus = EscapeCommFunction(udtPorts(intPortID).lngHandle, lngNewState)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSetLine (EscapeCommFunction)")
        GoTo Routine_Exit
    End If

    lngStatus = 0
        
Routine_Exit:
        CommSetLine = lngStatus
        Exit Function

Routine_Error:
        lngStatus = Err.Number
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommSetLine"
            .strErrorMessage = Err.Description
        End With
        Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommGetError - Get the last serial port error message.
'
' Parameters:
'   strMessage  - Error message from last serial port error.
'
' Returns:
'   Error Code  - Last serial port error code.
'-------------------------------------------------------------------------------
Public Function CommGetError(strMessage As String) As Long
    
    With udtCommError
        CommGetError = .lngErrorCode
        strMessage = "Error (" & CStr(.lngErrorCode) & "): " & .strFunction & _
            " - " & .strErrorMessage
    End With
    
End Function

Public Sub AutoRead_Click()
    Dim strvar As String
    Dim IntRowNow As Integer
    Dim ctlList As Control, varItem As Variant
    Dim ctlLineStatus As Control
    Dim rtcode As Integer
    Dim intRetryCounter As Integer
    Dim strShift As String
    Dim intHour As Integer
    Dim strLineStatus As String
    
    
    On Error GoTo ErrorHandler    ' Enable error-handling routine.
    
    BlnAutoRun = True

    LineStatusData.Selected(1) = True
    Do While BlnAutoRun
        ' Return Control object variable pointing to list box.
        Set ctlList = Forms![frmlineinfo]![LineStatusData]
        ' ctlList.Requery

        If ctlList.ListCount = 0 Then
            MsgBox "Line Status DataBase Empty", , "Warning"
            Exit Sub
        End If

        LineStatusData.Selected(1) = True
        intHour = Int(Hour(Now) / 8)
        intHour = IIf(intHour = 0, 3, intHour)
        strShift = "Shift" & intHour

        If strShift <> LineStatusData.Column(3) Then
        
            Dim fileNum As Integer
            'fileNum = FreeFile
            'Open "C:\Users\Nudam\Desktop\shiftChange.log" For Append As #fileNum
            
            Debug.Print (Now & "    Shift Changed")
            ' ShiftChange_Click
    
            strvar = "UPDATE ProductionLineStatus SET ShiftCounter = CounterStop - CounterStart"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strvar
            Debug.Print (Now & "    " & strvar)
            
            'Print #fileNum, Now & "    " & strvar
            TimerDelay (1)
    
            strvar = "INSERT INTO ProductionLineStatusHistory SELECT ProductionDate, ProductionLineNo, ShiftNO, "
            strvar = strvar & "MoldNo, ColorCode, SiloNo, LineStatus, StationNO, MachineModel, ShiftCounter, "
            strvar = strvar & "CounterStart, CounterStop, CounterLast, CycleTimeLast, StatusRemarks, MoldType FROM ProductionLineStatus"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strvar
    
    
            strvar = "UPDATE ProductionLineStatus SET CounterStart = CounterStop"
            strvar = strvar & ", ShiftNo='" & strShift & "', ProductionDate='" & IIf(intHour = 3, Date - 1, Date) & "'"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strvar
            Debug.Print (Now & "    " & strvar)
            
            'Print #fileNum, Now & "    " & strvar
            
            'Close #fileNum

            ctlList.Requery
        End If

        ' Enumerate through selected items.
        For IntRowNow = 1 To ctlList.ListCount
            
            If BlnAutoRun = False Then Exit Sub
            
            LineStatusData.Selected(IntRowNow) = True
            ClearRead_Click
            intRetryCounter = 0

            ' Read Counter
            Do While intRetryCounter < 3

                ClearRead_Click
                ReadCounter_Click

                ' TimerDelay (1)
                ' lngStatus = CommRead(intPortID, strIdentifier, 1)
                ' Forms![frmlineinfo]![txtNumber1] = Forms![frmlineinfo]![txtNumber1] & strIdentifier

                '                       ETX, End of Text
                If InStr(strIdentifier, Chr$(3)) < InStr(strIdentifier, " ") Then strIdentifier = Mid(strIdentifier, InStr(strIdentifier, " "), Len(strIdentifier))
                
                If Len(strIdentifier) < 8 Or InStr(strIdentifier, Chr$(3)) = 0 Or InStr(strIdentifier, " ") = 0 Then
                    strIdentifier = "0"
                Else
                    strvar = Mid(strIdentifier, InStr(strIdentifier, " "), InStr(strIdentifier, Chr$(3)) - InStr(strIdentifier, " "))
                    
                    If CheckNumeric(strvar) Then
                        strIdentifier = strvar
                        
                        Debug.Print (Now & "    [" & LineStatusData.Column(2) & "] Successfully read counter: " & Val(strvar) & ". CounterLast: " & LineStatusData.Column(13))
                        Forms![frmlineinfo]![ShotCountVal] = Val(strvar)
                        
                        strLineStatus = IIf(Val(strvar) > LineStatusData.Column(13), "Running", "Down")
                        
                        ' CounterLast means `CounterLatest`
                        strvar = "UPDATE ProductionLineStatus SET CounterLast=" & strvar & ", LineStatus='" & strLineStatus & "'"
                        strvar = strvar & " WHERE ID=" & str(LineStatusData.Column(0))
                        DoCmd.SetWarnings False
                        DoCmd.RunSQL strvar, 0
                        Debug.Print (Now & "    [" & LineStatusData.Column(2) & "] " & strvar)
                        
                        strvar = "UPDATE ProductionLineStatus SET CounterStop = CounterLast WHERE CounterLast > 0"
                        DoCmd.SetWarnings False
                        DoCmd.RunSQL strvar, 0
                        Debug.Print (Now & "    [" & LineStatusData.Column(2) & "] " & strvar)
                        
                        strvar = "INSERT INTO Productionlog (ProductionDate, ProductionLineNo, PollingRetry, ReadItem, ReadResult)"
                        strvar = strvar & "VALUES ('" & IIf(intHour = 3, Now - 1, Now) & "','" & LineStatusData.Column(2)
                        strvar = strvar & "', '" & intRetryCounter & "', 'Counter', '" & strIdentifier & "')"
                        DoCmd.SetWarnings False
                        DoCmd.RunSQL strvar, 0

                        strvar = TextDisplay
                        TextDisplay = strvar & ":=" & strIdentifier
                    End If
                    strIdentifier = ""
                End If
                intRetryCounter = IIf(strIdentifier = "0", intRetryCounter + 1, 5)
            Loop

            If intRetryCounter = 3 Then
                Debug.Print (Now & "    [" & LineStatusData.Column(2) & "] Failed to read counter")
                Forms![frmlineinfo]![ShotCountVal] = "ERR"

                strvar = "UPDATE ProductionLineStatus SET CounterLast = 0, LineStatus = 'Down'"
                strvar = strvar & " WHERE ID=" & str(LineStatusData.Column(0))
                DoCmd.SetWarnings False
                DoCmd.RunSQL strvar, 0
                Debug.Print (Now & "    [" & LineStatusData.Column(2) & "] " & strvar)
                
                strvar = "INSERT INTO ProductionLog (ProductionDate, ProductionLineNo, PollingRetry, ReadItem, ReadResult)"
                strvar = strvar & "VALUES ('" & IIf(intHour = 3, Now - 1, Now) & "', '" & LineStatusData.Column(2)
                strvar = strvar & "', '" & intRetryCounter & "', 'Counter', '" & strIdentifier & "')"
                DoCmd.SetWarnings False
                DoCmd.RunSQL strvar, 0
            End If

            strvar = TextDisplay
            TextDisplay = strvar & "   --- ProductionLineNo:=" & LineStatusData.Column(2)
            intRetryCounter = IIf(strLineStatus = "Running", 0, 5)

            ' Read CycleTime
            Do While intRetryCounter < 3
                ' Exit Do

                ClearRead_Click
                ReadCycleTime_Click

                ' TimerDelay (1)
                ' lngStatus = CommRead(intPortID, strIdentifier, 1)
                ' Forms![frmlineinfo]![txtNumber1] = Forms![frmlineinfo]![txtNumber1] & strIdentifier
                
                If InStr(strIdentifier, Chr$(3)) < InStr(strIdentifier, " ") Then strIdentifier = Mid(strIdentifier, InStr(strIdentifier, " "), Len(strIdentifier))
                
                If Len(strIdentifier) < 8 Or InStr(strIdentifier, Chr$(3)) = 0 Or InStr(strIdentifier, " ") = 0 Then
                    strIdentifier = "0"
                Else
                    strvar = Mid(strIdentifier, InStr(strIdentifier, " "), InStr(strIdentifier, Chr$(3)) - InStr(strIdentifier, " "))
                    
                    If CheckNumeric(strvar) Then
                        strIdentifier = strvar

                        Debug.Print (Now & "    [" & LineStatusData.Column(2) & "] Successfully read cycle time: " & Val(strvar))
                        Forms![frmlineinfo]![CycleTimeVal] = Val(strvar)

                        strvar = "UPDATE ProductionLineStatus SET CycleTimeLast = " & strvar
                        strvar = strvar & " WHERE ID=" & str(LineStatusData.Column(0))
                        DoCmd.SetWarnings False
                        DoCmd.RunSQL strvar, 0
                        Debug.Print (Now & "    [" & LineStatusData.Column(2) & "] " & strvar)
                        
                        strvar = "INSERT INTO ProductionLog (ProductionDate, ProductionLineNo, PollingRetry, ReadItem, ReadResult)"
                        strvar = strvar & " VALUES ('" & IIf(intHour = 3, Now - 1, Now) & "', '" & LineStatusData.Column(2)
                        strvar = strvar & "', '" & intRetryCounter & "', 'CycleTime', '" & strIdentifier & "')"
                        DoCmd.SetWarnings False
                        DoCmd.RunSQL strvar, 0

                        strvar = TextDisplay
                        TextDisplay = strvar & ":=" & strIdentifier
                    End If
                    strIdentifier = ""
                End If
                intRetryCounter = IIf(strIdentifier = "0", intRetryCounter + 1, 5)
            Loop

            If intRetryCounter = 3 Then
                Debug.Print (Now & "    [" & LineStatusData.Column(2) & "] Failed to read cycle time")
                Forms![frmlineinfo]![CycleTimeVal] = "ERR"

                strvar = "INSERT INTO ProductionLog (ProductionDate, ProductionLineNo, PollingRetry, ReadItem, ReadResult)"
                strvar = strvar & " VALUES ('" & IIf(intHour = 3, Now - 1, Now) & "', '" & LineStatusData.Column(2)
                strvar = strvar & "', '" & intRetryCounter & "', 'CycleTime', '" & strIdentifier & "')"
                DoCmd.SetWarnings False
                DoCmd.RunSQL strvar, 0
            End If

            If BlnAutoRun = False Then Exit Sub
        Next IntRowNow

        strIdentifier = ""
        
        strvar = "UPDATE ProductionLineStatus SET ShiftCounter = CounterLast - CounterStart WHERE CounterLast > 0"
        DoCmd.SetWarnings False
        DoCmd.RunSQL strvar, 0
        Debug.Print (Now & "    " & strvar)
    Loop
    Exit Sub

ErrorHandler:
        ' strvar = "Insert into RunErrorLog (EventTime, ErrorNo, ErrorText, ErrorSource) values ('" & Now & "','" & Err.Number & "','" & Err.Description & "','" & Err.Source & "')"
        ' DoCmd.RunSQL strvar, 0
        
        DoCmd.Close
        DoCmd.Quit
        Resume
End Sub

Private Sub ClearRead_Click()
    Dim strvar As String
    Dim retCode As Integer
    
    Forms![frmlineinfo]![txtNumber1] = ""
End Sub

Private Sub Form_Load()
    
    RefreshLineDetails.RefreshLineDetails
        
    ImageGreen.Visible = False
    ImageWarning.Visible = False
    ImageStop.Visible = True

    strGreenSign = "Green"
    strYellowSign = "Yellow"
    strStopSign = "Stop"
    BlnAutoRun = False
    Me.TimerInterval = 5000

    If udtPorts(intPortID).blnPortOpen Then CommClose (intPortID)

    If udtPorts(intPortID).blnPortOpen Then
        DoCmd.Close
        DoCmd.Quit
    End If

    blnrun = True
End Sub

Private Sub ChangeSign(StrSignNo As String)
    ImageGreen.Visible = False
    ImageWarning.Visible = False
    ImageStop.Visible = False

    Select Case StrSignNo
    Case strGreenSign
        ImageGreen.Visible = True
    Case strYellowSign
        ImageWarning.Visible = True
    Case strStopSign
        ImageStop.Visible = True
    End Select
End Sub


Private Sub Form_Timer()
    If blnrun = True Then
    'If True Then
        blnrun = False
        Me.TimerInterval = 1200000
        Me.SetFocus
        
        lngStatus = CommOpen(intPortID, "COM" & CStr(intPortID) & ":", "baud=9600 parity=E data=8 stop=1")
        If lngStatus <> 0 Then
            DoCmd.Close
            DoCmd.Quit
        End If
        
        ChangeSign (strGreenSign)
        DoCmd.OpenForm ("formRun")
        ' AutoRead.SetFocus
        ' SendKeys " ", False
    Else
        DoCmd.Close
        DoCmd.Quit
    End If
End Sub

Private Sub ImageGreen_Click()
    If udtPorts(intPortID).blnPortOpen Then CommClose (intPortID)

    ChangeSign (strStopSign)

    If udtPorts(intPortID).blnPortOpen Then
        MsgBox "cannot Close comm port "
        ChangeSign (strYellowSign)
    End If
End Sub

Private Sub ImageStop_Click()
    ChangeSign (strGreenSign)
    lngStatus = CommOpen(intPortID, "COM" & CStr(intPortID) & ":", "baud=19200 parity=E data=7 stop=1")

    If lngStatus <> 0 Then
        MsgBox "cannot open comm port "
        ChangeSign (strYellowSign)
    End If
End Sub

Private Sub ImageWarning_Click()
    If udtPorts(intPortID).blnPortOpen Then CommClose (intPortID)

    ChangeSign (strStopSign)

    If udtPorts(intPortID).blnPortOpen Then
        MsgBox "cannot Close comm port "
        ChangeSign (strYellowSign)
    End If
End Sub

Private Sub ReadCounter_Click()
    Dim strvar As String

    If Not udtPorts(intPortID).blnPortOpen Then
        MsgBox "Please Open Comm port", , "Warning"
        ChangeSign (strYellowSign)
        BlnAutoRun = False
    Else
        strIdentifier = ""
        strvar = ReadStation(LineStatusData.ItemData(LineStatusData.ListIndex + 1), "C307")
    End If
End Sub

Private Function ReadStation(strStationNo, strCommand As String)
    If IsNull(strStationNo) Or IsNull(strCommand) Then
        TextDisplay = "No Line Selected !"
    Else
        TextDisplay = "Reading ---> Station: " & strStationNo & " -- Item: " & strCommand
        lngStatus = CommWrite(intPortID, "{" + strStationNo + Chr$(4) + strCommand + Chr$(5) + Chr$(13))
        TimerDelay (1)
        lngStatus = CommRead(intPortID, strIdentifier, 1)
        Forms![frmlineinfo]![txtNumber1] = Forms![frmlineinfo]![txtNumber1] & strIdentifier
    End If
End Function

Private Function TimerDelay(intPauseTime)
    Dim intStartTime
    Dim intDate As Date
    
    intDate = Now
    
    intStartTime = Timer    ' Set start time.
    Do While Timer < intStartTime + intPauseTime
        DoEvents    ' Yield to other processes.
        If Day(intDate) <> Day(Now) Then
            Return
        End If
    Loop
End Function

Private Sub ReadCycleTime_Click()
    Dim strvar As String
    
    If Not udtPorts(intPortID).blnPortOpen Then
        MsgBox "Please Open Comm port", , "Warning"
        ChangeSign (strYellowSign)
        BlnAutoRun = False
    Else
        strIdentifier = ""
        strvar = ReadStation(LineStatusData.ItemData(LineStatusData.ListIndex + 1), "T300")
    End If
End Sub

Private Sub ShiftChange_Click()
    Dim strvar As String
    Dim strShift As String
    Dim intHour As Integer
    
    intHour = Int(Hour(Now) / 8)
    intHour = IIf(intHour = 0, 3, intHour)
    strShift = "Shift" & intHour
    
    strvar = "UPDATE ProductionLineStatus SET ShiftCounter=CounterStop - CounterStart"
    DoCmd.SetWarnings False
    DoCmd.RunSQL strvar, 0
    
    strvar = "INSERT INTO ProductionLineStatusHistory SELECT ProductionDate,ProductionLineNo,ShiftNO,"
    strvar = strvar & "MoldNo,ColorCode,SiloNo,LineStatus,StationNO,MachineModel,ShiftCounter,"
    strvar = strvar & "CounterStart,CounterStop,CounterLast,CycleTimeLast,StatusRemarks FROM ProductionLineStatus"
    DoCmd.SetWarnings False
    DoCmd.RunSQL strvar, 0
    
    strvar = "UPDATE ProductionLineStatus SET CounterStart = CounterStop"
    strvar = strvar & ", ShiftNo='" & strShift & "', ProductionDate='" & IIf(strShift = "3", Date - 1, Date) & "'"
    DoCmd.SetWarnings False
    DoCmd.RunSQL strvar, 0
End Sub

Private Sub StopAutoRead_Click()
    BlnAutoRun = False
End Sub

Private Function CheckNumeric(strTemp As String)
    Dim i As Integer
    
    CheckNumeric = True
    For i = 1 To Len(strTemp)
        If Mid(strTemp, i, 1) <> " " And Mid(strTemp, i, 1) <> "." And IsNumeric(Mid(strTemp, i, 1)) = False Then
            CheckNumeric = False
            i = Len(strTemp) + 1
        End If
    Next
End Function
