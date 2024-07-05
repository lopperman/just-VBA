VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  Logging
'  ** REQUIRES 'pbCommonUtil' module **
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbLOG.cls
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1

Public Enum LogType
    ltNOT_SET = -1
    ltTRACE = 0
    ltdebug = 1
    ltWARN = 2
    ltERROR = 3
    ltFORCED = 10
End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'       When CALLING 'Log'
'       will add log message to:
'       - DirectoryName:
'         Application.DefaultFilePath/[LOG_DIR]
'       - LogFileName=workbookname & '_LOG_'  & [YYYYMMDD]
'   Each time pbLOG is initialized, the 'Clean' funtion will be called and
'       will check and purge logs where the date [YYYYMMDD] in the filename
'       Is >= [LOG_MAXAGE_DAYS]
'       If[LOG_MAXAGE_DAYS] is 0 (zero), then all logs except current log
'           will be deleted
'   'LOG_MAXAGE_DAYS'
'       If set to 0 (zero), all previous logs matching log filename
'           pattern ([workbookname]_log_[yyyymmdd]) will
'           be deleted
'       if set to > 0, all previous logs matching log filename
'           pattern ([workbookname]_log_[yyyymmdd]) will
'           be delete if the datedifference between date
'           from log name (yyyymmdd) and NOW is
'           >= to 'LOG_MAXAGE_DAYS'
'       If set to < 0 (e.g.-1), then logs will not be cleaned automatically
    Private Const LOG_DIR As String _
        = "PBCOMMONLOG"
    Private Const LOG_MAXAGE_DAYS As Long _
        = 30
'   'pbLogFileNumber' is used to store FreeFile when
'       keeping logFile open. If you're expecting to write more
'       than a few log messages, performance is significatnly
'       increased when file is kept open
    Private logFileNumber
    Private lCurrentLogLevel As LogType
    Private lwarnCount As Long
    Private lerrorCount As Long
    Private wbNameNoExtension As String
    Private loggingInitialized As Boolean
'    Private WithEvents cEvents As pbCommonEvents
    Private lastLogMsg As Single
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'
'       PB COMMON LOGGING
'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Get Initialized() As Boolean
    Initialized = loggingInitialized
End Property


Public Function InitLogging(loggingLevel As LogType, Optional commonEvts As pbCommonEvents, Optional wkbk As Workbook, Optional runAutoClean As Boolean = True)
    On Error Resume Next
    If loggingInitialized = False Then
        lCurrentLogLevel = loggingLevel
        If wkbk Is Nothing Then
            Set wkbk = ThisWorkbook
        End If
        wbNameNoExtension = FileNameWithoutExtension(wkbk.Name)
        loggingInitialized = True
        lastLogMsg = Timer
        LOG "** Logging Initialized **, User: " & ENV_LogName, ltFORCED
        LOG "** Logging To: " & LogPath, ltFORCED
        LogLevel = loggingLevel
        LOG "** FileName **: " & wkbk.FullName, ltFORCED
        Set wkbk = Nothing
        If runAutoClean = True Then
            LogDelete
        End If
    End If
    If Err.number <> 0 Then
        loggingInitialized = False
        If IsDev Then
            Beep
            Stop
        End If
        Err.Clear
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   returns error object if not using pbcommong_log
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function LogPath() As String
    If Not loggingInitialized Then
        Err.Raise 1004, Description:="pbLog is not initialized!"
    End If
    Dim logFileName As String
    LogPath = PathCombine(False, Application.DefaultFilePath _
        , LOG_DIR _
        , ConcatWithDelim("_", wbNameNoExtension, "LOG", Format(Date, "YYYYMMDD") & ".log"))
End Function
Private Function LogDirectory() As String
    LogDirectory = PathCombine(True, Application.DefaultFilePath, LOG_DIR)
End Function

Private Function LogOpen()
    'if the log file is open and this method is called, will close and then re-open
    If VAL(logFileNumber) > 0 Then
        Close #logFileNumber
    End If
    logFileNumber = FreeFile
    On Error Resume Next
    Open LogPath For Append As #logFileNumber
    If Err.number <> 0 Then
        Err.Clear
        On Error GoTo 0
        CreateDirectory PathCombine(True, Application.DefaultFilePath, LOG_DIR)
        Open LogPath For Append As #logFileNumber
    End If
End Function
Public Function LogClose()
    On Error Resume Next
    If VAL(logFileNumber) > 0 Then
        Close #logFileNumber
        DoEvents
        logFileNumber = Empty
    End If
End Function

Public Function LaunchLog()
    On Error Resume Next
    #If Mac Then
        ThisWorkbook.FollowHyperlink LogPath
    #Else
        Dim pCmd As String
        pCmd = Concat("powershell -NoExit gc '", LogPath, "' -wait -tail 1")
        ''CreateObject("WScript.Shell").Exec (pCmd)
        Shell pCmd, vbNormalFocus
    #End If
End Function

Public Function ChangeLogLevel(Optional newLevel As LogType = LogType.ltNOT_SET, Optional userConfirmation As Boolean = False)
    If newLevel = ltNOT_SET And userConfirmation = False Then
        Exit Function
    End If
    If userConfirmation = False Then
        LogLevel = newLevel
    Else
        If newLevel <> ltNOT_SET Then
            If MsgBox("Change Log Level To: " & LogTypeDesc(newLevel) & "?", vbQuestion + vbYesNo + vbDefaultButton1, "LOG LEVEL") = vbYes Then
                LogLevel = newLevel
            End If
        Else
            Dim msg As String
            msg = "Enter Number for desired logging level: " & vbNewLine & vbNewLine & _
                "0 - TRACE (All Messages Logged, May Affect Performance)" & vbNewLine & vbNewLine & _
                "1 - DEBUG (Debugging Messages Logged)" & vbNewLine & vbNewLine & _
                "2 - WARNING (Default - Messages Logged for significant or uncommon operations)" & vbNewLine & vbNewLine & _
                "3 - ERROR (All Errors Logged)" & vbNewLine
            Dim userResp
            userResp = InputBox(msg, "Log Level", 2)
            If IsNumeric(userResp) Then
                If CLng(userResp) >= 0 And CLng(userResp) <= 4 Then
                    LogLevel = CLng(userResp)
                End If
            End If
        End If
    End If
End Function

Public Function GetCurrentLogLines() As Collection
    On Error Resume Next
    Dim resp As New Collection
    LogClose
    Dim hf As Integer: hf = FreeFile
    Dim lines() As String
    Dim i As Long
    Open LogPath For Input As #hf
    lines = Split(Input$(LOF(hf), #hf), vbNewLine)
    Close #hf
    For i = LBound(lines) To UBound(lines)
        resp.Add lines(i)
    Next i
    Set GetCurrentLogLines = resp
    
End Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Log [msg] to[pbLogPath()]
'   if there is an open fileNumber for log, that will be used.
'   to open file for append, call 'pbLogOpen'
'   to close file call 'pbLogClose'
'   if [toOpenFile] is > 0, must be a valid open file
'   if [keepFileOpen] = true, will not close file after write
'       ** YOU ARE RESPONSIBLE FOR CLOSING OPEN FILES
'       and keeping track of filenumber for further use!
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function LOG(msg, Optional lType As LogType = LogType.ltTRACE, Optional closeLog As Boolean = True)
    If CanLog = False Then Exit Function
    
    If IsDev Then
        If stg.SettingWithDefault(STG_DEV_OUTPUT_TRACE, False) = True Then
            Debug.Print Concat(Now(), " - " & "LogType: ", lType, " MSG: ", msg & " | " & GetPerfText)
        End If
    End If
    If Not loggingInitialized Then
        Err.Raise 1004, Description:="pbLog is not initialized!"
    End If
        
    If lType = ltWARN Then
        lwarnCount = lwarnCount + 1
    ElseIf lType = ltERROR Then
        lerrorCount = lerrorCount + 1
    End If
    msg = CStr(msg)
    If lType >= LogLevel Then
        If VAL(logFileNumber) = 0 Then
            LogOpen
        Else
            If (Timer - lastLogMsg) > 15 Then
                LogClose
                LogOpen
            End If
        End If
        msg = ConcatWithDelim(" | ", NowWithMS, LogTypeDesc(lType), SysStates, msg)
        Write #logFileNumber, msg
        If closeLog Then
            LogClose
        End If
    End If
End Function
Public Property Get LogWarningCount() As Long
    LogWarningCount = lwarnCount
End Property
Public Property Get LogErrorCount() As Long
    LogErrorCount = lerrorCount
End Property
Public Function LogTRACE(msg, Optional keepLogOpen As Boolean = False)
    LOG CStr(msg), lType:=ltTRACE, closeLog:=Not keepLogOpen
End Function
Public Function LogDEBUG(msg, Optional keepLogOpen As Boolean = False)
    LOG CStr(msg), lType:=ltdebug, closeLog:=Not keepLogOpen
End Function
Public Function LogWARN(msg, Optional keepLogOpen As Boolean = False)
    LOG CStr(msg), lType:=ltWARN, closeLog:=Not keepLogOpen
End Function
Public Function LogERROR(msg, Optional keepLogOpen As Boolean = False)
    If Not StringsMatch(msg, "Error", smStartsWithStr) Then
        msg = Concat("ERROR ", msg)
    End If
    LOG CStr(msg), lType:=ltERROR, closeLog:=Not keepLogOpen
End Function
Public Function LogFORCED(msg, Optional keepLogOpen As Boolean = False)
    LOG CStr(msg), lType:=ltFORCED, closeLog:=Not keepLogOpen
End Function

Public Function LogTypeDesc(lType As LogType) As String
    If lType = ltERROR Then
        LogTypeDesc = "*ERROR*"
    ElseIf lType = ltWARN Then
        LogTypeDesc = "*WARNING*"
    ElseIf lType = ltdebug Then
        LogTypeDesc = "*DEBUG*"
    ElseIf lType = ltFORCED Then
        LogTypeDesc = "*TRACE-FORCED*"
    Else
        LogTypeDesc = "*TRACE*"
    End If
End Function
Public Property Get LogLevel() As LogType
    LogLevel = lCurrentLogLevel
End Property
Public Property Let LogLevel(lgLevel As LogType)
    If Not loggingInitialized Then
        Err.Raise 1004, Description:="pbLog is not initialized!"
    End If
    If lgLevel <> LogLevel Then
        lCurrentLogLevel = lgLevel
        LOG "Log Level Changed: " & LogTypeDesc(lgLevel), ltFORCED
    End If
End Property
Public Function LogDelete(Optional removeAllButCurrent As Boolean = False, Optional overrideMaxAgeDays)
    If Not loggingInitialized Then
        Err.Raise 1004, Description:="pbLog is not initialized!"
    End If
    Dim maxAgeDays As Long: maxAgeDays = LOG_MAXAGE_DAYS
    If Not IsMissing(overrideMaxAgeDays) Then
        maxAgeDays = CLng(overrideMaxAgeDays)
    End If
    If removeAllButCurrent = True Then maxAgeDays = 0
    LOG ConcatWithDelim(" ", "Cleaning Log Files ('pbLogDelete')", "removeAllButCurrent:", removeAllButCurrent _
        , "overrideMaxAgeDays set?:", Not IsMissing(overrideMaxAgeDays), "maxAgeDays:", maxAgeDays), ltFORCED
    Dim testDt
    Dim logFileName As Variant
    Dim tmpText As String, tmpDimension As Long
    Dim delFilePath As String
    Dim logFiles() As Variant
    logFiles = GetFiles(LogDirectory)
    LOG "Log File Count Start: " & UBound(logFiles) - LBound(logFiles) + 1, ltFORCED
    For Each logFileName In GetFiles(LogDirectory)
        If StringsMatch(logFileName, wbNameNoExtension, smStartsWithStr) And _
            StringsMatch(logFileName, ".log", smEndWithStr) And StringsMatch(logFileName, "_log_", smContains) Then
            tmpDimension = UBound(Split(logFileName, delimiter:="_LOG_", compare:=vbTextCompare))
            tmpText = Split(logFileName, delimiter:="_LOG_", compare:=vbTextCompare)(tmpDimension)
            tmpText = Left(tmpText, 8)
            testDt = DateSerial(CLng(Left(tmpText, 4)), CLng(Mid(tmpText, 5, 2)), CLng(Mid(tmpText, 7, 2)))
            If DtDiff(dtDay, testDt, Date) > maxAgeDays Or maxAgeDays = 0 Then
                delFilePath = PathCombine(False, LogDirectory, logFileName)
                If Not StringsMatch(delFilePath, LogPath) Then
                    LOG "Deleting Log File: " & delFilePath, ltFORCED
                    DeleteFile delFilePath
                End If
            End If
        End If
    Next logFileName
    logFiles = GetFiles(LogDirectory)
    LOG "Log File Count End: " & UBound(logFiles) - LBound(logFiles) + 1, ltFORCED
End Function

Public Function LogOpenCurrent()
    On Error Resume Next
    Performance = PerfBlockAll
    ThisWorkbook.FollowHyperlink LogPath
    Performance = PerfNormal
    If Err.number <> 0 Then
        LOG "Could not open current log file.  Error: " & Err.number, ltERROR
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'       CLEAR (DELETE) CURRENT LOG FILE
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function LogClear()
    LogClose
    If FileExists(LogPath) Then
        DeleteFile LogPath
        LOG "Current Log File Cleared By: " & ENV_LogName, ltFORCED
    End If
End Function


'Private Sub cEvents_OnErrorHandlerBegin(ByVal errNumber As Variant, ByVal errDescription As Variant, ByVal errSource As Variant)
'    Me.LogERROR ConcatWithDelim(" - ", "ErrorNumber: " & errNumber, "Desc: " & errDescription, "Source: " & errSource)
'End Sub

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'       pbCommonEvents Handlers
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'Private Sub cEvents_OnLog(ByVal msgType As LogType, ByVal msg As Variant, ByVal keepLogOpen As Boolean)
'    LOG msg, lType:=msgType, closeLog:=Not keepLogOpen
'End Sub
'
'Private Sub cEvents_OnLogLevelChange(ByVal msgType As LogType)
'    LogLevel = msgType
'End Sub

Private Sub Class_Terminate()
    LogClose
    
End Sub

