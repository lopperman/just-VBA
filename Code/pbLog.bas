Attribute VB_Name = "pbLog"
'-----------------------------------------------------------------------------------------------------
'
' [just-VBA] pbLog
' pbLog v1.0.0
' (c) 2022 Paul Brower - https://github.com/lopperman/just-VBA
'
' @license GNU General Public License v3.0
'-----------------------------------------------------------------------------------------------------
Option Explicit
Option Private Module
Option Base 1
Option Compare Text



Private lLog As pbEventLog
Private shtDown As Boolean

Private logCountInfo As Long
Private logCountTrace As Long
Private logCountWarn As Long
Private logCountErr As Long

Public Enum LogLevelEnum
    'Logger will log if Log Event >= Currrent Log Level
    [_Default] = 3
    llNONE = 0
    llINFO = 1
    llTrace = 2
    llWarn = 3
    llError = 4
    llDev = 100
End Enum




Public Property Get InfoLogCount() As Long
    InfoLogCount = logCountInfo
End Property
Public Property Get TraceLogCount() As Long
    TraceLogCount = logCountTrace
End Property
Public Property Get WarnLogCount() As Long
    WarnLogCount = logCountWarn
End Property
Public Property Get ErrorLogCount() As Long
    ErrorLogCount = logCountErr
End Property

Public Function ShutDownLog()
    shtDown = True
    CloseLog
    CheckReady
    Set lLog = Nothing
End Function

Public Property Get LogFolderPath() As String
    LogFolderPath = PathCombine(True, Application.DefaultFilePath, ThisWorkbook.CodeName & "LOG")
End Property

Public Function LogFilesExist(curWorkbookOnly As Boolean) As Boolean
    If curWorkbookOnly = False Then
        LogFilesExist = FileExists(PathCombine(False, LogFolderPath, "*.log"), True)
    Else
        LogFilesExist = FileExists(PathCombine(False, LogFolderPath, LogInstance.WBNameClean & "*"), True)
    End If
End Function

Public Function CurrentLogFileName() As String
    On Error Resume Next
    If Not lLog Is Nothing Then
        CurrentLogFileName = LogInstance.CurrentLogFileFullName
    End If
    If Err.number <> 0 Then Err.Clear
End Function

Public Function LogFilesStringArr() As String()
    Dim fileV As Variant: fileV = LogFiles
    Dim retV() As String
    If UBound(fileV) - LBound(fileV) - 1 >= 0 Then
        ReDim retV(0 To UBound(fileV) - 1)
        Dim i As Long
        For i = LBound(fileV) To UBound(fileV)
            retV(i - 1) = fileV(i, 1)
        Next i
    End If
    LogFilesStringArr = retV
End Function

Public Function LogFiles() As Variant
On Error Resume Next
    Dim tmpCol As New Collection
    If DirectoryFileCount(LogFolderPath) > 0 Then
        Dim myPath As Variant
        myPath = PathCombine(True, LogFolderPath)
        ChDir LogFolderPath
        Dim myFile, MyName As String
        MyName = Dir(myPath, vbNormal)
        Do While MyName <> ""
            If (GetAttr(PathCombine(False, myPath, MyName)) And vbNormal) = vbNormal Then
                tmpCol.Add MyName
            End If
            MyName = Dir()
        Loop
    End If
    If Not Err.number = 0 Then Err.Clear
    
    If tmpCol.Count > 0 Then
        Dim tFiles() As Variant
        ReDim tFiles(1 To tmpCol.Count, 1 To 1)
        Dim i As Long
        For i = 1 To tmpCol.Count
            tFiles(i, 1) = tmpCol(i)
        Next i
        LogFiles = tFiles
    End If
    If Err.number <> 0 Then Err.Clear

End Function

Public Function DeleteOldLogFiles()
On Error GoTo E:
    Dim failed As Boolean
    Dim filesDeleted As Long
        
    Dim curFiles As Variant, fai As ArrInformation
    curFiles = LogFiles
    fai = ArrayInfo(curFiles)
    If fai.Dimensions > 0 Then
        Dim i As Long
        For i = fai.LBound_first To fai.Ubound_first
            Dim fName As String: fName = curFiles(i, 1)
            If StringsMatch(fName, "_", smContains) And Len(fName) > 20 Then
                Dim tDt As String: tDt = Right(fName, 20)
                If Mid(tDt, 1, 1) = "_" And Mid(tDt, 10, 1) = "_" Then
                    tDt = Mid(tDt, 2, 8)
                    Dim testDate As Variant
                    testDate = DateSerial(Int(Mid(tDt, 1, 4)), Int(Mid(tDt, 5, 2)), Int(Mid(tDt, 7, 2)))
                    If DateDiff("d", testDate, Date) > 15 Then
                        DeleteLogFile fName
                        filesDeleted = filesDeleted + 1
                    End If
                End If
            End If
        Next i
    End If
        
Finalize:
    On Error Resume Next
    If filesDeleted > 0 Then
        LogTrace "pbLog.DeleteOldLogFiles - Deleted: " & filesDeleted & " files older than 15 days"
    End If
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize
End Function

Public Function DeleteLogFile(logFileName As String) As Boolean
    If Len(CurrentLogFileName) > 0 Then
        If StringsMatch(CurrentLogFileName, logFileName, smEndWithStr) Then
            MsgBox_FT "Cannot delete log file: " & logFileName & ", becase it is currently the active logging file", vbOKOnly + vbExclamation, "OOPS"
            Exit Function
        End If
    End If
    If FileExists(PathCombine(False, LogFolderPath, logFileName)) Then
        DeleteLogFile = DeleteFile(PathCombine(False, LogFolderPath, logFileName))
    Else
        DeleteLogFile = True
    End If
End Function

Public Function LogInstance() As pbEventLog
    On Error Resume Next
    If Not shtDown Then
        If lLog Is Nothing Then
            Set lLog = New pbEventLog
        End If
        Set LogInstance = lLog
    End If
    If Err.number <> 0 Then Err.Clear
End Function

Public Function CloseAndUploadLog()
    On Error Resume Next
    If Not lLog Is Nothing Then
        Dim toUpload As String: toUpload = lLog.CurrentLogFileName
        lLog.CloseLog
        DoEvents
        If Len(LogFileUploadPath) > 0 And Len(toUpload) > 0 Then
            If FileExists(toUpload) Then
            'If MsgBox_FT(ErrorLogCount & " errors have been logged locally. Is it ok to upload this log file to the FinTool SharePoint site to help improve the Financial Tool?", vbExclamation + vbYesNo + vbDefaultButton1, "UPLOAD LOG FILE") = vbYes Then
               LogInstance.UploadLogs toUpload
               DoEvents
            End If
        End If
    End If
    DeleteOldLogFiles
    If Err.number <> 0 Then Err.Clear

End Function

Public Function CloseLog()
    On Error Resume Next
    If Not lLog Is Nothing Then
        'Dim toUpload As String: toUpload = lLog.CurrentLogFileName
        lLog.CloseLog
'        If ErrorLogCount > 0 Then
'            If Len(LogFileUploadPath) > 0 Then
'                If MsgBox_FT(ErrorLogCount & " errors have been logged locally. Is it ok to upload this log file to the FinTool SharePoint site to help improve the Financial Tool?", vbExclamation + vbYesNo + vbDefaultButton1, "UPLOAD LOG FILE") = vbYes Then
'                   LogInstance.UploadLogs toUpload
'                End If
'            End If
'        End If
    End If
    DeleteOldLogFiles
    If Err.number <> 0 Then Err.Clear
End Function

Public Function SetLogOptions(ByVal lvl As LogLevelEnum)
    On Error Resume Next
    If Not shtDown And (DebugMode Or DebugInfo) Then
        If Not lLog Is Nothing Then
            lLog.LogFlash lvl
        Else
            Set lLog = New pbEventLog
            lLog.LogLevel = lvl
        End If
    End If
    If Err.number <> 0 Then Err.Clear
End Function

Public Function LogInfo(ByVal msg As String, Optional force As Boolean = False)
    On Error Resume Next
    If Not shtDown And (DebugMode Or DebugInfo) Then
        logCountInfo = logCountInfo + 1
        LogInstance.Log llINFO, msg, force
    End If
    If Err.number <> 0 Then Err.Clear
End Function
Public Function LogTrace(ByVal msg As String, Optional force As Boolean = False)
    On Error Resume Next
    If Not shtDown And (DebugMode Or DebugInfo) Then
        logCountTrace = logCountTrace + 1
        LogInstance.Log llTrace, msg, force
    End If
    If Err.number <> 0 Then Err.Clear
End Function
Public Function LogWarn(ByVal msg As String, Optional force As Boolean = False)
    On Error Resume Next
    If Not shtDown Then
        logCountWarn = logCountWarn + 1
        LogInstance.Log llWarn, msg, force
    End If
    If Err.number <> 0 Then Err.Clear

End Function
Public Function LogError(ByVal msg As String, Optional force As Boolean = False)
    On Error Resume Next
    If Not shtDown Then
        If Err.number <> 0 Then msg = msg & " (" & Err.number & " - " & Err.Description & ")"
        ftBeep btError
        logCountErr = logCountErr + 1
        LogInstance.Log llError, msg, force
    End If
    If Err.number <> 0 Then Err.Clear
End Function

Public Function LogDEV(ByVal msg As String)
    On Error Resume Next
    'If CurrentUser is not DEV, send to 'LogTrace' -- If user has Tracing On, might as well send this to the log
    If Not IsDeveloper Then
        LogTrace msg
        Exit Function
    End If
    If Not shtDown Then
        LogInstance.Log llDev, msg, True
    End If
    If IsDEV Then Debug.Print msg
    If Err.number <> 0 Then Err.Clear

End Function

Public Function GetLog(Optional ByVal logFileName As Variant)
On Error Resume Next
    If IsMissing(logFileName) Then
        If Len(CurrentLogFileName) > 0 Then
            logFileName = PathCombine(False, LogFolderPath, logFileName)
        End If
    End If
    Dim dataArray As Variant
    Dim ff As Integer
    ff = FreeFile()
    Open logFileName For Input As ff
    dataArray = ArrArray(Split(Input$(LOF(1), ff), vbLf), aoNone)
    Close ff

    Dim daAI As ArrInformation
    daAI = ArrayInfo(dataArray)
    If daAI.Rows > 0 Then
        'EVENT TYPE, DATE, TIME, SYSSTATES, MSG
        Dim idx As Long, colIDX As Long, newARR As Variant
        ReDim newARR(1 To daAI.Rows, 1 To 5)
        
        Dim tLine As Variant
        For idx = daAI.LBound_first To daAI.Ubound_first
                tLine = ArrArray(Split(dataArray(idx, 1), ", "), aoNone, True)
                If UBound(tLine, 2) = 5 Then
                    newARR(idx, 1) = tLine(1, 1)
                    newARR(idx, 2) = tLine(1, 2)
                    newARR(idx, 3) = "TM " & CStr(tLine(1, 3))
                    newARR(idx, 4) = tLine(1, 5)
                    newARR(idx, 5) = tLine(1, 4)
                End If
        Next idx
    End If

    Dim wkbk As Workbook, lWS As Worksheet, tfileName As String
    tfileName = FileNameFromFullPath(CStr(logFileName))
    tfileName = Left(tfileName, InStr(1, tfileName, ".") - 1)
    Set wkbk = Workbooks.Add
    Set lWS = wkbk.Worksheets(1)
    lWS.Name = tfileName
    lWS.Range("A1") = "LOGTYPE"
    lWS.Range("B1") = "DATE"
    lWS.Range("C1") = "TIME"
    lWS.Range("D1") = "LOG MSG"
    lWS.Range("E1") = "SYSTRC"
    lWS.Range("A2:E2").Resize(rowSize:=UBound(newARR)).value = newARR
    
    If Err.number <> 0 Then Err.Clear

End Function


