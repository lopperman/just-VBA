VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbEventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [just-VBA] pbEventLog
' pbEventLog v1.0.0
' (c) 2022 Paul Brower - https://github.com/lopperman/just-VBA
'
' @license GNU General Public License v3.0
'-----------------------------------------------------------------------------------------------------
Option Explicit
Option Base 1
Option Compare Text

Private lLogLevel As LogLevelEnum
Private lLogLevelFriendly As String
Private lFile As Integer
Private lOpened As Boolean
Private lPathChecked As Boolean
Private lInitLogMsg As Boolean
Private lPathValid As Boolean
Private lLogFolderPath As String
Private lcurLogFullName As String


Public Property Let LogLevel(lgLvl As LogLevelEnum)
    lLogLevel = lgLvl
End Property
Public Property Get LogLevel() As LogLevelEnum
    If DebugMode Then
        lLogLevel = llINFO
    Else
        lLogLevel = llWarn
    End If
    LogLevel = lLogLevel
End Property

Public Property Get CurrentLogFileName() As String
    If lOpened Then
        CurrentLogFileName = FileNameFromFullPath(lcurLogFullName)
    End If
End Property

Public Property Get CurrentLogFileFullName() As String
    If lOpened Then
        CurrentLogFileFullName = lcurLogFullName
    End If
End Property

Public Function LogLevelString(Optional logLvl As LogLevelEnum) As String
    If IsMissing(logLvl) Or logLvl = 0 Then
        logLvl = LogLevel
    End If
    LogLevelString = pbLog.LogLevelToString(logLvl)
End Function



Public Function Log(ByVal lvl As LogLevelEnum, ByVal msg As String, Optional force As Boolean = False)
    If force = False And lvl < LogLevel Then Exit Function

    If InStr(1, msg, """", vbTextCompare) Then msg = Replace(msg, """", "")
    If InStr(1, msg, ",") > 0 Then msg = Replace(msg, ",", " ")
    msg = Replace(msg, "  ", " ")
    If lOpened = False Then
        If LogPathValid Then
            OpenLog
        End If
    End If
    If lOpened Then
        On Error Resume Next
        Print #lFile, ConcatWithDelim(", ", LogLevelString(lvl), TimeStamp, SystemStates, msg)
'            Write #lFile, ConcatWithDelim(", ", LogLevelString(lvl), TimeStamp, SystemStates, msg)
        If DebugMode And lvl >= LogLevelEnum.llWarn Then
            If IsDEV Then Debug.Print ConcatWithDelim(", ", LogLevelString(lvl), TimeStamp, SystemStates, msg)
        End If
        If Err.number <> 0 Then
            If IsDEV Then
                Stop
            End If
            Err.Clear
        End If
    End If
End Function

Public Property Get TimeStamp()
    Dim tNOW As Variant: tNOW = Now
    Dim tTimer As Double: tTimer = CDbl(Timer)
    TimeStamp = Format(tNOW, "yyyy-mm-dd, hh:nn:ss") & "." & Format((tTimer - Fix(tTimer)) * 1000, "000")
End Property

Private Function SystemStates() As String
    Dim tEv As String, tSc As String, tIn As String, tCa As String, retV As String
    tEv = IIf(Events, "Evts=Y", "Evtn=N")
    tSc = IIf(Application.ScreenUpdating, "Scrn=Y", "Scrn=N")
    tIn = IIf(Application.Interactive, "Intr=Y", "Intr=N")
    tCa = IIf(Application.Calculation = xlCalculationAutomatic, "Calc=Y", "Calc=N")
    retV = ConcatWithDelim(" ", tEv, tSc, tIn, tCa)
    retV = Concat("SysStates:(", retV, ")")
    SystemStates = retV
End Function

Public Function LogFlash(ByVal lvl As LogLevelEnum)
    If lOpened Then
        CloseLog
    End If
    OpenLog
    LogLevel = lvl
End Function

Public Property Get WBNameClean() As String
        Dim tName As String
        tName = ThisWorkbook.Name
        tName = ReplaceIllegalCharacters(tName, "")
        tName = Replace(tName, "'", "", compare:=vbTextCompare)
        tName = Replace(tName, "-", "", compare:=vbTextCompare)
        If InStr(1, tName, ".") > 1 Then
            tName = Left(tName, InStr(1, tName, ".") - 1)
        End If
        WBNameClean = tName
End Property

Private Function OpenLog()
On Error GoTo E:
    If LogPathValid Then
        If Len(lcurLogFullName) = 0 Then
            Dim logFileName As String
            'logFileName = Concat(WBNameClean, "_", Format(Now, "yyyymmdd"), ".log")
            logFileName = Concat(WBNameClean, "_", Format(Now, "yyyymmdd_hhnnss"), ".log")
            lcurLogFullName = PathCombine(False, LogFolderPath, logFileName)
        End If
        lFile = FreeFile()
        Open lcurLogFullName For Append Access Write As lFile
        lOpened = True
        If lInitLogMsg = False Then
            lInitLogMsg = True
            Log llINFO, Concat("*** FILE  ***", " - ", ThisWorkbook.Name, " - Version v", AppVersion), True
            Log llINFO, "*** LOCATION ***" & FullPathExcludingFileName(FullWbNameCorrected), True
        End If
    Else
        lOpened = False
    End If
    Exit Function
E:
    lOpened = False
    '   CANNOT WRITE ERROR HERE IF ERROR IS WITH OPEN LOG FILE
    Dim errInfo As String: errInfo = ErrString
    If IsDEV Then
        ftBeep btError
        Debug.Print "pbEventLog.OpenLog - ERROR: " & errInfo
        Err.Clear
    End If

End Function

Public Function CloseLog()
    On Error Resume Next
    lPathValid = False
    lPathChecked = False
    If lOpened Then
        Close lFile
        lOpened = False
        lInitLogMsg = False
        lcurLogFullName = vbNullString
        If Err.number <> 0 Then
            Err.Clear
        End If
    End If
End Function

Public Function LogPathValid() As Boolean
On Error Resume Next
    CheckReady
    If lPathChecked = False Then
        If Not DirectoryExists(LogFolderPath) Then
            CreateDirectory LogFolderPath
        End If
        lPathValid = DirectoryExists(LogFolderPath)
        lPathChecked = True
        LogPathValid = lPathValid
    Else
        LogPathValid = lPathValid
    End If
End Function

Private Sub Class_Initialize()
    CheckReady
    If LogPathValid Then
        If DebugMode Then
            lLogLevel = llTrace
        Else
            lLogLevel = LogLevelEnum.[_Default]
        End If
        lLogLevelFriendly = LogLevelString
        
    Else
        ftBeep btError
        'MsgBox_FT "Unable to Configure Log Folder at: " & LogFolderPath & vbNewLine & "This session will not Log events", vbExclamation + vbOKOnly, "OOPS"
        If IsDEV Then
            Debug.Print "Error Initializing Log File"
            Stop
        End If
        DoEvents
    End If
End Sub

Private Sub Class_Terminate()
    If lOpened Then
        CloseLog
    End If
End Sub

Public Function ToggleLogMode()
    ToggleLogLevel
    MsgBox_FT "Log Level Set To: " & FriendlyDebugMode, vbOKOnly + vbInformation, "LOG LEVEL"
End Function

Public Function ViewCurrentLog()
    Dim logName As String
    logName = CurrentLogFileFullName
    If StringsMatch(logName, ".log", smEndWithStr) Then
        LogInfo "Flushing Pending Writes ...", True
        LogInfo "Flushing Pending Writes ...", True
        LogInfo "Flushing Pending Writes ...", True
        ViewFormattedLog FileNameFromFullPath(logName)
    End If
End Function

Public Function ViewFormattedLog(Optional logName As Variant)
#If privateVersion Then
    On Error GoTo E:
        Dim failed As Boolean
        If pbLog.LogFilesExist(False) = False Then
            MsgBox_FT "There aren't any log files. Log files are created ANY time an error or warning exists, or you tracing enabled. (Turn on Tracing by clicking the TOGGLE DEBUG LOGGIN button, or by double-clicking the 'Tracing/Logging'area on the dashoard.", vbOKOnly + vbInformation, "COOL BEANS"
            Exit Function
        End If
        Dim pkg As Package
        If IsMissing(logName) Then
            Set pkg = BuildPackage(LogInstance, "ViewFormattedLog", VbMethod)
            Set pkg.targetSheet = wsDebug
            Wom.picklist.Initialize "Select Log File", "Select log file to format and view", pbLog.LogFiles, plSingle, pkg
            Exit Function
        End If
        
        Dim fullPath As String
        fullPath = PathCombine(False, LogFolderPath, logName)
        Dim logWB As Workbook, logWS As Worksheet
        Set logWB = Workbooks.Open(fullPath, ReadOnly:=True, AddToMru:=False)
        Dim logARR As Variant, ai As ArrInformation
        
        Dim rowCnt As Long
        rowCnt = LastPopulatedRow(logWB.Worksheets(1), 1)
        
        logARR = ArrRange(logWB.Worksheets(1).usedRange, aoNone)
        ai = ArrayInfo(logARR)
        logWB.Close SaveChanges:=False
        DoEvents
        Set logWB = Nothing
        Set logWB = Workbooks.Add
        Dim newArr As Variant
        ReDim newArr(1 To UBound(logARR) + 1, 1 To 5)
        Dim idx As Long
        Dim rw As String, rwarr As Variant, offsetIdx As Long
        newArr(1, 1) = "TYPE"
        newArr(1, 2) = "LogDate"
        newArr(1, 3) = "LogTime"
        newArr(1, 4) = "SysStates"
        newArr(1, 5) = "LogMessage"
        Dim lbRow As Long
        For idx = ai.LBound_first To ai.Ubound_first
            offsetIdx = idx + 1
            rw = CStr(logARR(idx, 1))
            rwarr = Split(rw, ",", compare:=vbTextCompare)
            lbRow = LBound(rwarr)
            If UBound(rwarr) - LBound(rwarr) + 1 = 5 Then
                newArr(offsetIdx, 1) = rwarr(lbRow)
                newArr(offsetIdx, 2) = rwarr(lbRow + 1)
                newArr(offsetIdx, 3) = rwarr(lbRow + 2)
                newArr(offsetIdx, 4) = rwarr(lbRow + 3)
                newArr(offsetIdx, 5) = rwarr(lbRow + 4)
                
            Else
                newArr(offsetIdx, 1) = rw
            End If
        Next idx
        Dim tarRng As Range
        Set tarRng = logWB.Worksheets(1).Range("A5:E5")
        Set tarRng = tarRng.Resize(rowSize:=UBound(newArr))
        tarRng.Value = newArr
        Dim lo As ListObject
        Set logWS = logWB.Worksheets(1)
        Set lo = logWS.ListObjects.Add(SourceType:=xlSrcRange, Source:=tarRng, XlListObjectHasHeaders:=xlYes)
        logWS.Name = "FinToolLog"
        lo.Name = "tblFTLog"
        lo.Range.EntireColumn.AutoFit
        logWB.Worksheets(1).Range("A1").Value = "Source Path: " & fullPath
        
        DoEvents
        logWB.Activate
        logWB.Windows(1).DisplayGridlines = False
        
        Set Wom.ActivateExtWorkbook = logWB
        Set logWB = Nothing
        
        
Finalize:
        On Error Resume Next
        
        Exit Function
E:
        failed = True
        ErrorCheck
        Resume Finalize:
#End If

End Function

Private Function SubmitLogFile(logName As String)
#If privateVersion Then
    
    BusyWait "Fetching log file info: " & PathCombine(False, pbLog.LogFolderPath, logName)
    
    Dim fullPath As String
    fullPath = PathCombine(False, LogFolderPath, logName)
    Dim logWB As Workbook, logWS As Worksheet
    Set logWB = Workbooks.Open(fullPath, ReadOnly:=True, AddToMru:=False)
    Dim logARR As Variant, ai As ArrInformation
    
    Dim rowCnt As Long
    rowCnt = LastPopulatedRow(logWB.Worksheets(1), 1)
    
    logARR = ArrRange(logWB.Worksheets(1).usedRange, aoNone)
    ai = ArrayInfo(logARR)
    logWB.Close SaveChanges:=False
    DoEvents
    BusyWait "Formatting as new file for upload: " & fullPath
    Set logWB = Nothing
    Set logWB = Workbooks.Add
    Dim newArr As Variant
    ReDim newArr(1 To UBound(logARR) + 1, 1 To 5)
    Dim idx As Long
    Dim rw As String, rwarr As Variant, offsetIdx As Long
    newArr(1, 1) = "TYPE"
    newArr(1, 2) = "LogDate"
    newArr(1, 3) = "LogTime"
    newArr(1, 4) = "SysStates"
    newArr(1, 5) = "LogMessage"
    Dim lbRow As Long
    For idx = ai.LBound_first To ai.Ubound_first
        offsetIdx = idx + 1
        rw = CStr(logARR(idx, 1))
        rwarr = Split(rw, ",", compare:=vbTextCompare)
        lbRow = LBound(rwarr)
        If UBound(rwarr) - LBound(rwarr) + 1 = 5 Then
            newArr(offsetIdx, 1) = rwarr(lbRow)
            newArr(offsetIdx, 2) = rwarr(lbRow + 1)
            newArr(offsetIdx, 3) = rwarr(lbRow + 2)
            newArr(offsetIdx, 4) = rwarr(lbRow + 3)
            newArr(offsetIdx, 5) = rwarr(lbRow + 4)
            
        Else
            newArr(offsetIdx, 1) = rw
        End If
    Next idx
    Dim tarRng As Range
    Set tarRng = logWB.Worksheets(1).Range("A5:E5")
    Set tarRng = tarRng.Resize(rowSize:=UBound(newArr))
    tarRng.Value = newArr
    Dim lo As ListObject
    Set logWS = logWB.Worksheets(1)
    Set lo = logWS.ListObjects.Add(SourceType:=xlSrcRange, Source:=tarRng, XlListObjectHasHeaders:=xlYes)
    logWS.Name = "FinToolLog"
    lo.Name = "tblFTLog"
    lo.Range.EntireColumn.AutoFit
    logWB.Worksheets(1).Range("A1").Value = "Source Path: " & fullPath
    
    DoEvents
    logWB.Windows(1).DisplayGridlines = False
        
    Dim fName As String
    fName = FileNameFromFullPath(fullPath)
    fName = Replace(fName, ".log", ".xlsx")
    fName = Replace(fName, ".xlsx", Format(Now, "_HHMMSS") & ".xlsx")
    fName = PathCombine(False, LogFileUploadPath, fName)
    fName = SimpleURLEncode(fName)
    BusyWait "Uploading to: " & fName
    logWB.SaveAs fName, AddToMru:=False
    DoEvents
    logWB.Close SaveChanges:=True
    Set logWB = Nothing
    DoEvents
#End If
End Function

Public Function UploadLogs(Optional LogFiles As Variant)
#If privateVersion Then
On Error GoTo E:
    
    Dim endingScreen As Worksheet
    Set endingScreen = ThisWorkbook.ActiveSheet

    Dim failed As Boolean
    If pbLog.LogFilesExist(False) = False Then
        Exit Function
    End If
    Dim pkg As Package
    If IsMissing(LogFiles) Then
        Set pkg = BuildPackage(LogInstance, "UploadLogs", VbMethod)
        Set pkg.targetSheet = wsDebug
        Wom.picklist.Initialize "Select Log File(s) For Upload", "Select 1 or more log files to upload", pbLog.LogFiles, plMultiple_MinimumOne, pkg
        Exit Function
    End If
    
    DoEvents
    ShowBusy endingScreen.Name, "* UPLOADING LOG FILES *"
    CheckReady
    
    If TypeName(LogFiles) = "String" Then
        SubmitLogFile CStr(LogFiles)
    Else
        Dim fileARR As Variant
        fileARR = ArrArray(LogFiles, aoNone)
        Dim lfAI As ArrInformation
        lfAI = ArrayInfo(fileARR)
        Dim idx As Long
        For idx = lfAI.LBound_first To lfAI.Ubound_first
            SubmitLogFile CStr(fileARR(idx, 1))
        Next idx
    End If

Finalize:
    On Error Resume Next
    If PackageMgr.PackageIsExecuting Then PackageMgr.RedirectExecutingTarget endingScreen
    CloseBusy forceFinalSheet:=endingScreen
    
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:

#End If
End Function
