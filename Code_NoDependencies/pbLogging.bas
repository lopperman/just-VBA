Attribute VB_Name = "pbLogging"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  An example on one way to do logging
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbLog.bas
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1

Private Enum strMatchEnum
    smEqual = 0
    smNotEqualTo = 1
    smContains = 2
    smStartsWithStr = 3
    smEndWithStr = 4
End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   'PBCOMMON_LOG'
'       When TRUE, CALLING 'pbCommonUtil.Log'
'       will add log message to:
'       - DirectoryName:
'         Application.DefaultFilePath/[LOG_DIR]
'       - LogFileName=workbookname & '_LOG_'  & [YYYYMMDD]
    Public Const PBCOMMON_LOG As Boolean = True
    Public Const LOG_DIR As String _
        = "PBCOMMONLOG"
'   'pbLogFileNumber' is used to store FreeFile when
'       keeping logFile open. If you're expecting to write more
'       than a few log messages, performance is significatnly
'       increased when file is kept open
    Private pbLogFileNumber
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'                       TESTS
Public Function TestLog()
    '   (Not Kept Open) - 1000 log messages in 6.082031 seconds
    '   (Kept Open) - 1000 log messages in  0.080078 seconds
    ' --- '
    Dim iMax As Long, i As Long
    iMax = 1000
    Dim timerStart, timerEnd
    timerStart = CDbl(Timer)
    'make sure log file is close
    pbLogClose
    For i = 1 To iMax
        pbLog Format(i, "00000") & " - " & Now & " - this is a test abcdefghijklmnopqrstuvwxyz 1234567890"
    Next i
    timerEnd = CDbl(Timer)
    Debug.Print "(Not Kept Open) - " & iMax & " log messages in " & Round(timerEnd - timerStart, 6) & " seconds"
    
    timerStart = CDbl(Timer)
    For i = 1 To iMax
        pbLog Format(i, "00000") & " - " & Now & " - this is a test abcdefghijklmnopqrstuvwxyz 1234567890", closeLog:=False
    Next i
    pbLogClose
    timerEnd = CDbl(Timer)
    Debug.Print "(Kept Open) - " & iMax & " log messages in  " & Round(timerEnd - timerStart, 6) & " seconds"

End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'
'       PB COMMON LOGGING
'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'    Public Const PBCOMMON_LOG As Boolean = True
'    Public Const LOG_DIR As String _
'        = "PBCOMMONLOG"
'    Public Const LOG_AUTO_PURGE As Boolean _
'        = True
'    Public Const LOG_MAXAGE_DAYS As Long _
'        = 30

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   returns error object if not using pbcommong_log
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function pbLOGPath(Optional wkbk As Workbook) As Variant
    If PBCOMMON_LOG = False Then
        pbLOGPath = CVErr(1028)
    Else
        Dim logFileName As String
        If wkbk Is Nothing Then
            logFileName = FileNameWithoutExtension(ThisWorkbook.Name)
        Else
            logFileName = FileNameWithoutExtension(wkbk.Name)
        End If
        pbLOGPath = PathCombine(False, Application.DefaultFilePath _
            , LOG_DIR _
            , ConcatWithDelim("_", logFileName, "LOG", Format(Date, "YYYYMMDD") & ".log"))
    End If
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'    Configure next 'FreeFile' for appending
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function pbLogOpen()
    If PBCOMMON_LOG = False Then Exit Function
    'if the log file is open and this method is called, will close and then re-open
    If Val(pbLogFileNumber) > 0 Then
        Close #pbLogFileNumber
    End If
    pbLogFileNumber = FreeFile
    Dim logPath As String
    logPath = CStr(pbLOGPath)
    On Error Resume Next
    Open pbLOGPath For Append As #pbLogFileNumber
    If Err.Number = 75 Then
        Err.Clear
        On Error GoTo 0
        CreateDirectory PathCombine(True, Application.DefaultFilePath, LOG_DIR)
        Open pbLOGPath For Append As #pbLogFileNumber
    End If
End Function

Public Function pbLogClose()
    If PBCOMMON_LOG = False Then Exit Function
    If Val(pbLogFileNumber) > 0 Then
        Close #pbLogFileNumber
        DoEvents
        pbLogFileNumber = Empty
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Log [msg] to[pbLogPath()]
'   if there is an open fileNumber for log, that will be used.
'   to open file for append, call 'pbLogOpen'
'   to close file call 'pbLogClose'
'   if [closeLog] = false, file will not close file after write
'       ** YOU ARE RESPONSIBLE FOR CLOSING OPEN FILES
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function pbLog(msg, Optional prependTimeStamp As Boolean = True, Optional closeLog As Boolean = True)
    If PBCOMMON_LOG = False Then Exit Function
    If Val(pbLogFileNumber) = 0 Then
        pbLogOpen
    End If
    If prependTimeStamp Then
        msg = ConcatWithDelim(NowWithMS, msg)
    End If
    Write #pbLogFileNumber, msg
    If closeLog Then
        pbLogClose
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Methods from pbCommon, made private here to remove
'   dependencies
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

Private Function FileNameWithoutExtension(ByVal fileName As String) As String
    If InStrRev(fileName, ".") > 0 Then
        Dim tmpExt As String
        tmpExt = Mid(fileName, InStrRev(fileName, "."))
        If Len(tmpExt) >= 2 Then
            fileName = Replace(fileName, tmpExt, vbNullString)
        End If
    End If
    FileNameWithoutExtension = fileName
End Function
Private Function Concat(ParamArray items() As Variant) As String
    Concat = Join(items, "")
End Function
Private Function ConcatWithDelim(ByVal delimeter As String, ParamArray items() As Variant) As String
    ConcatWithDelim = Join(items, delimeter)
End Function

Private Function StringsMatch( _
    ByVal checkString As Variant, ByVal _
    validString As Variant, _
    Optional smEnum As strMatchEnum = strMatchEnum.smEqual, _
    Optional compMethod As VbCompareMethod = vbTextCompare) As Boolean
    
'       IF NEEDED, PUT THIS ENUM AT TOP OF A STANDARD MODULE
        'Public Enum strMatchEnum
        '    smEqual = 0
        '    smNotEqualTo = 1
        '    smContains = 2
        '    smStartsWithStr = 3
        '    smEndWithStr = 4
        'End Enum
        
    Dim str1, str2
        
    str1 = CStr(checkString)
    str2 = CStr(validString)
    Select Case smEnum
        Case strMatchEnum.smEqual
            StringsMatch = StrComp(str1, str2, compMethod) = 0
        Case strMatchEnum.smNotEqualTo
            StringsMatch = StrComp(str1, str2, compMethod) <> 0
        Case strMatchEnum.smContains
            StringsMatch = InStr(1, str1, str2, compMethod) > 0
        Case strMatchEnum.smStartsWithStr
            StringsMatch = InStr(1, str1, str2, compMethod) = 1
        Case strMatchEnum.smEndWithStr
            If Len(str2) > Len(str1) Then
                StringsMatch = False
            Else
                StringsMatch = InStr(Len(str1) - Len(str2) + 1, str1, str2, compMethod) = Len(str1) - Len(str2) + 1
            End If
    End Select
End Function
Private Function PathCombine(includeEndSeparator As Boolean, ParamArray vals() As Variant) As String
    
    Dim tDelim As String, isHTTP As Boolean
    Dim i As Long
    Dim retV As String
    Dim dblPS As String
    Dim wrongPS As String
    For i = LBound(vals) To UBound(vals)
        If LCase(vals(i)) Like "*http*" Then
            isHTTP = True
            tDelim = "/"
            wrongPS = "\"
        End If
    Next i
    If Not isHTTP Then
        tDelim = Application.PathSeparator
        If InStr(1, "/", Application.PathSeparator) > 0 Then
            wrongPS = "\"
        Else
            wrongPS = "/"
        End If
    End If
    dblPS = tDelim & tDelim
    For i = LBound(vals) To UBound(vals)
        If i = LBound(vals) Then
            retV = CStr(vals(i))
            If Len(retV) = 0 Then retV = tDelim
        Else
            If Mid(retV, Len(retV)) = tDelim Then
                retV = retV & vals(i)
            Else
                retV = retV & tDelim & vals(i)
            End If
        End If
    Next i
    retV = Replace(retV, wrongPS, tDelim)
    If isHTTP Then
        retV = Replace(retV, "://", ":::")
        Do While InStr(1, retV, dblPS) > 0
            retV = Replace(retV, dblPS, tDelim)
        Loop
        retV = Replace(retV, ":::", "://")
    Else
        Do While InStr(1, retV, dblPS) > 0
            retV = Replace(retV, dblPS, tDelim)
        Loop
    End If
    If includeEndSeparator Then
        If Not Mid(retV, Len(retV)) = tDelim Then
            retV = retV & Application.PathSeparator
        End If
    Else
        'Remove it if it's there
        If Mid(retV, Len(retV)) = Application.PathSeparator Then
            retV = Mid(retV, 1, Len(retV) - 1)
        End If
    End If
    PathCombine = retV

End Function
Private Function CreateDirectory(fullPath As String) As Boolean
' IF 'fullPath' is not a valid directory but the '1 level back' IS a valid directory, then the last directory in 'fullPath' will be created
' Example: CreateDirectory("/Users/paul/Library/Containers/com.microsoft.Excel/Data/Documents/FinToolTemp/Logs")
    'If the 'FinToolTemp' directory exists, the Logs will be created if it is not already there.
'   Primary reason for not creating multiple directories in the path is issues with both PC and Mac for File System changes.
    
    Dim retV As Boolean

    If DirectoryExists(fullPath) Then
        retV = True
    Else
        Dim lastDirName As String, pathArr As Variant, checkFldrName As String
        fullPath = PathCombine(False, fullPath)
        If InStrRev(fullPath, Application.PathSeparator, Compare:=vbTextCompare) > InStr(1, fullPath, Application.PathSeparator, vbTextCompare) Then
            lastDirName = Left(fullPath, InStrRev(fullPath, Application.PathSeparator, Compare:=vbTextCompare) - 1)
            If DirectoryExists(lastDirName) Then
                On Error Resume Next
                MkDir fullPath
                If Err.Number = 0 Then
                    retV = DirectoryExists(fullPath)
                End If
            End If
        End If
    End If
    CreateDirectory = retV
    If Err.Number <> 0 Then Err.Clear
End Function
Private Function DirectoryExists(dirPath As String) As Boolean
On Error Resume Next
    Dim retV As Boolean
    Dim lastDirName As String, pathArr As Variant, checkFldrName As String
    dirPath = PathCombine(False, dirPath)

    If InStr(dirPath, Application.PathSeparator) > 0 Then
        pathArr = Split(dirPath, Application.PathSeparator)
        checkFldrName = CStr(pathArr(UBound(pathArr)))
        retV = StrComp(Dir(dirPath & "*", vbDirectory), LCase(checkFldrName), vbTextCompare) = 0
        If Err.Number <> 0 Then
            Debug.Print "DirectoryExists: Err Getting Path: " & dirPath & ", " & Err.Number & " - " & Err.Description
        End If
    End If
    DirectoryExists = retV
    If Err.Number <> 0 Then Err.Clear
End Function
Private Function NowWithMS() As String
    NowWithMS = Format(Now, "yyyymmdd hh:mm:ss") & Right(Format(Timer, "0.000"), 4)
End Function

