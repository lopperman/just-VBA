Attribute VB_Name = "pbMiscUtil"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' pbMiscUtil v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' Misc utilities and helpers that can be part of external shared library
'
' @module pbMiscUtil
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' Enums and Types should be in the module where appropriate
Option Explicit
Option Compare Text
Option Base 1

'   DEFAULT NAME (NOT PATH) OF TOP LEVEL TEMP DIRETORY FOR CURRENT APP
Public Const TEMP_DIRECTORY_NAME As String = "FinToolTemp"
Public Const LOG_DIRECTORY_NAME As String = "Logs"

Private l_OperatingState As ftOperatingState
Private l_listObjDict As Dictionary

Public Function pbProtectSheet(ws As Worksheet) As Boolean
    'Replace with you own implementation
    'pbProtectSheet = True
    
    pbProtectSheet = ProtectSHT(ws)
End Function
Public Function pbUnprotectSheet(ws As Worksheet) As Boolean
    'Replace with you own implementation
    'pbpbUnprotectSheet = True
    
    pbUnprotectSheet = UnprotectSHT(ws)
End Function


Public Function IsFTClosing() As Variant
    IsFTClosing = (l_OperatingState = ftClosing)
End Function
Public Function IsFTOpening() As Variant
    IsFTOpening = (l_OperatingState = ftOpening)
End Function

Public Property Get ftState() As ftOperatingState
    ftState = l_OperatingState
    wsOpenClose.Calculate
End Property
Public Property Let ftState(ftsVal As ftOperatingState)
    l_OperatingState = ftsVal
End Property


Public Property Get DefaultTempTblPrefixes() As Variant()
    'DEFAULT PREFIXES TO INDICATE A TABLE/LISTOBJECT IS TEMPORARY
    'USED BY THE 'WT' FUNCTION - if optional parameter is excluded, WHICH STORES LIST OBJECT REFERENCES IN A DICTIONARY
    DefaultTempTblPrefixes = Array("tmp", "temp", "table")
End Property

Public Function TempDirName(Optional dirName As String = vbNullString) As String
    TempDirName = IIf(Not dirName = vbNullString, dirName, TEMP_DIRECTORY_NAME)
End Function


Public Function IsMac() As Boolean
'   Returns True If Mac OS
    #If Mac Then
        IsMac = True
    #End If
End Function

Public Function Max2(Val1, Val2)
' REPLACE WORKSHEET 'MAX' WITH THIS (MUCH BETTER PERFORMANCE FROM VBA)
    If Val1 > Val2 Then
        Max2 = Val1
    Else
        Max2 = Val2
    End If
End Function
Public Function Min2(Val1, Val2)
' REPLACE WORKSHEET 'MAX' WITH THIS (MUCH BETTER PERFORMANCE FROM VBA)
    If Val1 > Val2 Then
        Min2 = Val2
    Else
        Min2 = Val1
    End If
End Function

Public Function CallAppRun(wbName As String, procName As String, Optional raiseErrorOnFail As Boolean = False)
'   Execute a 'Public Workbook Sub or Function in Workbook 'wbName'
    On Error GoTo E:
    wbName = CleanSingleTicks(wbName)
    Application.Run ("'" & wbName & "'!'" & procName & "'")
    Exit Function
E:
    Beep
    If Not raiseErrorOnFail Then
        Err.Clear
        On Error GoTo 0
    Else
        Err.Raise Err.Number, Err.Description
    End If
End Function

Public Function UnsetWTDict() As Boolean
'   Clears Dictionary used for storing Global Reference to ListObject in Workbook
'   See 'wt' Function for more info
    If Not l_listObjDict Is Nothing Then
        l_listObjDict.RemoveAll
    End If
    Set l_listObjDict = Nothing
    UnsetWTDict = l_listObjDict Is Nothing
End Function

Public Function wt(listObjectName As String, ParamArray tempListObjPrefixes() As Variant) As ListObject
'   Return object reference to ListObject in 'ThisWorkbook' called [listObjectName]
'   This function exists to eliminate problem with getting a ListObject using the 'Range([list object name])
'       where the incorrect List Object could be returned if the ActiveWorkbook containst a list object
'       with the same name, and is not the intended ListObject
'  If temporary list object mayexists, include the prefixes (e.g. "tmp","temp") to identify and not add to dictionary
On Error GoTo E:
    
    Dim i As Long, t As ListObject, ignoreArr As Variant, ignoreAI As ArrInformation, ignoreIdx As Long, ignore As Boolean
    Dim sw As StopWatch
    
    'Force array to 2D
    If IsMissing(tempListObjPrefixes) Then
        ignoreArr = ArrArray(DefaultTempTblPrefixes, aoNone)
    Else
        ignoreArr = ArrParams(tempListObjPrefixes)
    End If
    ignoreAI = ArrayInfo(ignoreArr)
    
    If l_listObjDict Is Nothing Then
    '   If th Dictionary is Empty, we're opening file, givea small breather to the app
        DoEvents
        Set sw = New StopWatch
        sw.StartTimer
        Set l_listObjDict = New Dictionary
    
        For i = 1 To ThisWorkbook.Worksheets.Count
            For Each t In ThisWorkbook.Worksheets(i).ListObjects
                ignore = False
                If ignoreAI.Dimensions > 0 Then
                    For ignoreIdx = ignoreAI.LBound_first To ignoreAI.Ubound_first
                        If InStr(1, t.Name, CStr(ignoreArr(ignoreIdx, 1)), vbTextCompare) > 0 Then
                            ignore = True
                            Exit For
                        End If
                    Next ignoreIdx
                End If
                If Not ignore Then
                    Set l_listObjDict(t.Name) = t
                End If
           Next t
        Next i
        sw.StopTimer
        DoEvents
        If IsDEV And DebugMode Then
            Debug.Print "mdlGlobal.wt - Load All listObjects: " & sw.ResultAsTime
        End If
        Set sw = Nothing
    End If
    
    'this covers the temporary listobject which may not always be available
    If Not l_listObjDict.Exists(listObjectName) Then
        Dim tWS As Worksheet, tLO As ListObject, tIDX As Long, tLOIDX
        For tIDX = 1 To ThisWorkbook.Worksheets.Count
            If ThisWorkbook.Worksheets(tIDX).ListObjects.Count > 0 Then
                For tLOIDX = 1 To ThisWorkbook.Worksheets(tIDX).ListObjects.Count
                    If ThisWorkbook.Worksheets(tIDX).ListObjects(tLOIDX).Name = listObjectName Then
                        'DON'T ADD any tmp tables
                        Set wt = ThisWorkbook.Worksheets(tIDX).ListObjects(tLOIDX)
                        GoTo Finalize:
                    End If
                Next tLOIDX
            End If
        Next tIDX
    End If
    
Finalize:
    On Error Resume Next
    
    If l_listObjDict.Exists(listObjectName) Then
        Set wt = l_listObjDict(listObjectName)
    Else
        If IsDEV Then
            If wt Is Nothing Then
                Beep
            End If
           ' Beep
            'Stop
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    Beeper
    DebugPrint "Error getting list object " & listObjectName
    Err.Clear
End Function

' ~~~~~~~~~~   CLEAN SINGLE TICKS ~~~~~~~~~~'
Public Function CleanSingleTicks(wbName As String) As String
    Dim retV As String
    If InStr(wbName, "'") > 0 And InStr(wbName, "''") = 0 Then
        retV = Replace(wbName, "'", "''")
    Else
        retV = wbName
    End If
    CleanSingleTicks = retV
End Function
' ~~~~~~~~~~   FLAG ENUM COMPARE ~~~~~~~~~~'
Public Function FlagEnumCompare(theEnum As Variant, enumMember As Variant) As Boolean
'Use to check Bitwise enums
    FlagEnumCompare = CBool(theEnum And enumMember) = True
End Function

' ~~~~~~~~~~   INPUT BOX   ~~~~~~~~~~'
Public Function InputBox_FT(prompt As String, Optional title As String = "Financial Tool - Input Needed", Optional default As Variant, Optional inputType As ftInputBoxType) As Variant
    Beeper
    If inputType > 0 Then
        InputBox_FT = Application.InputBox(prompt, title:=title, default:=default, Type:=inputType)
    Else
        InputBox_FT = Application.InputBox(prompt, title:=title, default:=default)
    End If
    DoEvents
End Function
' ~~~~~~~~~~   MSG BOX   ~~~~~~~~~~'
Public Function MsgBox_FT(prompt As String, Optional buttons As VbMsgBoxStyle = vbOKOnly, Optional title As Variant) As Variant
    Dim evts As Boolean: evts = Events
    Dim screenUpd As Boolean: screenUpd = Application.ScreenUpdating
    EventsOff
    If Not ButtonOptionSet(buttons, vbSystemModal) Then buttons = buttons + vbSystemModal
    If Not ButtonOptionSet(buttons, vbMsgBoxSetForeground) Then buttons = buttons + vbMsgBoxSetForeground
    Beep
    If Not ThisWorkbook.ActiveSheet Is Application.ActiveSheet Then
        Application.ScreenUpdating = True
        ThisWorkbook.Activate
        DoEvents
        Application.ScreenUpdating = screenUpd
    End If
    MsgBox_FT = MsgBox(prompt, buttons, title)
    Events = evts
    DoEvents
End Function
' ~~~~~~~~~~   ASK YES NO ~~~~~~~~~~'
Public Function AskYesNo(msg As String, title As String, Optional defaultYES As Boolean = True) As Variant
    If IsMissing(title) Then
        title = "QUESTION"
    End If
    Beep
    If defaultYES Then
        AskYesNo = MsgBox_FT(msg, vbYesNo + vbQuestion, title)
    Else
        AskYesNo = MsgBox_FT(msg, vbYesNo + vbQuestion + vbDefaultButton2, title)
    End If
    DoEvents
End Function

' ~~~~~~~~~~   GET NEXT ID ~~~~~~~~~~'
Public Function GetNextID(table As ListObject, uniqueIdcolumnIdx As Long) As Long
'   Use to create next (Long) number for unique ROW id in a Range
On Error Resume Next
    Dim nextID As Long
    If table.listRows.Count > 0 Then
        nextID = Application.WorksheetFunction.Max(table.ListColumns(uniqueIdcolumnIdx).DataBodyRange)
    End If
    GetNextID = nextID + 1
    If Err.Number <> 0 Then Err.Clear
End Function


' ~~~~~~~~~~   ~~ ~~ ~~ ~~   ~~~~~~~~~~' ' ~~~~~~~~~~   ~~ ~~ ~~ ~~   ~~~~~~~~~~'
' ~~~~~~~~~~   FILE SYSTEM ~~~~~~~~~~' ' ~~~~~~~~~~   ~~ ~~ ~~ ~~   ~~~~~~~~~~'
' ~~~~~~~~~~   ~~ ~~ ~~ ~~   ~~~~~~~~~~' ' ~~~~~~~~~~   ~~ ~~ ~~ ~~   ~~~~~~~~~~'

' ~~~~~~~~~~   STARTING TEMP DIRECTORY FOR CURRENT APP ~~~~~~~~~~'
Public Property Get TempDirPath() As String
On Error Resume Next
    Dim tmpPath As String
    tmpPath = PathCombine(True, Application.DefaultFilePath, TEMP_DIRECTORY_NAME)
    If DirectoryExists(tmpPath) Then
        TempDirPath = tmpPath
    End If
    If Err.Number <> 0 Then Err.Clear
End Property

' ~~~~~~~~~~   CREATE THE ** LAST ** DIRECTORY IN 'fullPath' ~~~~~~~~~~'
Public Function CreateDirectory(fullPath As String) As Boolean
' IF 'fullPath' is not a valid directory but the '1 level back' IS a valid directory, then the last directory in 'fullPath' will be created
' Example: CreateDirectory("/Users/paul/Library/Containers/com.microsoft.Excel/Data/Documents/FinToolTemp/Logs")
    'If the 'FinToolTemp' directory exists, the Logs will be created if it is not already there.
'   Primary reason for not creating multiple directories in the path is issues with both PC and Mac for File System changes.
    
    DebugPrint ConcatWithDelim(", ", "pbMiscUtil.CreateDirectory", "CHECKING", fullPath)
    
    Dim retV As Boolean

    If DirectoryExists(fullPath) Then
        DebugPrint ConcatWithDelim(", ", "pbMiscUtil.CreateDirectory", fullPath, "aready exists")
        retV = True
    Else
        Dim lastDirName As String, pathArr As Variant, checkFldrName As String
        fullPath = PathCombine(False, fullPath)
        If InStrRev(fullPath, Application.PathSeparator, Compare:=vbTextCompare) > InStr(1, fullPath, Application.PathSeparator, vbTextCompare) Then
            lastDirName = left(fullPath, InStrRev(fullPath, Application.PathSeparator, Compare:=vbTextCompare) - 1)
            If DirectoryExists(lastDirName) Then
                On Error Resume Next
                DebugPrint ConcatWithDelim(", ", "pbMiscUtil.CreateDirectory", "Creating directory: ", fullPath)
                MkDir fullPath
                If Err.Number = 0 Then
                    DebugPrint ConcatWithDelim(", ", "pbMiscUtil.CreateDirectory", "Created: ", fullPath)
                
                    retV = DirectoryExists(fullPath)
                End If
            End If
        End If
    End If
    CreateDirectory = retV
    If Err.Number <> 0 Then Err.Clear
End Function

' ~~~~~~~~~~   Returns true if DIRECTORY path dirPath) Exists ~~~~~~~~~~'
Public Function DirectoryExists(dirPath As String) As Boolean
On Error Resume Next
    Dim retV As Boolean
    Dim lastDirName As String, pathArr As Variant, checkFldrName As String
    dirPath = PathCombine(False, dirPath)

    If InStr(dirPath, Application.PathSeparator) > 0 Then
        pathArr = Split(dirPath, Application.PathSeparator)
        checkFldrName = CStr(pathArr(UBound(pathArr)))
        retV = StrComp(Dir(dirPath & "*", vbDirectory), LCase(checkFldrName), vbTextCompare) = 0
        If Err.Number <> 0 Then
            DebugPrint "DirectoryExists: Err Getting Path: " & dirPath & ", " & Err.Number & " - " & Err.Description
        End If
    End If
    DirectoryExists = retV
    If Err.Number <> 0 Then Err.Clear
End Function

' ~~~~~~~~~~   Returns true if filePth Exists and is not a directory  ~~~~~~~~~~'
Public Function FileExists(filePth As String, Optional allowWildcardsForFile As Boolean = False) As Boolean
On Error Resume Next
    Dim retV As Boolean
    Dim lastDirName As String, pathArr As Variant, checkFlName As String
    filePth = PathCombine(False, filePth)

    If InStr(filePth, Application.PathSeparator) > 0 Then
        pathArr = Split(filePth, Application.PathSeparator)
        checkFlName = CStr(pathArr(UBound(pathArr)))
        Dim tmpReturnedFileName As String
        tmpReturnedFileName = Dir(filePth & "*", vbNormal)
        If allowWildcardsForFile = True And Len(tmpReturnedFileName) > 0 Then
            retV = True
        Else
            retV = StrComp(Dir(filePth & "*"), LCase(checkFlName), vbTextCompare) = 0
        End If
        If Err.Number <> 0 Then DebugPrint "DirectoryExists: Err Getting Path: " & filePth & ", " & Err.Number & " - " & Err.Description
    End If
    FileExists = retV
    If Err.Number <> 0 Then Err.Clear
End Function

' ~~~~~~~~~~   Create Valid File or Directory Path (for PC or Mac, local, or internet) from 1 or more arguments  ~~~~~~~~~~'
Public Function PathCombine(includeEndSeparator As Boolean, ParamArray vals() As Variant) As String
' COMBINE PATH AND/OR FILENAME SEGMENTS
' WORKS FOR MAC OR PC ('/' vs '\'), and for web url's
'
'   Debug.Print PathCombine(True, "/usr", "\\what", "/a//", "mess")
'      outputs:  /usr/what/a/mess/
'   Debug.Print PathCombine(False, "/usr", "\\what", "/a//", "mess", "word.docx/")
'      outputs: /usr/what/a/mess/word.docx
'   Debug.Print PathCombine(true,"https://www.google.com\badurl","gmail")
'       outputs:  https://www.google.com/badurl/gmail/
    
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

Public Function WorksheetExists(sName As String, Optional wbk As Workbook) As Boolean
On Error Resume Next
    If wbk Is Nothing Then
        Set wbk = ThisWorkbook
    End If
    Dim ws As Worksheet
    Set ws = wbk.Worksheets(sName)
    If Err.Number = 0 Then
        WorksheetExists = True
    End If
    Set ws = Nothing
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function WorkbookIsOpen(wkBkName As String, Optional verifyFullNameURLEncoded As String = vbNullString) As Boolean

    Dim i As Long, retV As Boolean
    For i = 1 To Application.Workbooks.Count
        If StrComp(LCase(wkBkName), LCase(Application.Workbooks(i).Name), vbTextCompare) = 0 Then
            If Len(verifyFullNameURLEncoded) > 0 Then
                If StrComp(LCase(verifyFullNameURLEncoded), LCase(Application.Workbooks(i).FullNameURLEncoded), vbTextCompare) = 0 Then
                    retV = True
                    Exit For
                End If
            Else
                retV = True
            End If
        End If
    Next i
    
    If retV = False And LCase(wkBkName) Like "*.xlam" Then
        'This covers Addins, which by default will not show up in regular enumeration of Workbooks
        '   Requires an Explicit 'Workbooks([addin workbook name])'
        On Error Resume Next
        Dim tmpWB As Workbook
        Set tmpWB = Workbooks(wkBkName)
        If Err.Number = 0 And Not tmpWB Is Nothing Then
            retV = True
        End If
        Set tmpWB = Nothing
    End If
    
    WorkbookIsOpen = retV
If Err.Number <> 0 Then Err.Clear
End Function

Public Function FirstMondayOfMonth(dtVal As Variant) As Variant
    Dim firstOfMonth As Variant, tMonday As Variant
    firstOfMonth = DateSerial(DatePart("yyyy", dtVal), DatePart("m", dtVal), 1)
    tMonday = GetMondayOfWeek(firstOfMonth)
    If DatePart("m", firstOfMonth) = DatePart("m", tMonday) Then
        FirstMondayOfMonth = tMonday
    Else
        FirstMondayOfMonth = DateAdd("d", 7, tMonday)
    End If
End Function

Public Function GetSundayOfWeek(inputDate As Variant) As Date
    Dim processDt As Variant
    If TypeName(inputDate) = "String" Then
        processDt = DateValue(inputDate)
    Else
        processDt = inputDate
    End If

    If Weekday(processDt, vbMonday) = 7 Then
        GetSundayOfWeek = processDt
    Else
        GetSundayOfWeek = DateAddDays(GetMondayOfWeek(processDt), 6)
    End If
End Function

Public Function GetMondayOfWeek(inputDate As Variant) As Date

    Dim processDt As Variant
    If TypeName(inputDate) = "String" Then
        processDt = DateValue(inputDate)
    Else
        processDt = inputDate
    End If

    If Weekday(processDt, vbMonday) = 1 Then
        GetMondayOfWeek = processDt
    Else
        GetMondayOfWeek = DateAddDays(processDt, 1 - Weekday(processDt, vbMonday))
    End If
    
End Function
Public Function DateAddDays(dt As Variant, addDays As Double) As Date
On Error Resume Next
    Dim newDt As Double
    newDt = CDbl(dt) + addDays
    DateAddDays = CDate(newDt)
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function FullPathExcludingFileName(fullFileName As String) As String
On Error Resume Next
    Dim tPath As String, tFileName As String, fNameStarts As Long
    tFileName = FileNameFromFullPath(fullFileName)
    fNameStarts = InStr(fullFileName, tFileName)
    tPath = Mid(fullFileName, 1, fNameStarts - 1)
    FullPathExcludingFileName = tPath
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function FileNameFromFullPath(fullFileName As String) As String
On Error Resume Next
    Dim sepChar As String
    sepChar = Application.PathSeparator
    If LCase(fullFileName) Like "*http*" Then
        sepChar = "/"
    End If
    Dim lastSep As Long: lastSep = Strings.InStrRev(fullFileName, sepChar)
    Dim shortFName As String:  shortFName = Mid(fullFileName, lastSep + 1)
    FileNameFromFullPath = shortFName
    If Err.Number <> 0 Then Err.Clear
End Function
Public Function ChooseFolder(choosePrompt As String) As String
'   Get User-Selected Directory name (MAC and PC Supported)
On Error Resume Next
    Beep
    Dim retV As Variant

    #If Mac Then
        retV = MacScript("choose folder with prompt """ & choosePrompt & """ as string")
        If Len(retV) > 0 Then
            retV = MacScript("POSIX path of """ & retV & """")
        End If
    #Else
        Dim fldr As FileDialog
        Dim sItem As String
        Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
        With fldr
            .title = choosePrompt
            .AllowMultiSelect = False
            .InitialFileName = Application.DefaultFilePath
            If .Show <> -1 Then GoTo NextCode
            retV = .SelectedItems(1)
        End With
NextCode:
        Set fldr = Nothing
    #End If
    
    ChooseFolder = retV
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function ChooseFile(choosePrompt As String, Optional fileExt As String = vbNullString) As String
'   Get User-Select File Name (MAC and PC Supported)
On Error Resume Next
    Beep
    Dim retV As Variant

    #If Mac Then
        retV = MacScript("choose file with prompt """ & choosePrompt & """ as string")
        If Len(retV) > 0 Then
            retV = MacScript("POSIX path of """ & retV & """")
        End If
    #Else
        Dim fldr As FileDialog
        Dim sItem As String
        Set fldr = Application.FileDialog(msoFileDialogFilePicker)
        With fldr
            .title = choosePrompt
            If Not fileExt = vbNullString Then
                .Filters.Clear
                .Filters.add "Files", fileExt & "?", 1
            End If
            .AllowMultiSelect = False
            If .Show <> -1 Then GoTo NextCode
            retV = .SelectedItems(1)
        End With
NextCode:
        Set fldr = Nothing
    #End If
    
    ChooseFile = retV
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function WaitWithDoEvents(waitSeconds As Long)
'WAIT FOR N SECONDS WHILE ALLOWING OTHER EXCEL EVENT TO PROCESS
'PURPOSE IS TO ENABLE ENOUGHT TIME TO PASS FOR APPLICATION ONTIME TO TAKE HOLD
    Dim stTimer As Single
    stTimer = Timer
    DebugPrint "Start Wait: " & stTimer
    Do While Timer - stTimer < waitSeconds
        DoEvents
    Loop
    DebugPrint "End Wait: Waited For: " & Math.Round((Timer - stTimer), 3) & " seconds"
    
End Function

Public Function IsWorkbookOpen(wName As String, Optional checkCodeName As String = vbNullString) As Boolean

    Dim retV As Boolean

    Dim wIDX As Long
    For wIDX = 1 To Application.Workbooks.Count
        If StrComp(LCase(wName), LCase(Application.Workbooks(wIDX).Name), vbTextCompare) = 0 Then
            If checkCodeName <> vbNullString Then
                If StrComp(LCase(checkCodeName), LCase(Workbooks(wIDX).CodeName), vbTextCompare) = 0 Then
                    retV = True
                Else
                    retV = False
                End If
            Else
                retV = True
            End If
        End If
        If retV Then
            Exit For
        End If
    Next wIDX
    
    IsWorkbookOpen = retV

End Function

Public Function CallOnTime_TwoArg(wbName As String, procName As String, argVal1 As String, argVal2 As String, Optional secondsDelay As Long = 0)
    'FT HELPER NEEDS TO BE UPDATED AND TESTED BEFORE ALLOWING THE PARAMETER TO GO THROUGH
    wbName = CleanSingleTicks(wbName)
    argVal1 = CleanSingleTicks(argVal1)
    argVal2 = CleanSingleTicks(argVal2)
    
    Dim litDQ As String
    litDQ = """"
    
    Application.OnTime EarliestTime:=GetTimeDelay(secondsDelay), Procedure:="'" & wbName & "'!'" & procName & " " & litDQ & argVal1 & litDQ & "," & litDQ & argVal2 & litDQ & "'"
End Function

Public Function WrapExternalCall(wbName As String, procName As String, argVal As Variant) As String
    If TypeName(argVal) = "String" Then
        argVal = CleanSingleTicks(CStr(argVal))
        WrapExternalCall = "'" & wbName & "'!'" & procName & " " & """" & argVal & """'"
    Else
        WrapExternalCall = "'" & wbName & "'!'" & procName & " " & "" & argVal & "'"
    End If
End Function

Public Function CallOnTime_OneArg(wbName As String, procName As String, argVal As Variant, Optional secondsDelay As Long = 0)
    If IsDEV Then
        Beep
        Debug.Print " ***** DEV ***** See if OnTime can work properly as Application.Run"
    End If
    wbName = CleanSingleTicks(wbName)
    If TypeName(argVal) = "String" Then
        argVal = Replace(argVal, "'", "''")
        Application.OnTime EarliestTime:=GetTimeDelay(secondsDelay), Procedure:="'" & wbName & "'!'" & procName & " " & """" & argVal & """'"
    Else
        Application.OnTime EarliestTime:=GetTimeDelay(secondsDelay), Procedure:="'" & wbName & "'!'" & procName & " " & "" & argVal & "'"
    End If
End Function

Public Function GetTimeDelay(Optional secondsDelay As Long = 0) As Date
    If secondsDelay > 59 Then secondsDelay = 59
    If secondsDelay < 0 Then secondsDelay = 0
    GetTimeDelay = Now + TimeValue("00:00:" & Format(secondsDelay, "00"))
End Function

Public Function CallOnTime(wbName As String, procName As String, Optional secondsDelay As Long = 0)
    wbName = CleanSingleTicks(wbName)
    Dim tProc As String
    tProc = "'" & wbName & "'!'" & procName & "'"
    Application.OnTime EarliestTime:=GetTimeDelay(secondsDelay), Procedure:=tProc
End Function

