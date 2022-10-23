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
Private l_pbPackageRunning As Boolean

Private l_preventProtection As Boolean
Private l_OperatingState As ftOperatingState


    
Public Property Get ENV_User() As String
    ENV_User = VBA.Interaction.Environ("USER")
End Property

Public Function ENV_LogName() As String
On Error Resume Next
    #If Mac Then
        ENV_LogName = VBA.Interaction.Environ("LOGNAME")
    #Else
        ENV_LogName = VBA.Interaction.Environ("USERNAME")
    #End If
End Function

Public Property Get ENV_HOME() As String
    ENV_HOME = VBA.Interaction.Environ("HOME")
End Property
Public Property Get ENV_TEMPDIR() As String
    ENV_TEMPDIR = VBA.Interaction.Environ("TMPDIR")
End Property


    
'Private l_pbPackageRunning As Boolean
Public Property Get pbPackageRunning() As Boolean
    pbPackageRunning = l_pbPackageRunning
End Property
Public Property Let pbPackageRunning(vl As Boolean)
    l_pbPackageRunning = vl
End Property
    
    
    

    
    

    

    
Public Property Get DBLQUOTE() As String
    DBLQUOTE = Chr(34)
End Property
    
    

    
    
Public Function DtAdd(intervalType As DateDiffType, _
    number As Variant, ByVal dt As Variant) As Variant
    
    Dim retVal As Variant
    
    Select Case intervalType
        Case DateDiffType.dtday
            retVal = DateAdd("d", number, dt)
        Case DateDiffType.dtDayOfYear
            retVal = DateAdd("y", number, dt)
        Case DateDiffType.dtHour
            retVal = DateAdd("h", number, dt)
        Case DateDiffType.dtMinute
            retVal = DateAdd("n", number, dt)
        Case DateDiffType.dtMonth
            retVal = DateAdd("m", number, dt)
        Case DateDiffType.dtQuarter
            retVal = DateAdd("q", number, dt)
        Case DateDiffType.dtSecond
            retVal = DateAdd("s", number, dt)
        Case DateDiffType.dtWeekday
            retVal = DateAdd("w", number, dt)
        Case DateDiffType.dtWeek
            retVal = DateAdd("ww", number, dt)
        Case DateDiffType.dtYear
            retVal = DateAdd("yyyy", number, dt)
    End Select
    
    DtAdd = retVal
    
End Function

Public Function DtPart(thePart As DateDiffType, dt1 As Variant, _
    Optional ByVal firstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional ByVal firstWeekOfYear As VbFirstWeekOfYear = VbFirstWeekOfYear.vbFirstJan1) As Variant
    Select Case thePart
        Case DateDiffType.dtDate_NoTime
            DtPart = DateSerial(DtPart(dtYear, dt1), DtPart(dtMonth, dt1), DtPart(dtday, dt1))
        Case DateDiffType.dtday
            DtPart = DatePart("d", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtDayOfYear
            DtPart = DatePart("y", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtHour
            DtPart = DatePart("h", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtMinute
            DtPart = DatePart("n", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtMonth
            DtPart = DatePart("m", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtQuarter
            DtPart = DatePart("q", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtSecond
            DtPart = DatePart("s", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtWeek
            DtPart = DatePart("ww", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtWeekday
            DtPart = DatePart("w", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtYear
            DtPart = DatePart("yyyy", dt1, firstDayOfWeek, firstWeekOfYear)
    End Select
End Function

Public Function DtDiff(diffType As DateDiffType, _
    dt1 As Variant, Optional ByVal dt2 As Variant, _
    Optional firstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional firstWeekOfYear As VbFirstWeekOfYear = VbFirstWeekOfYear.vbFirstJan1, _
    Optional returnFraction As Boolean = False) As Variant
' ~~~ FRACTIONAL RETURN VALUES ONLY SUPPORTED FOR
'        minutes, hours, days, weeks
' ~~~ note:  fractionals are based on type of date/time component
' ~~~ for example, if the difference in time was 2 minutes, 30 seconds
' ~~~ and you were returning Minutes as a fractions, the return value would
' ~~~ be 2.5 (for 2 1/2 minutes)
    
    If IsMissing(dt2) Then dt2 = Now
    Dim retVal As Variant
    Dim tmpVal1 As Variant
    Dim tmpVal2 As Variant
    Dim tmpRemain As Variant
    
    Select Case diffType
        Case DateDiffType.dtSecond
            retVal = DateDiff("s", dt1, dt2)
        Case DateDiffType.dtWeekday
            retVal = DateDiff("w", dt1, dt2)
        Case DateDiffType.dtMinute
            If returnFraction Then
                ' fractions based on SECONDS (60)
                tmpVal1 = DtDiff(dtSecond, dt1, dt2)
                tmpVal2 = tmpVal1 - (DateDiff("n", dt1, dt2) * 60)
                If tmpVal2 > 0 Then
                    retVal = DateDiff("n", dt1, dt2) + (tmpVal2 / 60)
                Else
                    retVal = DateDiff("n", dt1, dt2)
                End If
            Else
                retVal = DateDiff("n", dt1, dt2)
            End If
        Case DateDiffType.dtHour
                ' fractions based on MINUTES (60)
            If returnFraction Then
                tmpVal1 = DtDiff(dtMinute, dt1, dt2)
                tmpVal2 = tmpVal1 - (DateDiff("h", dt1, dt2) * 60)
                If tmpVal2 > 0 Then
                    retVal = DateDiff("h", dt1, dt2) + (tmpVal2 / 60)
                Else
                    retVal = DateDiff("h", dt1, dt2)
                End If
            Else
                retVal = DateDiff("h", dt1, dt2)
            End If
        Case DateDiffType.dtday
                ' fractions based on HOURS (24)
            If returnFraction Then
                tmpVal1 = DtDiff(dtHour, dt1, dt2)
                tmpVal2 = tmpVal1 - (DateDiff("d", dt1, dt2) * 24)
                If tmpVal2 > 0 Then
                    retVal = DateDiff("d", dt1, dt2) + (tmpVal2 / 24)
                Else
                    retVal = DateDiff("d", dt1, dt2)
                End If
            Else
                retVal = DateDiff("d", dt1, dt2)
            End If
        Case DateDiffType.dtWeek
                ' fractions based on DAYS (7)
            If returnFraction Then
                tmpVal1 = DtDiff(dtday, dt1, dt2)
                tmpVal2 = tmpVal1 - (DateDiff("ww", dt1, dt2, firstDayOfWeek, firstWeekOfYear) * 7)
                If tmpVal2 > 0 Then
                    retVal = DateDiff("ww", dt1, dt2, firstDayOfWeek, firstWeekOfYear) + (tmpVal2 / 7)
                Else
                    retVal = DateDiff("ww", dt1, dt2, firstDayOfWeek, firstWeekOfYear)
                End If
            Else
                retVal = DateDiff("ww", dt1, dt2, firstDayOfWeek, firstWeekOfYear)
            End If
        Case DateDiffType.dtMonth
            retVal = DateDiff("m", dt1, dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtQuarter
            retVal = DateDiff("q", dt1, dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtYear
            retVal = DateDiff("yyyy", dt1, dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtDayOfYear
            retVal = DateDiff("y", dt1, dt1, firstDayOfWeek, firstWeekOfYear)
    End Select
    
    DtDiff = retVal
    
End Function

Public Sub CheckReady(Optional timeoutSec As Long = 20)
    On Error Resume Next
    If timeoutSec > 30 Then timeoutSec = 30
    Dim curTmr As Single, notReadyLogged As Boolean
    curTmr = Timer
    Do While Application.Ready = False
        If notReadyLogged = False Then
            notReadyLogged = True
        End If
        If Timer - curTmr >= timeoutSec Then
            Exit Do
        End If
        DoEvents
    Loop
    If Not Err.number = 0 Then
        Err.Clear
    End If
End Sub

Public Function InVisibleRange(activeSheetAddress As String, Optional scrollTo As Boolean = False) As Boolean
On Error Resume Next
    If Not ThisWorkbook.ActiveSheet Is Nothing Then
        If Intersect(ThisWorkbook.Windows(1).VisibleRange, ThisWorkbook.ActiveSheet.Range(activeSheetAddress).Cells(1, 1)) Is Nothing Then
            InVisibleRange = False
        Else
            InVisibleRange = True
        End If
    End If
    
    If InVisibleRange = False And scrollTo = True Then
        Dim scrn As Boolean: scrn = Application.ScreenUpdating
        Application.ScreenUpdating = True
        Application.GoTo Reference:=ThisWorkbook.ActiveSheet.Range(activeSheetAddress).Cells(1, 1), Scroll:=True
        DoEvents
        Application.ScreenUpdating = scrn
    End If
    
    If Err.number <> 0 Then
        Trace ConcatWithDelim(", ", "Error pbMiscUtil.InVisibleRange", "Address: ", ActiveSheetName, activeSheetAddress, Err.number, Err.Description), forceWrite:=True, forceDebug:=True
        Err.Clear
    End If
End Function

'Public Function FullNameCorrectedByName(wbName) As String
'    If WorkbookIsOpen(wbName) Then
'        FullNameCorrectedByName = wbName
'    End If
'End Function

Public Function FullWbNameCorrected(Optional wkbk As Workbook) As String
On Error Resume Next
    Dim fName As String
    If wkbk Is Nothing Then
        fName = ThisWorkbook.FullName
    Else
        fName = wkbk.FullName
    End If
    If Len(fName) > 0 Then
        If InStr(1, fName, "http", vbTextCompare) > 0 Then
            fName = Replace(fName, " ", "%20", compare:=vbTextCompare)
        End If
    End If
    FullWbNameCorrected = fName
    If Err.number <> 0 Then Err.Clear
End Function


Public Function SimpleURLEncode(ByVal fPath As String) As String
    If Len(fPath) > 0 Then
        If InStr(1, fPath, "http", vbTextCompare) > 0 Then
            fPath = Replace(fPath, " ", "%20", compare:=vbTextCompare)
        End If
    End If
    SimpleURLEncode = fPath
    If Err.number <> 0 Then Err.Clear

End Function



    
    Public Function StringsMatch( _
        ByVal str1 As Variant, ByVal _
        str2 As Variant, _
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
            
        str1 = CStr(str1)
        str2 = CStr(str2)
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
    ftBeep btError
    If Not raiseErrorOnFail Then
        Err.Clear
        On Error GoTo 0
    Else
        Err.Raise Err.number, Err.Description
    End If
End Function


Public Function wt(listObjectName As String, ParamArray tempListObjPrefixes() As Variant) As ListObject
'   Return object reference to ListObject in 'ThisWorkbook' called [listObjectName]
'   This function exists to eliminate problem with getting a ListObject using the 'Range([list object name])
'       where the incorrect List Object could be returned if the ActiveWorkbook containst a list object
'       with the same name, and is not the intended ListObject
'  If temporary list object mayexists, include the prefixes (e.g. "tmp","temp") to identify and not add to dictionary
On Error GoTo E:
    Static l_listObjDict As Dictionary
    Dim ws As Worksheet, t As ListObject, ignoreArr As Variant, ignoreAI As ArrInformation, ignoreIdx As Long, ignore As Boolean
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
        Set l_listObjDict = New Dictionary
        For Each ws In ThisWorkbook.Worksheets
            For Each t In ws.ListObjects
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
        Next ws
        DoEvents
    End If
    
    'this covers the temporary listobject which may not always be available
    If Not l_listObjDict.Exists(listObjectName) Then
        Dim tWS As Worksheet, tLO As ListObject
        For Each tWS In ThisWorkbook.Worksheets
            For Each tLO In tWS.ListObjects
                If StrComp(tLO.Name, listObjectName, vbTextCompare) = 0 Then
                    'DON'T ADD any tmp tables
                    Set wt = tLO
                    GoTo Finalize:
                End If
            Next tLO
        Next tWS
    End If
    
Finalize:
    On Error Resume Next
    
    If l_listObjDict.Exists(listObjectName) Then
        Set wt = l_listObjDict(listObjectName)
    End If
    If Err.number <> 0 Then Err.Clear
    Exit Function
E:
    ftBeep btError
    DebugPrint "Error getting list object " & listObjectName
    Err.Clear
End Function


' ~~~ ~~ FLAG ENUM COMPARE ~~~ ~~~
Public Function EnumCompare(theEnum As Variant, enumMember As Variant, Optional ByVal iType As ecComparisonType = ecComparisonType.ecOR) As Boolean
    Dim c As Long
    c = theEnum And enumMember
    EnumCompare = IIf(iType = ecOR, c <> 0, c = enumMember)
End Function


' ~~~~~~~~~~   INPUT BOX   ~~~~~~~~~~'
Public Function InputBox_FT(prompt As String, Optional title As String = "Financial Tool - Input Needed", Optional default As Variant, Optional inputType As ftInputBoxType) As Variant
    ftBeep btMsgBoxChoice
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
    If Not EnumCompare(buttons, vbSystemModal) Then buttons = buttons + vbSystemModal
    If Not EnumCompare(buttons, vbMsgBoxSetForeground) Then buttons = buttons + vbMsgBoxSetForeground
    If EnumCompare(buttons, vbOKOnly) Then
        ftBeep btMsgBoxOK
    Else
        ftBeep btMsgBoxChoice
    End If
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
    ftBeep btMsgBoxChoice
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
    If Err.number <> 0 Then Err.Clear
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
    If Err.number <> 0 Then Err.Clear
End Property






Public Function WorksheetExists(sName As String, Optional wbk As Workbook) As Boolean
On Error Resume Next
    If wbk Is Nothing Then
        Set wbk = ThisWorkbook
    End If
    Dim ws As Worksheet
    Set ws = wbk.Worksheets(sName)
    If Err.number = 0 Then
        WorksheetExists = True
    End If
    Set ws = Nothing
    If Err.number <> 0 Then Err.Clear
End Function



Public Function WorkbookIsOpen(ByVal wkBkName As String, Optional checkCodeName As String = vbNullString) As Boolean
On Error Resume Next
    wkBkName = FileNameFromFullPath(wkBkName)
    If Not Workbooks(wkBkName) Is Nothing Then
        WorkbookIsOpen = True
    Else
        WorkbookIsOpen = False
    End If
    If Err.number <> 0 Then
        WorkbookIsOpen = False
    End If
    If WorkbookIsOpen And Len(checkCodeName) > 0 Then
        If StringsMatch(Workbooks(wkBkName).CodeName, checkCodeName) Then
            WorkbookIsOpen = True
        Else
            WorkbookIsOpen = False
        End If
    End If
    If Err.number <> 0 Then Err.Clear
    Exit Function

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
    If Err.number <> 0 Then Err.Clear
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
'    Debug.Print "End Wait: Waited For: " & Math.Round((Timer - stTimer), 3) & " seconds"
    
End Function



Public Function CallOnTime_TwoArg(ByVal wbName As String, ByVal procName As String, ByVal argVal1 As String, ByVal argVal2 As String, Optional ByVal secondsDelay As Long = 0)
    'FT HELPER NEEDS TO BE UPDATED AND TESTED BEFORE ALLOWING THE PARAMETER TO GO THROUGH
    On Error Resume Next
    Dim litDQ As String
    litDQ = """"
    
    wbName = wbName
    If TypeName(argVal1) = "String" Then
        If StringsMatch(argVal1, ".xlam", smContains) Or StringsMatch(argVal1, ".xlsm", smContains) Then
            argVal1 = argVal1
        End If
    End If
    If TypeName(argVal2) = "String" Then
        If StringsMatch(argVal2, ".xlam", smContains) Or StringsMatch(argVal2, ".xlsm", smContains) Then
            argVal2 = argVal2
        End If
    End If
    
    wbName = CleanSingleTicks(wbName)
    argVal1 = CleanSingleTicks(argVal1)
    argVal2 = CleanSingleTicks(argVal2)
    
    DoEvents
    Application.ONTIME EarliestTime:=GetTimeDelay(secondsDelay), Procedure:="'" & wbName & "'!'" & procName & " " & litDQ & argVal1 & litDQ & "," & litDQ & argVal2 & litDQ & "'"
    If Err.number <> 0 Then Err.Clear
    
End Function
'
'Public Function CallOnTime_TwoArg(wbName As String, procName As String, argVal1 As String, argVal2 As String, Optional secondsDelay As Long = 0)
'    'FT HELPER NEEDS TO BE UPDATED AND TESTED BEFORE ALLOWING THE PARAMETER TO GO THROUGH
'    wbName = CleanSingleTicks(wbName)
'    argVal1 = CleanSingleTicks(argVal1)
'    argVal2 = CleanSingleTicks(argVal2)
'
'    Dim litDQ As String
'    litDQ = """"
'
'    Application.ONTIME EarliestTime:=GetTimeDelay(secondsDelay), Procedure:="'" & wbName & "'!'" & procName & " " & litDQ & argVal1 & litDQ & "," & litDQ & argVal2 & litDQ & "'"
'End Function

Public Function WrapExternalCall(wbName As String, procName As String, argVal As Variant) As String
    If TypeName(argVal) = "String" Then
        argVal = CleanSingleTicks(CStr(argVal))
        WrapExternalCall = "'" & wbName & "'!'" & procName & " " & """" & argVal & """'"
    Else
        WrapExternalCall = "'" & wbName & "'!'" & procName & " " & "" & argVal & "'"
    End If
End Function

Public Function CallOnTime_OneArg(wbName As String, procName As String, argVal As Variant, Optional secondsDelay As Long = 0)
On Error Resume Next
    Dim litDQ As String
    litDQ = """"
    
    wbName = wbName
    If TypeName(argVal) = "String" Then
        If StringsMatch(argVal, ".xlam", smContains) Or StringsMatch(argVal, ".xlsm", smContains) Then
            argVal = argVal
        End If
    End If
    If TypeName(argVal) = "String" Then
        'argVal = Replace(argVal, "'", "''")
        Application.ONTIME EarliestTime:=GetTimeDelay(secondsDelay), Procedure:="'" & wbName & "'!'" & procName & " " & litDQ & argVal & litDQ & "'"
    Else
        Application.ONTIME EarliestTime:=GetTimeDelay(secondsDelay), Procedure:="'" & wbName & "'!'" & procName & " " & argVal & "'"
    End If

    If Err.number <> 0 Then Err.Clear
End Function

Public Function GetTimeDelay(Optional secondsDelay As Long = 0) As Date
    If secondsDelay > 59 Then secondsDelay = 59
    If secondsDelay < 0 Then secondsDelay = 0
    GetTimeDelay = Now + TimeValue("00:00:" & Format(secondsDelay, "00"))
End Function

Public Function CallOnTime(wbName As String, procName As String, Optional secondsDelay As Long = 0)
On Error Resume Next
    wbName = wbName
    Dim tProc As String
    tProc = "'" & wbName & "'!'" & procName & "'"
    Application.ONTIME EarliestTime:=GetTimeDelay(secondsDelay), Procedure:=tProc
    If Err.number <> 0 Then Err.Clear
End Function

Public Property Get ActiveSheetName() As String
    If Not ThisWorkbook.ActiveSheet Is Nothing Then
        ActiveSheetName = ThisWorkbook.ActiveSheet.Name
    End If
End Property


Public Property Get Events() As Boolean
    Events = Application.EnableEvents
End Property
Public Property Let Events(evtOn As Boolean)
    If Not Application.EnableEvents = evtOn Then
        Application.EnableEvents = evtOn
    End If
End Property
Public Function EventsOff()
    Events = False
End Function
Public Function EventsOn()
    Events = True
End Function



'   Example Usage: ConcatWithDelim(", ","Why","Doesn't","VBA","Have","This")
'       outputs:  Why, Doesn't, VBA, Have, This
Public Function ConcatWithDelim(ByVal delimeter As String, ParamArray Items() As Variant) As String
    ConcatWithDelim = Join(Items, delimeter)
End Function

'RETURN STRING FOR EACH ROW REPRESENTED IN RANGE, vbNewLine as Line Delimeter
'   Example 1 (Get your Column Names for a list object)
'       Dim lo as ListObject
'       Set lo = wsTeamInfo.ListOobjects("tblTeamInfo")
'       DebugPrint ConcatRange(lo.HeaderRowRange)
'           outputs:  StartDt|EndDt|Project|Employee|Role|BillRate|EstCostRt|ActCostRt|Active|TaskName|SegName|AllocPerc|Utilization|Bill_Hrs|NonBill_Hrs|CfgID|ActiveHidden|Updated
'   Example 2 (let's grab some weird ranges)
'       Dim rng As Range
'       Set rng = wsDashboard.Range("E49:J50")
'       Set rng = Union(rng, wsDashboard.Range("L60:Q60"))
'       DebugPrint ConcatRange(rng)
'           Outputs:
'               8/16/21|8/22/21|Actual|0|0|0
'               8/23/21|8/29/21|Actual|23762.5|13799.5|9963
'               386274.85|18276.05|10631.35|7644.7|0.4182906043702|
Public Function ConcatRange(rng As Range, Optional delimeter As String = "|") As String
    Dim rngArea As Range, rRow As Long, rCol As Long, retV As String, rArea As Long
    For rArea = 1 To rng.Areas.Count
        For rRow = 1 To rng.Areas(rArea).Rows.Count
            If Len(retV) > 0 Then
                retV = retV & vbNewLine
            End If
            For rCol = 1 To rng.Areas(rArea).Columns.Count
                If rCol = 1 Then
                    retV = ConcatWithDelim("", retV, rng.Areas(rArea)(rRow, rCol).value)
                Else
                    retV = ConcatWithDelim(delimeter, retV, rng.Areas(rArea)(rRow, rCol).value)
                End If
            Next rCol
        Next rRow
    Next rArea
    
    ConcatRange = retV

End Function

'   Example Usage: Dim msg as string: msg = "Hello There today's date is: ": DebugPrint Concat(msg,Date)
'       outputs: Hello There today's date is: 5/24/22
Public Function Concat(ParamArray Items() As Variant) As String
    Concat = Join(Items, "")
End Function

Public Function ftCreateWorkbook(Optional ByVal tmplPath As Variant) As Workbook
On Error Resume Next
    Dim retWB As Workbook
    If IsMissing(tmplPath) Then
        Set retWB = Workbooks.Add
        DoEvents
    Else
        Dim tFullPath As String
        tFullPath = PathCombine(False, tmplPath)
        Set retWB = Workbooks.Add(Template:=tFullPath)
        DoEvents
    End If
    Set ftCreateWorkbook = retWB
    If Not Err.number = 0 Then
        Err.Clear
    End If
End Function

' BEGIN ~~~ ~~~ SHEET PROTECTION    ~~~ ~~~
Public Function ProtectWS(ws As Worksheet _
    , Optional protOpt As ProtectionTemplate = ProtectionTemplate.ptDefault _
    , Optional pwdOption As ProtectionPWD = ProtectionPWD.pwStandard _
    , Optional customTemplate As SheetProtection) As Boolean
    
    Dim pwd As String
    Select Case pwdOption
        Case ProtectionPWD.pwStandard
            pwd = CFG_PROTECT_PASSWORD
        Case ProtectionPWD.pwExport
            pwd = CFG_PROTECT_PASSWORD_EXPORT
        Case ProtectionPWD.pwMisc
            pwd = CFG_PROTECT_PASSWORD_MISC
        Case ProtectionPWD.pwLog
            'Most Secure
            pwd = CFG_P_LOG
    End Select
    
    Dim prt As SheetProtection
    Select Case protOpt
        Case ProtectionTemplate.ptDefault
            prt = ProtectShtDefault
        Case ProtectionTemplate.ptDenyFilterSort
            prt = ProtectShtCustom(False)
        Case ProtectionTemplate.ptAllowFilterSort
            prt = ProtectShtCustom(True)
        Case ProtectionTemplate.ptCustom
            prt = customTemplate
    End Select

End Function

Public Function ProtectShtCustom(allowFilterSort As Boolean) As SheetProtection
    Dim protSht As SheetProtection
    protSht = ProtectShtDefault
    If allowFilterSort = False Then
        If EnumCompare(protSht, SheetProtection.psAllowFiltering) Then
            protSht = protSht - SheetProtection.psAllowFiltering
        End If
        If EnumCompare(protSht, SheetProtection.psAllowSorting) Then
            protSht = protSht - SheetProtection.psAllowSorting
        End If
    Else
        If Not EnumCompare(protSht, SheetProtection.psAllowFiltering) Then
            protSht = protSht + SheetProtection.psAllowFiltering
        End If
        If Not EnumCompare(protSht, SheetProtection.psAllowSorting) Then
            protSht = protSht + SheetProtection.psAllowSorting
        End If
    End If
    
    ProtectShtCustom = protSht
    
End Function

Public Function ProtectShtDefault( _
    Optional pContents As Boolean = True, _
    Optional pUsePassword As Boolean = True, _
    Optional pDrawingObjects As Boolean = False, _
    Optional pScenarios As Boolean = False, _
    Optional pUserInterfaceOnly As Boolean = True, _
    Optional pAllowFormattingCells As Boolean = True, _
    Optional pAllowFormattingColumns As Boolean = True, _
    Optional pAllowFormattingRows As Boolean = True, _
    Optional pAllowInsertingColumns As Boolean = False, _
    Optional pAllowInsertingRows As Boolean = False, _
    Optional pAllowInsertingHyperlinks As Boolean = False, _
    Optional pAllowDeletingColumns As Boolean = False, _
    Optional pAllowDeletingRows As Boolean = False, _
    Optional pAllowSorting As Boolean = True, _
    Optional pAllowFiltering As Boolean = True, _
    Optional pAllowUsingPivotTables As Boolean = False) As SheetProtection

    Dim protSht As SheetProtection
    protSht = protSht + IIf(pContents, SheetProtection.psContents, 0)
    protSht = protSht + IIf(pUsePassword, SheetProtection.psUsePassword, 0)
    protSht = protSht + IIf(pDrawingObjects, SheetProtection.psDrawingObjects, 0)
    protSht = protSht + IIf(pScenarios, SheetProtection.psScenarios, 0)
    protSht = protSht + IIf(pUserInterfaceOnly, SheetProtection.psUserInterfaceOnly, 0)
    protSht = protSht + IIf(pAllowFormattingCells, SheetProtection.psAllowFormattingCells, 0)
    protSht = protSht + IIf(pAllowFormattingColumns, SheetProtection.psAllowFormattingColumns, 0)
    protSht = protSht + IIf(pAllowFormattingRows, SheetProtection.psAllowFormattingRows, 0)
    protSht = protSht + IIf(pAllowInsertingColumns, SheetProtection.psAllowInsertingColumns, 0)
    protSht = protSht + IIf(pAllowInsertingRows, SheetProtection.psAllowInsertingRows, 0)
    protSht = protSht + IIf(pAllowInsertingHyperlinks, SheetProtection.psAllowInsertingHyperlinks, 0)
    protSht = protSht + IIf(pAllowDeletingColumns, SheetProtection.psAllowDeletingColumns, 0)
    protSht = protSht + IIf(pAllowDeletingRows, SheetProtection.psAllowDeletingRows, 0)
    protSht = protSht + IIf(pAllowSorting, SheetProtection.psAllowSorting, 0)
    protSht = protSht + IIf(pAllowFiltering, SheetProtection.psAllowFiltering, 0)
    protSht = protSht + IIf(pAllowUsingPivotTables, SheetProtection.psAllowUsingPivotTables, 0)
    
    ProtectShtDefault = protSht
End Function


' END ~~~ ~~~ SHEET PROTECTION    ~~~ ~~~



Public Property Get StartupPath() As String
    StartupPath = PathCombine(True, Application.StartupPath)
End Property

Public Property Let PreventProtection(preventProtect As Boolean)
    l_preventProtection = preventProtect
End Property
Public Property Get PreventProtection() As Boolean
    PreventProtection = l_preventProtection
End Property

' BEGIN ~~~ ~~~ OPERATING STATE    ~~~ ~~~
Public Function IsFTClosing() As Variant
    IsFTClosing = (l_OperatingState = ftClosing)
End Function
Public Function IsFTOpening() As Variant
    IsFTOpening = (l_OperatingState = ftOpening)
End Function

Public Property Get ftState() As ftOperatingState
    ftState = l_OperatingState
End Property
Public Property Let ftState(ftsVal As ftOperatingState)
    l_OperatingState = ftsVal
End Property
' END ~~~ ~~~ OPERATING STATE    ~~~ ~~~

Public Function URLEncode(ByRef txt As String) As String
    Dim buffer As String, i As Long, c As Long, n As Long
    buffer = String$(Len(txt) * 12, "%")
 
    For i = 1 To Len(txt)
        c = AscW(Mid$(txt, i, 1)) And 65535
 
        Select Case c
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95  ' Unescaped 0-9A-Za-z-._ '
                n = n + 1
                Mid$(buffer, n) = ChrW(c)
            Case Is <= 127            ' Escaped UTF-8 1 bytes U+0000 to U+007F '
                n = n + 3
                Mid$(buffer, n - 1) = Right$(Hex$(256 + c), 2)
            Case Is <= 2047           ' Escaped UTF-8 2 bytes U+0080 to U+07FF '
                n = n + 6
                Mid$(buffer, n - 4) = Hex$(192 + (c \ 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case 55296 To 57343       ' Escaped UTF-8 4 bytes U+010000 to U+10FFFF '
                i = i + 1
                c = 65536 + (c Mod 1024) * 1024 + (AscW(Mid$(txt, i, 1)) And 1023)
                n = n + 12
                Mid$(buffer, n - 10) = Hex$(240 + (c \ 262144))
                Mid$(buffer, n - 7) = Hex$(128 + ((c \ 4096) Mod 64))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case Else                 ' Escaped UTF-8 3 bytes U+0800 to U+FFFF '
                n = n + 9
                Mid$(buffer, n - 7) = Hex$(224 + (c \ 4096))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
        End Select
    Next
    URLEncode = Left$(buffer, n)
End Function

Public Function SanitizeAlpha(ByVal vl As String, Optional alsoAllowChars As String = vbNullString) As String
'   strips out EVERYTHING that isn't A-Z
    Dim retV As String
    retV = vl
    If Len(retV) = 0 Then
        retV = vbNullString
        SanitizeAlpha = retV
        Exit Function
    End If
    Dim validChars As String: validChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    If Len(alsoAllowChars) > 0 Then validChars = validChars & alsoAllowChars
    Dim i As Long
    For i = Len(retV) To 1 Step -1
        If Not StringsMatch(validChars, Mid(retV, i, 1), smContains, vbBinaryCompare) Then
            retV = Replace(retV, Mid(retV, i, 1), "", compare:=vbBinaryCompare)
        End If
    Next i
    SanitizeAlpha = Trim(retV)
End Function


Function ReplaceIllegalCharacters(ByVal strIn As String, ByVal strChar As String, Optional ByVal padSingleQuote As Boolean = True, Optional useForSpecialChars As Variant) As String
    Dim strSpecialChars As String
    Dim i As Long
    If IsMissing(useForSpecialChars) Then
        strSpecialChars = "~""#%&*:<>?{|}/\[]" & Chr(10) & Chr(13)
    Else
        strSpecialChars = useForSpecialChars
    End If

    For i = 1 To Len(strSpecialChars)
        strIn = Replace(strIn, Mid$(strSpecialChars, i, 1), strChar)
    Next
    
    If padSingleQuote And InStr(1, strIn, "''") = 0 Then
        strIn = CleanSingleTicks(strIn)
    End If
    
    ReplaceIllegalCharacters = strIn
End Function
