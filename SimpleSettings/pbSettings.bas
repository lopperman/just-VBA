Attribute VB_Name = "pbSettings"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
' Manage settings at a Workbook level
' Keeps Settings Updated in Hidden Worksheet and
' Automatically Keep Dictionary Synchronized with Settings
' For fast access
' (Automatically  builds and configures setting sheet if not exist)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbSettings.bas
'  license GNU General Public License v3.0
'  Updated 14-Apr-2023
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Option Explicit
Option Compare Text
Option Base 1


Private Const SETTING_LSTOBJNAME As String = "tblPBSETTINGS"
Public Const SETTING_WSNAME As String = "pb-Settings"
Private lsettingDict As Dictionary
Private lLO As ListObject

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
' Enum for StringsMatch Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Enum strMatchEnum
    smEqual = 0
    smNotEqualTo = 1
    smContains = 2
    smStartsWithStr = 3
    smEndWithStr = 4
End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
' Return Setting Value, if [keyName] Exists, otherwise
'   If [defaultVal] was provided, return defaultVal, otherwise
'   return Empty
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function GetValue(keyName, Optional defaultVal)
    CheckDict
    If KeyExists(keyName) Then
        GetValue = lsettingDict(keyName)
    ElseIf Not IsMissing(defaultVal) Then
        GetValue = defaultVal
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Create or Update Setting [keyName] to be[keyValue]
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function SetValue(keyName, keyValue)
    CheckDict
    CreateOrUpdate keyName, keyValue
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Return TRUE if [keyName] exists in Setting collection
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function KeyExists(keyName) As Boolean
    CheckDict
    KeyExists = lsettingDict.Exists(keyName)
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   will delete setting [keyName] if it exists
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function Delete(keyName)
    CheckDict
    If KeyExists(keyName) Then
        lsettingDict.Remove (keyName)
        Dim rng As Range
        Set rng = SettingLO.ListColumns(1).Range.Find(keyName, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=False)
        If Not rng Is Nothing Then
            rng.EntireRow.Delete xlShiftUp
        End If
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Make the Settings Sheet Temporarily Visible
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function ShowSettingsSheet()
    MsgBox "Settings Sheet Will Automatically be Hidden Next Time Any Settings Function Is Called"
    If Not SettingSheet Is Nothing Then
        SettingSheet.Visible = xlSheetVisible
        SettingSheet.Activate
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Returns Number of Settings
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Property Get SettingCount() As Long
    CheckDict
    SettingCount = lsettingDict.Count
End Property


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Returns All Settings Keys/Values in a 2-dimension array
'   e.g.
'       Dim tmpKey, tmpVal
'       tmpKey = AllSettings(1,1) 'first setting key
'       tmpVal = AllSettings(1,2) 'first setting value
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function AllSettings() As Variant
    CheckDict
    AllSettings = SettingLO.ListColumns(1).DataBodyRange.Resize(ColumnSize:=2).value
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Create or Update Setting [stgKey] to be [stgVal]
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function CreateOrUpdate(stgKey, stgVal)
    CheckDict
    Dim tArr As Variant, idx
    Dim evts As Boolean, scrn As Boolean, intr As Boolean
    evts = Application.EnableEvents
    scrn = Application.ScreenUpdating
    intr = Application.Interactive
    SysMode
    If KeyExists(stgKey) Then
        tArr = SettingLO.DataBodyRange.value
        For idx = LBound(tArr) To UBound(tArr)
            If StringsMatch(tArr(idx, 1), stgKey) Then
                SettingLO.ListRows(idx).Range.Offset(ColumnOffset:=1).Resize(ColumnSize:=2).value = Array(stgVal, Now)
                lsettingDict(stgKey) = stgVal
                Exit For
            End If
        Next idx
    Else
        AddSettingRow stgKey, stgVal
        lsettingDict.Add stgKey, stgVal
    End If
    SysMode evts, scrn, intr
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   (Private) Get reference to the settings ListObject
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function SettingLO() As ListObject
    If lLO Is Nothing Then
        Set lLO = SettingSheet.ListObjects(SETTING_LSTOBJNAME)
    End If
    Set SettingLO = lLO
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   (Private) Ensures SettingsSheet, Settings ListObject, and
'       default values exist
'   Loads all settings into dictionary for fast retrieval
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function CheckDict()
    If SettingSheet Is Nothing Then
        CreateSettingsSheet
    End If
    If Not SettingSheet.Visible = xlSheetVeryHidden Then
        SettingSheet.Visible = xlSheetVeryHidden
    End If
    If lsettingDict Is Nothing Then
      Set lsettingDict = New Dictionary
      lsettingDict.CompareMode = TextCompare
      LoadSettingsDict
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   (Private) Adds a new settings row to settings listobject
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function AddSettingRow(stgKey, stgVal)
    Dim evts As Boolean: evts = Application.EnableEvents
    Application.EnableEvents = False
    With SettingLO
        .Resize .Range.Resize(RowSize:=.Range.Rows.Count + 1)
        .ListRows(.ListRows.Count).Range.value = Array(stgKey, stgVal, Now)
    End With
    Application.EnableEvents = evts
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   (Private) Function that loads settings into dictionary
'   If no settings rows exists, will create a row with
'   'Version' Settings Key
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function LoadSettingsDict()
    Dim tArr As Variant, idx
    Set lsettingDict = Nothing
    Set lsettingDict = New Dictionary
    lsettingDict.CompareMode = TextCompare
    If SettingLO.ListRows.Count = 0 Then
        AddSettingRow "VERSION", CDbl(0.01)
    End If
    tArr = SettingLO.DataBodyRange.value
    For idx = LBound(tArr) To UBound(tArr)
        If Not lsettingDict.Exists(tArr(idx, 1)) Then
            lsettingDict.Add tArr(idx, 1), tArr(idx, 2)
        End If
    Next idx
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   (Private) If Settings Worksheet does not exist, this method
'   is used to create the new worksheet with settings listobject
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function CreateSettingsSheet() As Boolean
On Error Resume Next
    Dim tWS As Worksheet
    Dim tLO As ListObject
    Dim evts As Boolean, scrn As Boolean, intr As Boolean
    evts = Application.EnableEvents
    scrn = Application.ScreenUpdating
    intr = Application.Interactive
    SysMode
    Set tWS = ThisWorkbook.Worksheets.Add
    tWS.Name = SETTING_WSNAME
    tWS.Parent.Windows(1).DisplayGridlines = False
    tWS.Move After:=LastVisibleSheet
    Set tWS = Nothing
    If Not SettingSheet Is Nothing Then
        With SettingSheet
            .Range("A1:C1").value = Array("SettingKey", "SettingValue", "Updated")
            .Range("A2:C2").value = Array("justVBA-GitHub", "https://github.com/lopperman/just-VBA", Now)
            .Range("A3:C3").value = Array("VERSION", CDbl(0.01), Now)
            Set tLO = .ListObjects.Add(SourceType:=xlSrcRange, Source:=.Range("A1:C3"), XLListObjectHasHeaders:=xlYes)
            tLO.Name = SETTING_LSTOBJNAME
            tLO.Range.EntireColumn.AutoFit
            Set tLO = Nothing
        End With
    End If
    SysMode evts, scrn, intr
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   (Private) Returns Settings Worksheet, if Found
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function SettingSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SETTING_WSNAME Then
            Set SettingSheet = ws
            Exit For
        End If
    Next ws
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   (Private) Returns the last visible Worksheet in current
'   workbook.
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function LastVisibleSheet() As Worksheet
    Dim lstIndex As Long
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And ws.Index > lstIndex Then
            lstIndex = ws.Index
        End If
    Next ws
    Set LastVisibleSheet = ThisWorkbook.Worksheets(lstIndex)
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   (Private) Used to turn off Events and Screen, or restore
'   Events and Screen settings to previous values
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function SysMode(Optional evts As Boolean = False, Optional scrn As Boolean = False, Optional intr As Boolean = False)
    Application.EnableEvents = evts
    Application.ScreenUpdating = scrn
    Application.Interactive = intr
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   (Private) my string comparison function
'   Made private to not interfere with version you might have
'   in other modules
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function StringsMatch( _
    ByVal checkString As Variant, ByVal _
    validString As Variant, _
    Optional smEnum As strMatchEnum = strMatchEnum.smEqual, _
    Optional compMethod As VbCompareMethod = vbTextCompare) As Boolean
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


