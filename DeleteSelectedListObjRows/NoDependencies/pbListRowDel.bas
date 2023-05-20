Attribute VB_Name = "pbListRowDel"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  Self-Contained Module for Delete User-Selected
'   rows from List Objects
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbListRowDel.bas
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1


Private ignoreProtectSheets As New Collection


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   used to cache protection settings for specific worksheet
'   Key=WorkbookName.WorksheetCodeName
'   Value = 1-based Array:
'   (1)=password,(2)=SheetProtection enum
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private protSettingCache As New Collection

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   GENERALIZED CONSTANTS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Const CACHED_PROTECT_PWD_POSITION As Long = 1
Public Const CACHED_PROTECT_ENUM_POSITION As Long = 2
Public Const CFG_PROTECT_PASSWORD As String = "00000"

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   CUSTOM ERROR CONSTANTS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   *** GENERIC ERRORS ONLY
'   *** STARTS AT 1001 - (513-1000 RESERVED FOR APPLICATION SPECIFIC ERRORS)
'Public Const ERR_OBSOLETE As Long = vbObjectError + 1001
'Public Const ERR_MISSING_EXPECTED_WORKSHEET As Long = vbObjectError + 1002
'Public Const ERR_HOME_SHEET_NOT_SET As Long = vbObjectError + 1003
'Public Const ERR_HOME_SHEET_ALREADY_SET As Long = vbObjectError + 1004
'Public Const ERR_INVALID_CALLER_SOURCE As Long = vbObjectError + 1005
'Public Const ERR_ACTION_PATH_NOT_DEFINED As Long = vbObjectError + 1006
'Public Const ERR_UNABLE_TO_GET_WORKSHEET_BUTTON As Long = vbObjectError + 1007
'Public Const ERR_RANGE_SRC_TARGET_MISMATCH = vbObjectError + 1008
'Public Const ERR_INVALID_FT_ARGUMENTS = vbObjectError + 1009
'Public Const ERR_NO_ERROR As String = "(INFO)"
'Public Const ERR_ERROR As String = "(ERROR)"
'Public Const ERR_INVALID_SETTING_OPERATION As Long = vbObjectError + 1010
'Public Const ERR_TARGET_OBJECT_MUST_BE_EMPTY As Long = vbObjectError + 1011
'Public Const ERR_EXPECTED_SHEET_NOT_ACTIVE As Long = vbObjectError + 1012
'Public Const ERR_INVALID_RANGE_SIZE = vbObjectError + 1013
'Public Const ERR_IMPORT_COLUMN_ALREADY_MATCHED = vbObjectError + 1014
'Public Const ERR_RANGE_AREA_COUNT = vbObjectError + 1015
'Public Const ERR_LIST_OBJECT_RESIZE_CANNOT_DELETE = vbObjectError + 1016
'Public Const ERR_LIST_OBJECT_RESIZE_INVALID_ARGUMENTS = vbObjectError + 1017
'Public Const ERR_INVALID_ARRAY_SIZE = vbObjectError + 1018
'Public Const ERR_EXPECTED_MANUAL_CALC_NOT_FOUND = vbObjectError + 1019
'Public Const ERR_WORKSHEET_OBJECT_NOT_FOUND = vbObjectError + 1020
'Public Const ERR_NOT_IMPLEMENTED_YET = vbObjectError + 1021
''CLASS INSTANCE PROHIBITED, FOR CLASS MODULE WITH Attribute VB_PredeclaredId = True
'Public Const ERR_CLASS_INSTANCE_PROHIBITED = vbObjectError + 1022
'Public Const ERR_CONTROL_STATE = vbObjectError + 1023
'Public Const ERR_PREVIOUS_PerfState_EXISTS = vbObjectError + 1024
'Public Const ERR_FORCED_WORKSHEET_NOT_FOUND = vbObjectError + 1025
'Public Const ERR_CANNOT_CHANGE_PACKAGE_PROPERTY = vbObjectError + 1026
'Public Const ERR_INVALID_PACKAGE_OPERATION = vbObjectError + 1027

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   GENERALIZED ENUMS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Enum DateDiffType
    dtSecond
    dtMinute
    dtHour
    dtDay
    dtWeek
    dtMonth
    dtYear
    dtQuarter
    dtDayOfYear
    dtWeekday
    dtDate_NoTime
    dtTimeString
End Enum
    
Public Enum ecComparisonType
    ecOR = 0 'default
    ecAnd
End Enum

Public Enum strMatchEnum
    smEqual = 0
    smNotEqualTo = 1
    smContains = 2
    smStartsWithStr = 3
    smEndWithStr = 4
End Enum

'
'    Public Enum ProtectionTemplate
'        ptDefault = 0
'        ptAllowFilterSort = 1
'        ptDenyFilterSort = 2
'        ptCustom = 3
'    End Enum
'
    
    Public Enum SheetProtection
        psContents = 2 ^ 0
        psUsePassword = 2 ^ 1
        psDrawingObjects = 2 ^ 2
        psScenarios = 2 ^ 3
        psUserInterfaceOnly = 2 ^ 4
        psAllowFormattingCells = 2 ^ 5
        psAllowFormattingColumns = 2 ^ 6
        psAllowFormattingRows = 2 ^ 7
        psAllowInsertingColumns = 2 ^ 8
        psAllowInsertingRows = 2 ^ 9
        psAllowInsertingHyperlinks = 2 ^ 10
        psAllowDeletingColumns = 2 ^ 11
        psAllowDeletingRows = 2 ^ 12
        psAllowSorting = 2 ^ 13
        psAllowFiltering = 2 ^ 14
        psAllowUsingPivotTables = 2 ^ 15
    End Enum
    Public Enum FlagEnumModify
        feVerifyEnumExists
        feVerifyEnumRemoved
    End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'
'       DELETE USER-SELECTED LIST OBJECT ROW(S)
'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
    '   Delete any row from DataBodyRange of a ListObject where at least one cell
    '       of row is Selected.  (Does not require contiguous selection area)
    '   If 'canDeleteMultiple' = False, then then only the first row in the ListObject
    '       that has any cells selected will be deleted
    '   If 'requestConfirmation' = True (Default) then user will be prompted to
    '       confirm delete action
    '   If the list object is on a Protected Worksheet, the (optional) password
    '       should be included ('protectionPwd').  If unprotecting the worksheet
    '       is necessary, it will be reprotected after the rows re deleted, and
    '       will be reprotected with the same protection options in place before
    '       this function was called
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
    Public Function DeleteSelectedListObjRows( _
        lstObj As ListObject _
        , Optional ByVal selectedRng As Range _
        , Optional ByVal canDeleteMultiple As Boolean = True _
        , Optional ByVal requestConfirmation As Boolean = True _
        , Optional ByVal protectionPwd As String = CFG_PROTECT_PASSWORD)
        '--- --- ---'
        On Error GoTo E:
        Dim failed As Boolean
        Dim evts As Boolean, scrn As Boolean, calc As XlCalculation
        Dim protOptions As SheetProtection, protErr As Boolean, protErrDesc
        protOptions = CurrentProtectionOptions(lstObj)
        '--- --- ---'
        'Check for error conditions'
        If selectedRng Is Nothing Then
            Set selectedRng = Selection
        End If
        If Not selectedRng.Worksheet Is lstObj.Range.Worksheet Then
            Err.Raise 1004, Description:="'DeleteSelectedListObjRows' Error - [selectedRng] Worksheet must match Worksheet containing [lstObj]"
        End If
        '--- --- ---'
        If protOptions > 0 Then
            On Error Resume Next
            UnprotectSheet lstObj.Range.Worksheet, pwd:=CStr(protectionPwd)
            If Err.number <> 0 Then
                protErr = True
                protErrDesc = Err.number & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo -1
            If protErr Then
                Err.Raise 1004, Description:="'DeleteSelectedListObjRows' Error unprotected sheet"
            End If
        End If
        On Error GoTo E:
        '--- --- ---'
        evts = Events
        scrn = Application.ScreenUpdating
        calc = Application.Calculation
        EventsOff
        Screen = False
        Application.Calculation = xlCalculationManual
        '--- --- ---'
        If lstObj.ListRows.Count > 0 Then
            Dim minRowIdx As Long, hdrRow As Long
            Dim delColl As New Collection
            Dim intRng As Range, rngArea As Range, tmpRng As Range, delRng As Range
            hdrRow = lstObj.HeaderRowRange.Row
            Set intRng = Intersect(Selection, lstObj.DataBodyRange)
            If Not intRng Is Nothing Then
                For Each rngArea In intRng.Areas
                    For Each tmpRng In rngArea.Resize(ColumnSize:=1)
                        If minRowIdx = 0 Or tmpRng.Row - hdrRow < minRowIdx Then
                            minRowIdx = tmpRng.Row - hdrRow
                        End If
                        If Not CollectionKeyExists(delColl, CStr(tmpRng.Row - hdrRow)) Then
                            delColl.Add lstObj.ListRows(tmpRng.Row - hdrRow).Range, key:=CStr(tmpRng.Row - hdrRow)
                        End If
                    Next tmpRng
                Next rngArea
            End If
        End If
        Set delRng = Nothing
        If delColl.Count > 1 And canDeleteMultiple = False Then
             Set delRng = lstObj.ListRows(minRowIdx).Range
        ElseIf delColl.Count > 0 Then
            For Each tmpRng In delColl
                If delRng Is Nothing Then
                    Set delRng = tmpRng
                Else
                    Set delRng = Union(delRng, tmpRng)
                End If
            Next tmpRng
        End If
        If Not delRng Is Nothing Then
            Dim doAction As Boolean
            delRng.Select
            Application.ScreenUpdating = True
            If requestConfirmation Then
                If MsgBox_FT("Delete the " & RangeRowCount(delRng) & " selected row(s)?", vbYesNo + vbDefaultButton2 + vbQuestion, "Delete") = vbYes Then
                    doAction = True
                End If
            Else
                doAction = True
            End If
            If doAction Then
                Application.ScreenUpdating = False
                delRng.Delete xlShiftUp
                Selection(1, 1).Select
            End If
        End If
Finalize:
        On Error Resume Next
        '--- --- ---'
        ''If 'protOptions' is > 0, then worksheet was protected and needs to be re-protected
        If protOptions > 0 Then
            ProtectSheet lstObj.Range.Worksheet, options:=protOptions, password:=CStr(protectionPwd)
        End If
        
        Events = evts
        Screen = scrn
        Application.Calculation = calc
        
        
        Exit Function
E:
        failed = True
        '--- --- ---'
        ''Add Preferred Error Handling Here
        MsgBox "An Error Occurred running 'DeleteSelectedListObjRows'"
        Resume Finalize:
    End Function

    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''  Count Unique Rows In Range
    ''  (Works for Ranges with multiple areas)
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function RangeRowCount(ByVal rng As Range) As Long
        If rng Is Nothing Then Exit Function
        Dim rngColl As New Collection
        Dim tArea As Range
        Dim tRow As Range
        For Each tArea In rng.Areas
            For Each tRow In tArea.Resize(ColumnSize:=1)
                If Not CollectionKeyExists(rngColl, CStr(tRow.Row)) Then
                    rngColl.Add tRow, key:=CStr(tRow.Row)
                End If
            Next tRow
        Next tArea
        RangeRowCount = rngColl.Count
        Set rngColl = Nothing
    End Function

    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''  Count Unique Columns In Range
    ''  (Works for Ranges with multiple areas)
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function RangeColumnCount(ByVal rng As Range) As Long
        If rng Is Nothing Then Exit Function
        Dim rngColl As New Collection
        Dim tArea As Range
        Dim tCol As Range
        For Each tArea In rng.Areas
            For Each tCol In tArea.Resize(RowSize:=1)
                If Not CollectionKeyExists(rngColl, CStr(tCol.column)) Then
                    rngColl.Add tCol, key:=CStr(tCol.column)
                End If
            Next tCol
        Next tArea
        RangeColumnCount = rngColl.Count
        Set rngColl = Nothing
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'
'   WORKSHEET PROTECTION
'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Return Default 'SheetProtection' Flag Enum
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Get DefaultProtectOptions() As SheetProtection
    DefaultProtectOptions = _
        psAllowFiltering _
        + psAllowFormattingCells _
        + psAllowFormattingColumns _
        + psAllowFormattingRows _
        + psDrawingObjects _
        + psUserInterfaceOnly _
        + psContents _
        + psAllowSorting
End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  ** CurrentProtectionOptions **
''  Builds and Returns a 'SheetProtection' enum based on
''      current protection options of a protected worksheet
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function CurrentProtectionOptions(lstObj As ListObject) As SheetProtection

    Dim protectDrawingObjects As Boolean
    Dim protectContents As Boolean
    Dim protectScenarios As Boolean
    Dim userInterfaceOnly As Boolean
    Dim allowFormattingCells As Boolean
    Dim allowFormattingColumns As Boolean
    Dim allowFormattingRows As Boolean
    Dim allowInsertingColumns As Boolean
    Dim allowInsertingRows As Boolean
    Dim allowInsertingHyperlinks As Boolean
    Dim allowDeletingColumns As Boolean
    Dim allowDeletingRows As Boolean
    Dim allowSorting As Boolean
    Dim allowFiltering As Boolean
    Dim allowUsingPivotTables As Boolean
    
    Dim spEnum As SheetProtection
    With lstObj.Range.Worksheet
        If .protectContents Then
            spEnum = EnumModify(spEnum, SheetProtection.psUserInterfaceOnly, feVerifyEnumExists)
            If .protectDrawingObjects Then spEnum = EnumModify(spEnum, SheetProtection.psDrawingObjects, feVerifyEnumExists)
            If .protectContents Then spEnum = EnumModify(spEnum, SheetProtection.psContents, feVerifyEnumExists)
            If .protectScenarios Then spEnum = EnumModify(spEnum, SheetProtection.psScenarios, feVerifyEnumExists)
            
            If .Protection.allowFormattingCells Then spEnum = EnumModify(spEnum, SheetProtection.psAllowFormattingCells, feVerifyEnumExists)
            If .Protection.allowFormattingColumns Then spEnum = EnumModify(spEnum, SheetProtection.psAllowFormattingColumns, feVerifyEnumExists)
            If .Protection.allowFormattingRows Then spEnum = EnumModify(spEnum, SheetProtection.psAllowFormattingRows, feVerifyEnumExists)
            If .Protection.allowInsertingColumns Then spEnum = EnumModify(spEnum, SheetProtection.psAllowInsertingColumns, feVerifyEnumExists)
            If .Protection.allowInsertingRows Then spEnum = EnumModify(spEnum, SheetProtection.psAllowInsertingRows, feVerifyEnumExists)
            If .Protection.allowInsertingHyperlinks Then spEnum = EnumModify(spEnum, SheetProtection.psAllowInsertingHyperlinks, feVerifyEnumExists)
            If .Protection.allowDeletingColumns Then spEnum = EnumModify(spEnum, SheetProtection.psAllowDeletingColumns, feVerifyEnumExists)
            If .Protection.allowDeletingRows Then spEnum = EnumModify(spEnum, SheetProtection.psAllowDeletingRows, feVerifyEnumExists)
            If .Protection.allowSorting Then spEnum = EnumModify(spEnum, SheetProtection.psAllowSorting, feVerifyEnumExists)
            If .Protection.allowFiltering Then spEnum = EnumModify(spEnum, SheetProtection.psAllowFiltering, feVerifyEnumExists)
            If .Protection.allowUsingPivotTables Then spEnum = EnumModify(spEnum, SheetProtection.psAllowUsingPivotTables, feVerifyEnumExists)
        End If
    End With

    CurrentProtectionOptions = spEnum

End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  ** UnprotectSheet **
''  Will use CFG_PROTECT_PASSWORD Constant if separate
''  Password is not supplied, and if a cached password
''  does not exist
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function UnprotectSheet(ByRef wksht As Worksheet _
    , Optional pwd As String = CFG_PROTECT_PASSWORD)
    '   If a cached configuration exists for sheet, use that password instead of default
        If wksht.protectContents Then
            If HasCachedProtectOptions(wksht) Then
                Dim cachedOpt() As Variant
                cachedOpt = CachedProtectionOptions(wksht)
                pwd = CStr(cachedOpt(CACHED_PROTECT_PWD_POSITION))
            End If
            wksht.Unprotect password:=pwd
        End If
End Function
Public Function UnprotectAllSheets(Optional wkbk As Workbook, Optional unhideAll As Boolean = False)
    If wkbk Is Nothing Then
        Set wkbk = ThisWorkbook
    End If
    Dim tWS As Worksheet
    For Each tWS In wkbk.Worksheets
        UnprotectSheet tWS
        If Not tWS.Visible = xlSheetVisible And unhideAll Then
            tWS.Visible = xlSheetVisible
        End If
    Next tWS
End Function
Public Function AddSheetProtectionCache(ByRef wksht As Worksheet, pwd As String, protectOptions As SheetProtection, Optional allowReplace As Boolean = False)
    Dim colKey As String, Item() As Variant, existKey As Boolean
    colKey = KeyCachedProtection(wksht)
    existKey = CollectionKeyExists(protSettingCache, colKey)
    If existKey And allowReplace = False Then
        Err.Raise 457
    ElseIf existKey Then
        protSettingCache.Remove (colKey)
    End If
    protSettingCache.Add Array(pwd, protectOptions), key:=colKey
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Returns Arraycached SheetProtection options, otherwise '0' (zero)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function CachedProtectionOptions(ByRef wksht As Worksheet) As Variant
    Dim Item() As Variant, tKey
    tKey = KeyCachedProtection(wksht)
    If CollectionKeyExists(protSettingCache, tKey) Then
        CachedProtectionOptions = protSettingCache(tKey)
    Else
        CachedProtectionOptions = CVErr(1004)
    End If
End Function
Private Function HasCachedProtectOptions(ByRef wksht As Worksheet) As Boolean
    If CollectionKeyExists(protSettingCache, KeyCachedProtection(wksht)) Then
        HasCachedProtectOptions = True
    End If
End Function

Private Function KeyCachedProtection(ByRef wksht As Worksheet) As String
    KeyCachedProtection = ConcatWithDelim(".", wksht.Parent.Name, wksht.CodeName)
End Function

'   ** ProtectSheet **
'       - Will use CFG_PROTECT_PASSWORD Constant if separate
'       Password is not supplied
'       - If options is not 0 (not set) or default, and
'       'enableOptionsCache' is True, then SheetProtection
'       options will be saved for [wksht] and will be used
'       next time protect is called for [wksht] without
'       any custom protect options
Public Function ProtectSheet(ByRef wksht As Worksheet _
    , Optional options As SheetProtection _
    , Optional password = CFG_PROTECT_PASSWORD _
    , Optional enableOptionsCache As Boolean = True)
    
    Dim protectDrawingObjects As Boolean
    Dim protectContents As Boolean
    Dim protectScenarios As Boolean
    Dim userInterfaceOnly As Boolean
    Dim allowFormattingCells As Boolean
    Dim allowFormattingColumns As Boolean
    Dim allowFormattingRows As Boolean
    Dim allowInsertingColumns As Boolean
    Dim allowInsertingRows As Boolean
    Dim allowInsertingHyperlinks As Boolean
    Dim allowDeletingColumns As Boolean
    Dim allowDeletingRows As Boolean
    Dim allowSorting As Boolean
    Dim allowFiltering As Boolean
    Dim allowUsingPivotTables As Boolean

    If options = 0 Then
        If HasCachedProtectOptions(wksht) Then
           Dim cOptions() As Variant
           cOptions = CachedProtectionOptions(wksht)
           password = cOptions(CACHED_PROTECT_PWD_POSITION)
           options = cOptions(CACHED_PROTECT_ENUM_POSITION)
        Else
            options = DefaultProtectOptions
        End If
    End If
    If options <> DefaultProtectOptions And enableOptionsCache = True Then
        AddSheetProtectionCache wksht, CStr(password), options, allowReplace:=True
    End If
    
    protectDrawingObjects = EnumCompare(options, SheetProtection.psDrawingObjects)
    protectContents = EnumCompare(options, SheetProtection.psContents)
    protectScenarios = EnumCompare(options, SheetProtection.psScenarios)
    userInterfaceOnly = EnumCompare(options, SheetProtection.psUserInterfaceOnly)
    allowFormattingCells = EnumCompare(options, SheetProtection.psAllowFormattingCells)
    allowFormattingColumns = EnumCompare(options, SheetProtection.psAllowFormattingColumns)
    allowFormattingRows = EnumCompare(options, SheetProtection.psAllowFormattingRows)
    allowInsertingColumns = EnumCompare(options, SheetProtection.psAllowInsertingColumns)
    allowInsertingRows = EnumCompare(options, SheetProtection.psAllowInsertingRows)
    allowInsertingHyperlinks = EnumCompare(options, SheetProtection.psAllowInsertingHyperlinks)
    allowDeletingColumns = EnumCompare(options, SheetProtection.psAllowDeletingColumns)
    allowDeletingRows = EnumCompare(options, SheetProtection.psAllowDeletingRows)
    allowSorting = EnumCompare(options, SheetProtection.psAllowSorting)
    allowFiltering = EnumCompare(options, SheetProtection.psAllowFiltering)
    allowUsingPivotTables = EnumCompare(options, SheetProtection.psAllowUsingPivotTables)

    ProtectSheet = ProtectSheet2(wksht, password, protectDrawingObjects, _
        protectContents, protectScenarios, userInterfaceOnly, allowFormattingCells, _
        allowFormattingColumns, allowFormattingRows, allowInsertingColumns, _
        allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, _
        allowDeletingRows, allowSorting, allowFiltering, allowUsingPivotTables)

End Function

Private Function ProtectSheet2(ByRef wksht As Worksheet _
    , Optional password = CFG_PROTECT_PASSWORD _
    , Optional protectDrawingObjects As Boolean = True _
    , Optional protectContents As Boolean = True _
    , Optional protectScenarios As Boolean = False _
    , Optional userInterfaceOnly As Boolean = True _
    , Optional allowFormattingCells As Boolean = True _
    , Optional allowFormattingColumns As Boolean = True _
    , Optional allowFormattingRows As Boolean = True _
    , Optional allowInsertingColumns As Boolean = False _
    , Optional allowInsertingRows As Boolean = False _
    , Optional allowInsertingHyperlinks As Boolean = False _
    , Optional allowDeletingColumns As Boolean = False _
    , Optional allowDeletingRows As Boolean = False _
    , Optional allowSorting As Boolean = True _
    , Optional allowFiltering As Boolean = True _
    , Optional allowUsingPivotTables As Boolean = False)

    With wksht
       .Protect password:=password _
        , DrawingObjects:=protectDrawingObjects _
        , Contents:=protectContents _
        , Scenarios:=protectScenarios _
        , userInterfaceOnly:=userInterfaceOnly _
        , allowFormattingCells:=allowFormattingCells _
        , allowFormattingColumns:=allowFormattingColumns _
        , allowFormattingRows:=allowFormattingRows _
        , allowInsertingColumns:=allowInsertingColumns _
        , allowInsertingRows:=allowInsertingRows _
        , allowInsertingHyperlinks:=allowInsertingHyperlinks _
        , allowDeletingColumns:=allowDeletingColumns _
        , allowDeletingRows:=allowDeletingRows _
        , allowSorting:=allowSorting _
        , allowFiltering:=allowFiltering _
        , allowUsingPivotTables:=allowUsingPivotTables
    End With
End Function

Public Function RemoveGridLines(Optional wkbkOrWksht)
    On Error Resume Next
    Dim tmpWkbk As Workbook
    Dim tmpWksht As Worksheet
    Dim view As WorksheetView
    If StringsMatch(TypeName(wkbkOrWksht), "Workbook") Then
        Set tmpWkbk = wkbkOrWksht
    ElseIf StringsMatch(TypeName(wkbkOrWksht), "Worksheet") Then
        Set tmpWksht = wkbkOrWksht
    End If
    If tmpWkbk Is Nothing And tmpWksht Is Nothing Then
        Set tmpWkbk = ThisWorkbook
    End If
    If Not tmpWkbk Is Nothing Then
        For Each view In tmpWkbk.Windows(1).SheetViews
            If view.DisplayGridlines = True Then
                view.DisplayGridlines = False
            End If
        Next
    ElseIf Not tmpWksht Is Nothing Then
        For Each view In tmpWksht.Parent.Windows(1).SheetViews
            If view.Sheet Is tmpWksht Then
                view.DisplayGridlines = False
                Exit For
            End If
        Next
    End If
    Set tmpWkbk = Nothing
    Set tmpWksht = Nothing
    If Err.number <> 0 Then Err.Clear
End Function



    Public Function StringsMatch( _
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

Public Function ConcatWithDelim(ByVal delimeter As String, ParamArray items() As Variant) As String
    ConcatWithDelim = Join(items, delimeter)
End Function

'RETURN STRING FOR EACH ROW REPRESENTED IN RANGE, vbNewLine as Line Delimeter
'   Example 1 (Get your Column Names for a list object)
'       Dim lo as ListObject
'       Set lo = wsTeamInfo.ListOobjects("tblTeamInfo")
'       Debug.Print ConcatRange(lo.HeaderRowRange)
'           outputs:  StartDt|EndDt|Project|Employee|Role|BillRate|EstCostRt|ActCostRt|Active|TaskName|SegName|AllocPerc|Utilization|Bill_Hrs|NonBill_Hrs|CfgID|ActiveHidden|Updated
'   Example 2 (let's grab some weird ranges)
'       Dim rng As Range
'       Set rng = wsDashboard.Range("E49:J50")
'       Set rng = Union(rng, wsDashboard.Range("L60:Q60"))
'       Debug.Print ConcatRange(rng)
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
                    retV = ConcatWithDelim("", retV, rng.Areas(rArea)(rRow, rCol).Value)
                Else
                    retV = ConcatWithDelim(delimeter, retV, rng.Areas(rArea)(rRow, rCol).Value)
                End If
            Next rCol
        Next rRow
    Next rArea
    ConcatRange = retV
End Function
Public Function Concat(ParamArray items() As Variant) As String
    Concat = Join(items, "")
End Function
Public Function Concat_1DArray(dArr, Optional delimiter As String = " | ")
On Error Resume Next
    Concat_1DArray = Join(dArr, delimiter)
End Function

Public Property Get ENV_User() As String
    ENV_User = VBA.Interaction.Environ("USER")
End Property

Public Function ENV_LogName() As String
    #If Mac Then
        ENV_LogName = VBA.Interaction.Environ("LOGNAME")
    #Else
        ENV_LogName = VBA.Interaction.Environ("USERNAME")
    #End If
End Function

Public Property Get DBLQUOTE() As String
    DBLQUOTE = Chr(34)
End Property

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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   FLAG ENUM COMPARE
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function EnumCompare(theEnum As Variant, enumMember As Variant, Optional ByVal iType As ecComparisonType = ecComparisonType.ecOR) As Boolean
    Dim c As Long
    c = theEnum And enumMember
    EnumCompare = IIf(iType = ecOR, c <> 0, c = enumMember)
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   FLAG ENUM - ADD/REMOVE SPECIFIC ENUM MEMBER
'   (Works with any flag enum)
'   e.g. If you have vbMsgBoxStyle enum and want to make sure
'   'DefaultButton1' is included
'   msgBtnOption = vbYesNo + vbQuestion
'   msgBtnOption = EnumModify(msgBtnOption,vbDefaultButton1,feVerifyEnumExists)
'   'now includes vbDefaultButton1, would not modify enum value if it already
'   contained vbDefaultButton1
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function EnumModify(theEnum, enumMember, modifyType As FlagEnumModify) As Long
    Dim Exists As Boolean
    Exists = EnumCompare(theEnum, enumMember)
    If Exists And modifyType = feVerifyEnumRemoved Then
        theEnum = theEnum - enumMember
    ElseIf Exists = False And modifyType = feVerifyEnumExists Then
        theEnum = theEnum + enumMember
    End If
    EnumModify = theEnum
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   MESSAGE BOX REPLACEMENT
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function MsgBox_FT(prompt As String, Optional buttons As VbMsgBoxStyle = vbOKOnly, Optional title As Variant) As Variant
    Dim evts As Boolean: evts = Events
    Dim screenUpd As Boolean: screenUpd = Application.ScreenUpdating
    EventsOff
    If Not EnumCompare(buttons, vbSystemModal) Then buttons = buttons + vbSystemModal
    If Not EnumCompare(buttons, vbMsgBoxSetForeground) Then buttons = buttons + vbMsgBoxSetForeground
    If EnumCompare(buttons, vbOKOnly) Then
        Beep
    End If
    If Not ThisWorkbook.activeSheet Is Application.activeSheet Then
        Application.ScreenUpdating = True
        ThisWorkbook.Activate
        DoEvents
        Application.ScreenUpdating = screenUpd
    End If
    MsgBox_FT = MsgBox(prompt, buttons, title)
    Events = evts
    DoEvents
End Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   ASK YES/NO
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   GET NEXT ID
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function GetNextID(table As ListObject, uniqueIdcolumnIdx As Long) As Long
'   Use to create next (Long) number for unique ROW id in a Range
On Error Resume Next
    Dim nextID As Long
    If table.ListRows.Count > 0 Then
        nextID = Application.WorksheetFunction.Max(table.ListColumns(uniqueIdcolumnIdx).DataBodyRange)
    End If
    GetNextID = nextID + 1
    If Err.number <> 0 Then Err.Clear
End Function

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
Public Property Get Screen() As Boolean
    Screen = Application.ScreenUpdating
End Property
Public Property Let Screen(vl As Boolean)
    Application.ScreenUpdating = vl
    Application.Interactive = vl
    Application.Cursor = IIf(vl, xlDefault, xlWait)
End Property
Public Function ScreenOn()
    Screen = True
End Function
Public Function ScreenOff()
    Screen = False
End Function

Public Property Get VisibleWorksheets() As Long
    Dim tIDX As Long, retV As Long
    For tIDX = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Worksheets(tIDX).Visible = xlSheetVisible Then
            retV = retV + 1
        End If
    Next tIDX
    VisibleWorksheets = retV
End Property



''Public Function TestCollectionItemByKey()
''    Dim c As New Collection, i
''    For i = 1 To 10
''        c.Add i, key:="K" & i
''    Next i
''    Debug.Assert CollectionItemByKey(c, "K5") = 5
''    Debug.Assert IsError(CollectionItemByKey(c, "K11"))
''End Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Returns item from collection by Key
'   If [key] does not exist in collection, error object with
'   error code 1004 is return
'   suggested use:
'
'   Dim colItem as Variant
'   colItem = CollectionItemByKey([collection], [expectedKey])
'
'   'If expecting object, use 'Set'
'    Set colItem = CollectionItemByKey([collection], [expectedKey])
'
'   If Not IsError(colItem) Then
'       'value was returned
'   Else
'       'error was returned
'   End if
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function CollectionItemByKey(ByRef col As Collection, key)
On Error Resume Next
    If IsObject(col(key)) Then
        If Err.number = 0 Then
            Set CollectionItemByKey = col(key)
        End If
    Else
        If Err.number = 0 Then
            CollectionItemByKey = col(key)
        End If
    End If
    If Err.number <> 0 Then
        Err.Clear
        CollectionItemByKey = CVErr(1004)
    End If
End Function
Public Function CollectionKeyExists(ByRef col As Collection, key)
On Error Resume Next
    If IsObject(col(key)) Then
        If Err.number = 0 Then
            CollectionKeyExists = True
        Else
            CollectionKeyExists = False
        End If
    Else
        If Err.number = 0 Then
            CollectionKeyExists = True
        Else
            CollectionKeyExists = False
        End If
    End If
    If Err.number <> 0 Then
        Err.Clear
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Get All  'Environ' Settings, or
'   Get 'Environ' Settings matches or partially matches 'filter'
'   filter applied to [Key] by default, to apply to value(s), use
'       filterKey:=False
'   Returns Collection of 1-based Arrays (1=Key, 2=Val) with [Key] = Environ Key, and
'       [Item] = Environ Value
'   If using filterKey with exactMatch:=True, will return 1 or no matching items
'   When mutiple items are matched (exactMatch:=False),
'       will return 0 to Many items
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   FYI: TO GET USERNAME:
'       ON PC USE:  =GetEnvironSettings("username")
'       ON MAC USE: =GetEnvironSettings("user")
'
'   (VBA.Interaction.Environ([KEY] is case sensitive, however using this
'       'GetEnvironSettings' is not case sensitive)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   EXAMPLE OF USAGE
'       ** Print all Environ Key/Value values to Immediate Window
'       Dim v
'       For Each v In GetEnvironSettings
'           Debug.Print v(1), v(2) 'Key, Value
'       Next v
'       ** Print all Environ Key/Value values
'       Dim v
'       For Each v In GetEnvironSettings
'           Debug.Print v(1) & ", " & v(2)
'       Next v
'       ** Get Environ setting value for key=USER
'       Dim v, tempCol as Collection
'       Set tempCol = GetEnvironSettings("user", exactMatch:=True)
'           THIS
'       If tempCol.Count = 1 Then
'           Debug.Print tempCol(1)(2)
'       End if
'           OR THIS
'       v = tempCol(1)
'           Debug.Print v(2)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function GetEnvironSettings( _
    Optional Filter As String, _
    Optional exactMatch As Boolean = True, _
    Optional filterKey As Boolean = True) As Collection

    Dim tColl As New Collection, arrItem, eq As String, eqPosition As Long
    Dim EnvString, i As Long, tKey, tVal, tFilterItem, canReturn As Boolean
    i = 1
    eq = "="
    EnvString = Environ(i)
    Do While EnvString <> vbNullString
        canReturn = False
        eqPosition = InStr(EnvString, eq)
        If eqPosition > 0 Then
            tKey = Left(EnvString, (eqPosition - 1))
            tVal = Mid(EnvString, (eqPosition + 1))
            If Not StringsMatch(Filter, vbNullString) Then
                tFilterItem = IIf(filterKey, tKey, tVal)
                If exactMatch And StringsMatch(tFilterItem, Filter, smEqual) Then
                    canReturn = True
                ElseIf Not exactMatch And StringsMatch(tFilterItem, Filter, smContains) Then
                    canReturn = True
                End If
            Else
                canReturn = True
            End If
        End If
        If canReturn Then
            tColl.Add Array(tKey, tVal), key:=tKey
        End If
        
        i = i + 1
        EnvString = Environ(i)
    Loop
    
    Set GetEnvironSettings = tColl
    Set tColl = Nothing

End Function




' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Get any list object in workbook by name
'   by default, will cache reference to list object
'   set [cacheReference] to False if looking for a list object
'   that is known to be temporary
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function wt(lstObjName As String, Optional cacheReference As Boolean = True) As ListObject
    Static loCol As New Collection
    If cacheReference And CollectionKeyExists(loCol, lstObjName) Then
        Set wt = loCol(lstObjName)
    Else
        Dim tmpWS As Worksheet, tmpLO As ListObject
        For Each tmpWS In ThisWorkbook.Worksheets
            For Each tmpLO In tmpWS.ListObjects
                If StringsMatch(tmpLO.Name, lstObjName) Then
                    If cacheReference Then
                        loCol.Add tmpLO, key:=lstObjName
                    End If
                    Set wt = tmpLO
                End If
                If Not wt Is Nothing Then Exit For
            Next tmpLO
            If Not wt Is Nothing Then Exit For
        Next tmpWS
    End If
End Function


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'   Scroll any active sheet to desired location
'    - Does not change previous worksheet selection
'    - Optionally set selection range, if desired ('selectRng')
'
'   Can use for scrolling only, worksheets do not have to have split panes
'
'   Use 'splitOnRow' and/or 'splitOnColumn' to guarantee split is correct
'    - By default split panes will be frozen.  Pass in arrgument: 'freezePanes:=False'
'      to make sure split panes are not frozen
'
'   By Default, if a splitRow/Column is not specific, but one existrs, it will be
'   left alone.  To remove split panes that should not exist by default,
'   pass in 'removeUnspecified:=True'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Public Function Scroll(wksht As Worksheet _
    , Optional splitOnRow As Long _
    , Optional splitOnColumn As Long _
    , Optional freezePanes As Boolean = True _
    , Optional removeUnspecified As Boolean _
    , Optional selectRng As Range)
    ' --- '
    On Error GoTo E:
    'The Worksheet you are scrolling must be the ActiveSheet'
    If Not ActiveWindow.activeSheet Is wksht Then Exit Function
    ' --- '
    Dim failed As Boolean
    Dim evts As Boolean, scrn As Boolean, scrn2 As Boolean
    evts = Application.EnableEvents
    scrn = Application.ScreenUpdating
    scrn2 = Application.Interactive
    ' --- '
    Dim pnIdx As Long
    With ActiveWindow
        'Scroll All Panes to the left, to the top'

        For pnIdx = 1 To .Panes.Count
'            .SmallScroll ToRight:=1
'            .SmallScroll Down:=1
            If pnIdx = 1 Or pnIdx = 3 Then
                .Panes(pnIdx).ScrollRow = 1
            ElseIf .splitRow > 0 Then
                .Panes(pnIdx).ScrollRow = .splitRow + 1
            End If
            If pnIdx = 1 Or pnIdx = 2 Then
                .Panes(pnIdx).ScrollColumn = 1
            ElseIf .splitColumn > 0 Then
                .Panes(pnIdx).ScrollColumn = .splitColumn + 1
            End If
        Next pnIdx
        'Ensure split panes are in the right place
        If splitOnRow > 0 And Not .splitRow = splitOnRow Then
            .splitRow = splitOnRow
        ElseIf splitOnRow = 0 And .splitRow <> 0 And removeUnspecified Then
            .splitRow = 0
        End If
        If splitOnColumn > 0 And Not .splitColumn = splitOnColumn Then
            .splitColumn = splitOnColumn
        ElseIf splitOnColumn = 0 And .splitColumn <> 0 And removeUnspecified Then
            .splitColumn = 0
        End If
        If splitOnColumn > 0 Or splitOnRow > 0 Then
            If Not .freezePanes = freezePanes Then
                .freezePanes = freezePanes
            End If
        End If
    End With
    If Not selectRng Is Nothing Then
        If selectRng.Worksheet Is wksht Then
            selectRng.Select
        End If
    End If
Finalize:
    On Error Resume Next
    Application.EnableEvents = evts
    Application.ScreenUpdating = scrn
    Application.Interactive = scrn2
    Application.StatusBar = False
    Exit Function
E:
    'Implement Own Error Handling'
    failed = True
    MsgBox "Error in 'Scroll' Function: " & Err.number & " - " & Err.Description
    Err.Clear
    Resume Finalize:
End Function





' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'
'       UNCATEGORIZED
'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function ShowAllSheets(Optional wkbk As Workbook)
    If wkbk Is Nothing Then
        UnhideSheets ThisWorkbook
    Else
        UnhideSheets wkbk
    End If
End Function
Public Function UnhideSheets(wkbk As Workbook, ParamArray keepHiddenSheets() As Variant)
    Dim evts, scrn
    evts = Events
    scrn = Screen
    EventsOff
    ScreenOff
    
    Dim ws As Worksheet, hideColl As New Collection
    If UBound(keepHiddenSheets) >= LBound(keepHiddenSheets) Then
        Dim sht
        For Each sht In keepHiddenSheets
            If StringsMatch(TypeName(sht), "Worksheet") Then
                hideColl.Add sht.CodeName, key:=sht.CodeName
            ElseIf StringsMatch(TypeName(sht), "String") Then
                For Each ws In wkbk.Worksheets
                    If StringsMatch(ws.CodeName, sht) Or StringsMatch(ws.Name, sht) Then
                        hideColl.Add ws.CodeName, key:=ws.CodeName
                    End If
                Next ws
            End If
        Next sht
    End If
    
    Dim activeSheet As Worksheet
    Set activeSheet = wkbk.activeSheet
    For Each ws In wkbk.Worksheets
        If Not ws.Visible = xlSheetVisible Then
            If Not CollectionKeyExists(hideColl, ws.CodeName) Then
                ws.Visible = xlSheetVisible
            End If
        End If
    Next ws
    If hideColl.Count > 0 Then
        For Each ws In wkbk.Worksheets
            If CollectionKeyExists(hideColl, ws.CodeName) Then
                If ws.Visible = xlSheetVisible Then
                    ws.Visible = xlSheetVeryHidden
                End If
            End If
        Next ws
    End If
    
    Events = evts
    Screen = scrn
End Function


