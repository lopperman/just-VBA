Attribute VB_Name = "pbCommonUtil"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  Common Methods & Utililities
'  Most Modules/Classes in just-VBA Library are dependent
'  on this common module
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbCommonUtil.bas
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   used to cache protection settings for specific worksheet
'   Key=WorkbookName.WorksheetCodeName
'   Value = 1-based Array:
'   (1)=password,(2)=SheetProtection enum
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private protSettingCache As New Collection

Private l_IsDeveloper As Boolean

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   GENERALIZED CONSTANTS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Const CFG_PROTECT_PASSWORD As String = "00000"
Public Const CFG_PROTECT_PASSWORD_EXPORT As String = "000001"
Public Const CFG_PROTECT_PASSWORD_MISC As String = "0000015"
Public Const CFG_PROTECT_PASSWORD_VBA = "0123210"
Public Const TEMP_DIRECTORY_NAME2 As String = "VBATemp"

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   CUSTOM ERROR CONSTANTS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   *** GENERIC ERRORS ONLY
'   *** STARTS AT 1001 - (513-1000 RESERVED FOR APPLICATION SPECIFIC ERRORS)
Public Const ERR_OBSOLETE As Long = vbObjectError + 1001
Public Const ERR_MISSING_EXPECTED_WORKSHEET As Long = vbObjectError + 1002
Public Const ERR_HOME_SHEET_NOT_SET As Long = vbObjectError + 1003
Public Const ERR_HOME_SHEET_ALREADY_SET As Long = vbObjectError + 1004
Public Const ERR_INVALID_CALLER_SOURCE As Long = vbObjectError + 1005
Public Const ERR_ACTION_PATH_NOT_DEFINED As Long = vbObjectError + 1006
Public Const ERR_UNABLE_TO_GET_WORKSHEET_BUTTON As Long = vbObjectError + 1007
Public Const ERR_RANGE_SRC_TARGET_MISMATCH = vbObjectError + 1008
Public Const ERR_INVALID_FT_ARGUMENTS = vbObjectError + 1009
Public Const ERR_NO_ERROR As String = "(INFO)"
Public Const ERR_ERROR As String = "(ERROR)"
Public Const ERR_INVALID_SETTING_OPERATION As Long = vbObjectError + 1010
Public Const ERR_TARGET_OBJECT_MUST_BE_EMPTY As Long = vbObjectError + 1011
Public Const ERR_EXPECTED_SHEET_NOT_ACTIVE As Long = vbObjectError + 1012
Public Const ERR_INVALID_RANGE_SIZE = vbObjectError + 1013
Public Const ERR_IMPORT_COLUMN_ALREADY_MATCHED = vbObjectError + 1014
Public Const ERR_RANGE_AREA_COUNT = vbObjectError + 1015
Public Const ERR_LIST_OBJECT_RESIZE_CANNOT_DELETE = vbObjectError + 1016
Public Const ERR_LIST_OBJECT_RESIZE_INVALID_ARGUMENTS = vbObjectError + 1017
Public Const ERR_INVALID_ARRAY_SIZE = vbObjectError + 1018
Public Const ERR_EXPECTED_MANUAL_CALC_NOT_FOUND = vbObjectError + 1019
Public Const ERR_WORKSHEET_OBJECT_NOT_FOUND = vbObjectError + 1020
Public Const ERR_NOT_IMPLEMENTED_YET = vbObjectError + 1021
'CLASS INSTANCE PROHIBITED, FOR CLASS MODULE WITH Attribute VB_PredeclaredId = True
Public Const ERR_CLASS_INSTANCE_PROHIBITED = vbObjectError + 1022
Public Const ERR_CONTROL_STATE = vbObjectError + 1023
Public Const ERR_PREVIOUS_PerfState_EXISTS = vbObjectError + 1024
Public Const ERR_FORCED_WORKSHEET_NOT_FOUND = vbObjectError + 1025
Public Const ERR_CANNOT_CHANGE_PACKAGE_PROPERTY = vbObjectError + 1026
Public Const ERR_INVALID_PACKAGE_OPERATION = vbObjectError + 1027

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   GENERALIZED TYPES
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Type KVP
  key As String
  Value As Variant
End Type
    
Public Type ftFound
    matchExactFirstIDX As Long
    matchExactLastIDX As Long
    matchSmallerIDX As Long
    matchLargerIDX As Long
    realRowFirst As Long
    realRowLast As Long
    realRowSmaller As Long
    realRowLarger As Long
End Type
    
Public Type LocationStart
    Left As Long
    Top As Long
End Type
    
Public Type ArrInformation
    Rows As Long
    Columns As Long
    Dimensions As Long
    Ubound_first As Long
    LBound_first As Long
    UBound_second As Long
    LBound_second As Long
    IsArray As Boolean
End Type
    
Public Type AreaStruct
    rowStart As Long
    RowEnd As Long
    ColStart As Long
    ColEnd As Long
    rowCount As Long
    columnCount As Long
End Type
    
Public Type RngInfo
    Rows As Long
    Columns As Long
    AreasSameRows As Boolean
    AreasSameColumns As Boolean
    Areas As Long
End Type

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
    
Public Enum NullableBool
    [_Default] = 0
    triNULL = 0
    triTRUE = 1
    triFALSE = 2
End Enum

Public Enum ExtendedBool
    ebTRUE = 2 ^ 0
    ebFALSE = 2 ^ 1
    ebPartial = 2 ^ 2
    ebERROR = 2 ^ 3
    ebNULL = 2 ^ 4
End Enum
    
Public Enum PicklistMode
    plSingle = 0
    plMultiple_MinimumNone = -1
    plMultiple_MinimumOne = 1
End Enum

Public Enum ecComparisonType
    ecOR = 0 'default
    ecAnd
End Enum

Public Enum MergeRangeEnum
    mrDefault_MergeAll = 0
    mrUnprotect = 2 ^ 0
    mrClearFormatting = 2 ^ 1
    mrClearContents = 2 ^ 2
    mrMergeAcrossOnly = 2 ^ 3
End Enum

Public Enum InitActionEnum
    [_DefaultInvalid] = 0
    iaAutoCode
    iaEventResponse
    iaButtonClick
    iaManual
End Enum

Public Enum ReportPeriod
    frpDay = 1
    frpWeek = 2
    frpGLPeriod = 3
    frpCalMonth = 4
End Enum
    
Public Enum strMatchEnum
    smEqual = 0
    smNotEqualTo = 1
    smContains = 2
    smStartsWithStr = 3
    smEndWithStr = 4
End Enum

Public Enum ftOperatingState
    [_ftunknown] = -1
    ftOpening = 0
    ftRunning = 1
    ftClosing = 2
    ftUpgrading = 3
    ftResetting = 4
    ftImporting = 5
End Enum
    
    Public Enum ProtectionTemplate
        ptDefault = 0
        ptAllowFilterSort = 1
        ptDenyFilterSort = 2
        ptCustom = 3
    End Enum
    
    Public Enum protectionPwd
        pwStandard = 1
        pwExport = 2
        pwMisc = 3
        pwVBAOnly = 4
    End Enum
    
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
    
    Public Enum RangeFunctionOperator
        Min = 1
        Max = 2
        Sum = 3
        Count = 4
        CountUnique = 5
        CountBlank = 6
    End Enum
    
    Public Enum btnLocationEnum
        Beneath = 1
        ToTheRight
    End Enum
    
    Public Enum color
        Aqua = 42
        Black = 1
        Blue = 5
        BlueGray = 47
        BrightGreen = 4
        Brown = 53
        cream = 19
        DarkBlue = 11
        DarkGreen = 51
        DarkPurple = 21
        DarkRed = 9
        DarkTeal = 49
        DarkYellow = 12
        Gold = 44
        Gray25 = 15
        Gray40 = 48
        Gray50 = 16
        Gray80 = 56
        Green = 10
        Indigo = 55
        Lavender = 39
        LightBlue = 41
        LIGHtgreen = 35
        LightLavender = 24
        LightOrange = 45
        LightTurquoise = 20
        LightYellow = 36
        Lime = 43
        NavyBlue = 23
        OliveGreen = 52
        Orange = 46
        PaleBlue = 37
        Pink = 7
        Plum = 18
        PowderBlue = 17
        red = 3
        Rose = 38
        SALMON = 22
        SeaGreen = 50
        SkyBlue = 33
        Tan = 40
        Teal = 14
        Turquoise = 8
        Violet = 13
        White = 2
        Yellow = 6
    End Enum

    Public Enum ftInputBoxType
    '   DO NOT EDIT ENUM VALUES
    '   THESE TIE TO SPECIFIC
    '   INPUTBOX TYPES
        ftibFormula = 0
        ftibNumber = 1
        ftibString = 2
        ftibLogicalValue = 4
        ftibCellReference = 8
        ftibErrorValue = 16
        ftibArrayOfValues = 64
    End Enum
    
    Public Enum ListReturnType
        lrtArray = 1
        lrtDictionary = 2
        lrtCollection = 3
    End Enum
    
    Public Enum XMatchMode
    '   DO NOT EDIT ENUM VALUES
        exactMatch = 0
        ExactMatchOrNextSmaller = -1
        ExactMatchOrNextLarger = 1
        WildcardCharacterMatch = 2
    End Enum
    
    Public Enum XSearchMode
    '   DO NOT EDIT ENUM VALUES
        searchFirstToLast = 1
        searchLastToFirst = -1
        searchBinaryAsc = 2
        searchBinaryDesc = -2
    End Enum
    
    Public Enum ftActionType
        ftaADD = 1
        ftaEDIT
        ftaDELETE
    End Enum
    
    Public Enum MatchTypeEnum
        mtAll = 1
        mtAny = 2
        mtNONE = 3
    End Enum

    Public Enum ListObjCompareEnum
        locName = 2 ^ 0
        locColumnCount = 2 ^ 1
        locColumnNames = 2 ^ 2
        locColumnOrder = 2 ^ 3
        locRowCount = 2 ^ 4
    End Enum

    Public Enum ArrayOptionFlags
        aoNone = 0
        aoUnique = 2 ^ 0
        aoUniqueNoSort = 2 ^ 1
        aoAreaSizesMustMatch = 2 ^ 2
        aoVisibleRangeOnly = 2 ^ 3
        aoIncludeListObjHeaderRow = 2 ^ 4
    End Enum

    Public Enum ftMinMax
        minValue = 1
        maxValue = 2
    End Enum
    Public Enum HolidayEnum
        holidayName = 1
        holidayDT = 2
    End Enum

    Public Enum BeepType
        btMsgBoxOK = 0
        btMsgBoxChoice = 1
        btError = 2
        btBusyWait = 3
        btButton = 4
        btForced = 5
    End Enum
    
    Public Enum ErrorOptionsEnum
        ftDefaults = 2 ^ 0
        ftERR_ProtectSheet = 2 ^ 1
        ftERR_MessageIgnore = 2 ^ 2
        ftERR_NoBeeper = 2 ^ 3
        ftERR_DoNotCloseBusy = 2 ^ 4
        ftERR_ResponseAllowBreak = 2 ^ 5
    End Enum

    Private lBypassOnCloseCheck As Boolean

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'
'       DATES & TIME UTILITIES
'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

Public Function DtAdd(intervalType As DateDiffType, _
    number As Variant, ByVal dt As Variant) As Variant
    Dim retVal As Variant
    
    Select Case intervalType
        Case DateDiffType.dtDay
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
            DtPart = DateSerial(DtPart(dtYear, dt1), DtPart(dtMonth, dt1), DtPart(dtDay, dt1))
        Case DateDiffType.dtDay
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
    Optional returnFraction As Boolean = False, _
    Optional absoluteVal As Boolean = True) As Variant
    
'       ** absoluteVal (default = True) **
'       When true, will always return a value >= 0
'       When false, will return negative value if [dt2] argument is smaller than [dt1]
'
'       FRACTIONAL RETURN VALUES
'           SUPPORTED FOR: minutes, hours, days, weeks
'       ** No Rounding Occurs When returnFraction = False **
'           e.g.    if return DtDiff(dtDay,...) any value less than 1 will return 0,
'                   any value > 1 and less than 2 will return 1, etc
'       note:  fractionals are based on type of date/time component
'       for example, if the difference in time was 2 minutes, 30 seconds
'       and you were returning Minutes as a fractions, the return value would
'       be 2.5 (for 2 1/2 minutes)
    
    If IsMissing(dt2) Then dt2 = Now
    Dim retVal As Variant
    Dim tmpVal1 As Variant
    Dim tmpVal2 As Variant
    Dim tmpRemain As Variant
    
    Dim recArgFirstDay As VbDayOfWeek, recArgFirstWeek As VbFirstWeekOfYear, recArgFraction As Boolean, recArgAbsolute As Boolean
    recArgFirstDay = firstDayOfWeek
    recArgFirstWeek = firstWeekOfYear
    recArgFraction = returnFraction
    recArgAbsolute = absoluteVal
    
    Select Case diffType
        Case DateDiffType.dtSecond
            retVal = DateDiff("s", dt1, dt2)
        Case DateDiffType.dtWeekday
            retVal = DateDiff("w", dt1, dt2)
        Case DateDiffType.dtMinute
            If returnFraction Then
                ' fractions based on SECONDS (60)
                tmpVal1 = DtDiff(dtSecond, dt1, dt2, firstDayOfWeek:=recArgFirstDay, firstWeekOfYear:=recArgFirstWeek, returnFraction:=recArgFraction, absoluteVal:=recArgAbsolute)
                tmpVal2 = tmpVal1 - (DateDiff("n", dt1, dt2) * 60)
                If tmpVal2 <> 0 Then
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
                tmpVal1 = DtDiff(dtMinute, dt1, dt2, firstDayOfWeek:=recArgFirstDay, firstWeekOfYear:=recArgFirstWeek, returnFraction:=recArgFraction, absoluteVal:=recArgAbsolute)
                tmpVal2 = tmpVal1 - (DateDiff("h", dt1, dt2) * 60)
                If tmpVal2 <> 0 Then
                    retVal = DateDiff("h", dt1, dt2) + (tmpVal2 / 60)
                Else
                    retVal = DateDiff("h", dt1, dt2)
                End If
            Else
                retVal = DateDiff("h", dt1, dt2)
            End If
        Case DateDiffType.dtDay
                ' fractions based on HOURS (24)
            If returnFraction Then
                tmpVal1 = DtDiff(dtHour, dt1, dt2, firstDayOfWeek:=recArgFirstDay, firstWeekOfYear:=recArgFirstWeek, returnFraction:=recArgFraction, absoluteVal:=recArgAbsolute)
                tmpVal2 = tmpVal1 - (DateDiff("d", dt1, dt2) * 24)
                If tmpVal2 <> 0 Then
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
                tmpVal1 = DtDiff(dtDay, dt1, dt2, firstDayOfWeek:=recArgFirstDay, firstWeekOfYear:=recArgFirstWeek, returnFraction:=recArgFraction, absoluteVal:=recArgAbsolute)
                tmpVal2 = tmpVal1 - (DateDiff("ww", dt1, dt2, firstDayOfWeek, firstWeekOfYear) * 7)
                If tmpVal2 <> 0 Then
                    retVal = DateDiff("ww", dt1, dt2, firstDayOfWeek, firstWeekOfYear) + (tmpVal2 / 7)
                Else
                    retVal = DateDiff("ww", dt1, dt2, firstDayOfWeek, firstWeekOfYear)
                End If
            Else
                retVal = DateDiff("ww", dt1, dt2, firstDayOfWeek, firstWeekOfYear)
            End If
        Case DateDiffType.dtMonth
            retVal = DateDiff("m", dt1, dt2)
        Case DateDiffType.dtQuarter
            retVal = DateDiff("q", dt1, dt2, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtYear
            retVal = DateDiff("yyyy", dt1, dt2, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtDayOfYear
            retVal = DateDiff("y", dt1, dt2, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtTimeString
            Dim tHrs As Long, tMin As Long, tSec As Single, tAdj As Single
            tHrs = DtDiff(dtHour, dt1, dt2)
            tMin = DtDiff(dtMinute, dt1, dt2) - (tHrs * 60)
            tSec = DtDiff(dtSecond, dt1, dt2) - (((tHrs * 60) * 60) + (tMin * 60))
            DtDiff = Format(tHrs, "00:") & Format(tMin, "00.") & Format(tSec, "00")
    End Select
    If diffType <> dtTimeString Then
        If absoluteVal Then
            DtDiff = Abs(retVal)
        Else
            DtDiff = retVal
        End If
    End If
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'
'   WORKSHEET PROTECTION
'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

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

'   ** UnprotectSheet **
'   Will use CFG_PROTECT_PASSWORD Constant if separate
'   Password is not supplied
Public Function UnprotectSheet(ByRef wksht As Worksheet _
    , Optional pwd As String = CFG_PROTECT_PASSWORD)
    wksht.Unprotect password:=pwd

End Function

Public Function AddSheetProtectionCache(ByRef wksht As Worksheet, pwd As String, protectOptions As SheetProtection, Optional allowReplace As Boolean = False)
    Dim colKey As String, item() As Variant, existKey As Boolean
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
    Dim item() As Variant, tKey
    tKey = KeyCachedProtection(wksht)
    If CollectionKeyExists(protSettingCache, tKey) Then
        CachedProtectionOptions = protSettingCache(tKey)
    Else
        CachedProtectionOptions = CVErr(1004)
    End If
End Function
Private Function ProtectOptionsCached(ByRef wksht As Worksheet) As Boolean
    If CollectionKeyExists(protSettingCache, KeyCachedProtection(wksht)) Then
        ProtectOptionsCached = True
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
        If ProtectOptionsCached(wksht) Then
           Dim cOptions() As Variant
           cOptions = CachedProtectionOptions(wksht)
           password = cOptions(1)
           options = cOptions(2)
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
    



' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   PRIVATE VS OPEN
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
#If privateVersion Then
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   *** PRIVATE *** IMPLEMENTATION OF COMMON FUNCTIONS
'   These Methods Are Used Only When 'privateVersion = 1' Exists
'       in VBA Project Conditional Compilation Arguments
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ProtectSht(ByRef ws As Worksheet, Optional ByVal forceProtect As Boolean = False) As Boolean
        ProtectSht = ProtectShtPriv(ws, forceProtect:=forceProtect)
    End Function
    Public Function UnprotectSht(ByRef ws As Worksheet) As Boolean
        UnprotectSht = UnprotectSHTPriv(ws)
    End Function
    Public Property Get byPassOnCloseCheck() As Boolean
        byPassOnCloseCheck = lBypassOnCloseCheck
        If IsUpgrader Then byPassOnCloseCheck = True
    End Property
    Public Property Let byPassOnCloseCheck(bypassCheck As Boolean)
        lBypassOnCloseCheck = bypassCheck
    End Property
    Public Property Get DevUserNames() As String
        DevUserNames = DEV_USERNAME
    End Property
    Public Sub ftBeep(bpType As BeepType)
        Dim doBeep    As Boolean
        If ftSYS.IsInitialized = False Then
            Exit Sub
        End If
        Select Case bpType
            Case BeepType.btMsgBoxOK
                doBeep = (Setting2(seBeepMsgBoxOK) = True)
            Case BeepType.btError, BeepType.btForced
                doBeep = True
            Case BeepType.btMsgBoxChoice
                doBeep = (Setting2(sebeepmsgboxchoice) = True)
            Case BeepType.btBusyWait
                doBeep = (Setting2(sebeepbusywait) = True)
            Case BeepType.btButton
                doBeep = (Setting2(seBeepButton) = True)
        End Select
        If doBeep Then
            Beep
        End If
    End Sub
    Public Property Get AppVersion() As Variant
        AppVersion = FinToolVersion
    End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   *** END PRIVATE *** IMPLEMENTATION OF COMMON
'       FUNCTIONS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
#Else
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   *** PUBLIC *** IMPLEMENTATION OF COMMON FUNCTIONS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'    Public Function ProtectSht(ByRef ws As Worksheet, Optional ByVal forceProtect As Boolean = False) As Boolean
'        ProtectSht = True
'        Debug.Print "pbCommon.ProtectSht - Not Implemented in Open Source Version"
'    End Function
'    Public Function UnprotectSht(ByRef ws As Worksheet) As Boolean
'        UnprotectSht = True
'        Debug.Print "pbCommon.UnprotectSht - Not Implemented in Open Source Version"
'    End Function
    Public Property Get byPassOnCloseCheck() As Boolean
        byPassOnCloseCheck = lBypassOnCloseCheck
    End Property
    Public Property Let byPassOnCloseCheck(bypassCheck As Boolean)
        lBypassOnCloseCheck = bypassCheck
    End Property
    Public Property Get DevUserNames() As String
        DevUserNames = "paulbrower|OTHERLOGINS"
    End Property
    Public Sub ftBeep(bpType As BeepType)
        Dim doBeep    As Boolean
        Select Case bpType
            Case BeepType.btMsgBoxOK
            Case BeepType.btError, BeepType.btForced
            Case BeepType.btMsgBoxChoice
            Case BeepType.btBusyWait
            Case BeepType.btButton
        End Select
        Beep
    End Sub
    Public Property Get AppVersion() As Variant
        AppVersion = CDbl(1)
        Err.Raise 1004, Source:="pbCommon.AppVersion", Description:="Not Implemented"
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   *** END PUBLIC *** IMPLEMENTATION OF COMMON
'       FUNCTIONS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
#End If

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


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   COPY ENTIRE WORKSHEET TO A DIFFERENT WORKBOOK
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function CopyWorksheetUtil(sourceWksht As Worksheet, targetWB As Workbook)
On Error GoTo E:
    Dim failed As Boolean
    Dim evts As Boolean: evts = Application.EnableEvents
    Dim srcVis As Variant
    srcVis = sourceWksht.Visible
    
    Application.EnableEvents = False
    sourceWksht.Visible = xlSheetVisible
    
    If Not targetWB Is Nothing Then
        With sourceWksht
            .Copy After:=targetWB.Worksheets(1)
            DoEvents
        End With
    End If

Finalize:
    On Error Resume Next
    sourceWksht.Visible = srcVis
    If Not failed Then
        Beep
    End If
    Application.EnableEvents = evts
    
    Exit Function
E:
    failed = True
    sourceWksht.Visible = srcVis
    Application.EnableEvents = evts
    Err.Raise Err.number
    Resume Finalize:

End Function

Public Function ReplaceIllegalCharacters(ByVal strIn As String, ByVal strChar As String, Optional ByVal padSingleQuote As Boolean = True, Optional useForSpecialChars As Variant) As String
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

Public Property Get IsDEV() As Boolean
    Dim retV As Boolean
    Dim devNames
    devNames = Split(DevUserNames, "|", , vbTextCompare)
    Dim i As Long, compareEnvName
    For i = LBound(devNames) To UBound(devNames)
        compareEnvName = devNames(i)
        If StringsMatch(ENV_LogName, compareEnvName, smContains) Then
            retV = True
            Exit For
        End If
    Next i
    IsDEV = retV
    If Not IsDeveloper = retV Then IsDeveloper = retV
End Property

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

Public Function ConcatWithDelim(ByVal delimeter As String, ParamArray Items() As Variant) As String
    ConcatWithDelim = Join(Items, delimeter)
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
Public Function Concat(ParamArray Items() As Variant) As String
    Concat = Join(Items, "")
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

Public Property Get ENV_HOME() As String
    ENV_HOME = VBA.Interaction.Environ("HOME")
End Property
Public Property Get ENV_TEMPDIR() As String
    ENV_TEMPDIR = VBA.Interaction.Environ("TMPDIR")
End Property

Public Property Get DBLQUOTE() As String
    DBLQUOTE = Chr(34)
End Property

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
        Err.Clear
    End If
End Function

Public Function FullWbNameCorrected(Optional wkbk As Workbook) As String
On Error Resume Next
    Dim Fname As String
    If wkbk Is Nothing Then
        Fname = ThisWorkbook.FullName
    Else
        Fname = wkbk.FullName
    End If
    If Len(Fname) > 0 Then
        If InStr(1, Fname, "http", vbTextCompare) > 0 Then
            Fname = Replace(Fname, " ", "%20", Compare:=vbTextCompare)
        End If
    End If
    FullWbNameCorrected = Fname
    If Err.number <> 0 Then Err.Clear
End Function

Public Function SimpleURLEncode(ByVal fPath As String) As String
    If Len(fPath) > 0 Then
        If InStr(1, fPath, "http", vbTextCompare) > 0 Then
            fPath = Replace(fPath, " ", "%20", Compare:=vbTextCompare)
        End If
    End If
    SimpleURLEncode = fPath
    If Err.number <> 0 Then Err.Clear
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

    Function ReplaceIllegalCharacters2(strIn As String, strChar As String, Optional padSingleQuote As Boolean = True) As String
        Dim strSpecialChars As String
        Dim i As Long
        strSpecialChars = "~""#%&*:<>?{|}/\[]" & Chr(10) & Chr(13)
    
        For i = 1 To Len(strSpecialChars)
            strIn = Replace(strIn, Mid$(strSpecialChars, i, 1), strChar)
        Next
        
        If padSingleQuote And InStr(1, strIn, "''") = 0 Then
            strIn = CleanSingleTicks(strIn)
        End If
        
        ReplaceIllegalCharacters2 = strIn
    End Function

    ' ~~~~~~~~~~   CLEAN SINGLE TICKS ~~~~~~~~~~'
    Public Function CleanSingleTicks(ByVal wbName As String) As String
        Dim retV As String
        If InStr(wbName, "'") > 0 And InStr(wbName, "''") = 0 Then
            retV = Replace(wbName, "'", "''")
        Else
            retV = wbName
        End If
        CleanSingleTicks = retV
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
    Dim exists As Boolean
    exists = EnumCompare(theEnum, enumMember)
    If exists And modifyType = feVerifyEnumRemoved Then
        theEnum = theEnum - enumMember
    ElseIf exists = False And modifyType = feVerifyEnumExists Then
        theEnum = theEnum + enumMember
    End If
    EnumModify = theEnum
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   INPUT BOX
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function InputBox_FT(prompt As String, Optional title As String = "Financial Tool - Input Needed", Optional default As Variant, Optional inputType As ftInputBoxType, Optional useVBAInput As Boolean = False) As Variant
    ftBeep btMsgBoxChoice
    
    If useVBAInput Then
        InputBox_FT = VBA.InputBox(prompt, title:=title, default:=default)
    Else
        If inputType >= 0 Then
            InputBox_FT = Application.InputBox(prompt, title:=title, default:=default, Type:=inputType)
        Else
            InputBox_FT = Application.InputBox(prompt, title:=title, default:=default)
        End If
    End If
    
    DoEvents
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
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   ASK YES/NO
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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
    If Not DtPart(dtWeekday, processDt, firstDayOfWeek:=vbMonday) = 7 Then
        processDt = DtAdd(dtDay, 7 - DtPart(dtWeekday, processDt, firstDayOfWeek:=vbMonday), processDt)
    End If
    GetSundayOfWeek = processDt
End Function

Public Function GetMondayOfWeek(inputDate As Variant) As Date
    Dim processDt As Variant
    If TypeName(inputDate) = "String" Then
        processDt = DateValue(inputDate)
    Else
        processDt = inputDate
    End If
    Dim diffDays As Long
    diffDays = 1 - DtPart(dtWeekday, processDt, firstDayOfWeek:=vbMonday)
    If diffDays <> 0 Then
        processDt = DtAdd(dtDay, diffDays, processDt)
    End If
    GetMondayOfWeek = processDt
End Function
Public Function DateAddDays(dt As Variant, addDays As Double) As Date
    'KEEP FOR BACKWARDS COMPATIBILITY
    DateAddDays = DtAdd(dtDay, addDays, dt)
End Function

Public Function WaitWithDoEvents(waitSeconds As Long)
'WAIT FOR N SECONDS WHILE ALLOWING OTHER EXCEL EVENT TO PROCESS
'PURPOSE IS TO ENABLE ENOUGHT TIME TO PASS FOR APPLICATION ONTIME TO TAKE HOLD
    Dim stTimer As Single
    stTimer = Timer
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
        If StringsMatch(argVal1, ".xlam", strMatchEnum.smContains) Or StringsMatch(argVal1, ".xlsm", strMatchEnum.smContains) Then
            argVal1 = argVal1
        End If
    End If
    If TypeName(argVal2) = "String" Then
        If StringsMatch(argVal2, ".xlam", strMatchEnum.smContains) Or StringsMatch(argVal2, ".xlsm", strMatchEnum.smContains) Then
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
Public Property Get StartupPath() As String
    StartupPath = PathCombine(True, Application.StartupPath)
End Property
Public Function URLEncode(ByRef txt As String) As String
    Dim buffer As String, i As Long, c As Long, N As Long
    buffer = String$(Len(txt) * 12, "%")
 
    For i = 1 To Len(txt)
        c = AscW(Mid$(txt, i, 1)) And 65535
 
        Select Case c
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95  ' Unescaped 0-9A-Za-z-._ '
                N = N + 1
                Mid$(buffer, N) = ChrW(c)
            Case Is <= 127            ' Escaped UTF-8 1 bytes U+0000 to U+007F '
                N = N + 3
                Mid$(buffer, N - 1) = Right$(Hex$(256 + c), 2)
            Case Is <= 2047           ' Escaped UTF-8 2 bytes U+0080 to U+07FF '
                N = N + 6
                Mid$(buffer, N - 4) = Hex$(192 + (c \ 64))
                Mid$(buffer, N - 1) = Hex$(128 + (c Mod 64))
            Case 55296 To 57343       ' Escaped UTF-8 4 bytes U+010000 to U+10FFFF '
                i = i + 1
                c = 65536 + (c Mod 1024) * 1024 + (AscW(Mid$(txt, i, 1)) And 1023)
                N = N + 12
                Mid$(buffer, N - 10) = Hex$(240 + (c \ 262144))
                Mid$(buffer, N - 7) = Hex$(128 + ((c \ 4096) Mod 64))
                Mid$(buffer, N - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, N - 1) = Hex$(128 + (c Mod 64))
            Case Else                 ' Escaped UTF-8 3 bytes U+0800 to U+FFFF '
                N = N + 9
                Mid$(buffer, N - 7) = Hex$(224 + (c \ 4096))
                Mid$(buffer, N - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, N - 1) = Hex$(128 + (c Mod 64))
        End Select
    Next
    URLEncode = Left$(buffer, N)
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
            retV = Replace(retV, Mid(retV, i, 1), "", Compare:=vbBinaryCompare)
        End If
    Next i
    SanitizeAlpha = Trim(retV)
End Function




Public Property Get IsDeveloper() As Boolean
    IsDeveloper = l_IsDeveloper
End Property
Public Property Let IsDeveloper(devMode As Boolean)
    l_IsDeveloper = devMode
End Property


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'
'   FILE SYSTEM FUNCTIONS
'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function CopySheetToNewWB(ByVal ws As Worksheet, Optional filePath As Variant, Optional fileName As Variant)
On Error Resume Next
    Application.EnableEvents = False
    Dim newWB As Workbook
    Set newWB = Application.Workbooks.Add
    With ws
        .Copy Before:=newWB.Worksheets(1)
        DoEvents
    End With
    If IsMissing(filePath) Then filePath = Application.DefaultFilePath
    If IsMissing(fileName) Then fileName = ReplaceIllegalCharacters2(ws.Name, vbEmpty) & ".xlsx"
    newWB.SaveAs fileName:=PathCombine(False, filePath, fileName), FileFormat:=xlOpenXMLStrictWorkbook
    Application.EnableEvents = True
    If Not Err.number = 0 Then
        MsgBox "CopySheetToNewWB Error: " & Err.number & ", " & Err.Description
        Err.Clear
    End If

End Function




    Public Function SaveCopyToUserDocFolder(ByVal wb As Workbook, Optional fileName As Variant)
        SaveWBCopy wb, Application.DefaultFilePath, IIf(IsMissing(fileName), wb.Name, CStr(fileName))
    End Function

Public Function SaveWBCopy(ByVal wb As Workbook, dirPath As String, fileName As String)
On Error Resume Next
    Application.EnableEvents = False
    wb.SaveCopyAs PathCombine(False, dirPath, fileName)
    Application.EnableEvents = True
    If Not Err.number = 0 Then
        MsgBox "SaveWBCopy Error: " & Err.number & ", " & Err.Description
        Err.Clear
    End If
End Function

Public Function openPath(fldrPath As String)
'   Open Folder (MAC and PC Supported)
On Error Resume Next
    ftBeep btMsgBoxChoice
    Dim retV As Variant

    #If Mac Then
        Dim scriptStr As String
        scriptStr = "do shell script " & Chr(34) & "open " & fldrPath & Chr(34)
        MacScript (scriptStr)
    #Else
        Call Shell("explorer.exe " & fldrPath, vbNormalFocus)
    #End If
    
    If Err.number <> 0 Then
        Debug.Print "pbCommon.OpenFolder - Error Opening: (" & fldrPath & ") "
        Err.Clear
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Create Valid File or Directory Path (for PC or Mac, local,
'   or internet) from 1 or more arguments
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function PathCombine(includeEndSeparator As Boolean, ParamArray vals() As Variant) As String
    
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

Public Function FullPathExcludingFileName(fullFileName As String) As String
On Error Resume Next
    Dim tPath As String, tfileName As String, fNameStarts As Long
    tfileName = FileNameFromFullPath(fullFileName)
    fNameStarts = InStr(fullFileName, tfileName)
    tPath = Mid(fullFileName, 1, fNameStarts - 1)
    FullPathExcludingFileName = tPath
    If Err.number <> 0 Then Err.Clear
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
    If Err.number <> 0 Then Err.Clear
End Function
Public Function ChooseFolder(choosePrompt As String) As String
'   Get User-Selected Directory name (MAC and PC Supported)
On Error Resume Next
    ftBeep btMsgBoxChoice
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
    If Err.number <> 0 Then Err.Clear
End Function

Public Function RequestFileAccess(ParamArray files() As Variant)

    #If Mac Then
        'Declare Variables?
        Dim fileAccessGranted As Boolean
        Dim filePermissionCandidates
    
        'Create an array with file paths for the permissions that are needed.?
    '    filePermissionCandidates = Array("/Users//Desktop/test1.txt", "/Users//Desktop/test2.txt")
        filePermissionCandidates = files
    
        'Request access from user.?
        fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)
        'Returns true if access is granted; otherwise, false.
    #End If
End Function

Public Function FileNameWithoutExtension(ByVal fileName As String) As String
    If InStrRev(fileName, ".") > 0 Then
        Dim tmpExt As String
        tmpExt = Mid(fileName, InStrRev(fileName, "."))
        If Len(tmpExt) >= 2 Then
            fileName = Replace(fileName, tmpExt, vbNullString)
        End If
    End If
    FileNameWithoutExtension = fileName
End Function

Public Function SaveFileAs(savePrompt As String, Optional ByVal defaultFileName, Optional ByVal fileExt) As String
On Error Resume Next
    ftBeep btMsgBoxChoice
    Dim retV As Variant
        
    #If Mac Then
        If Len(fileExt) > 0 Then
            
            fileExt = Replace(Replace(fileExt, "*", ""), ".", "")
            retV = Application.GetSaveAsFilename(InitialFileName:=IIf(IsMissing(defaultFileName), "", defaultFileName), FileFilter:=IIf(IsMissing(fileExt), "", fileExt), ButtonText:="USE THIS NAME")
        Else
            retV = Application.GetSaveAsFilename(InitialFileName:=IIf(IsMissing(defaultFileName), "", defaultFileName), FileFilter:=IIf(IsMissing(fileExt), "", fileExt), ButtonText:="USE THIS NAME")
        End If
    #Else
NextCode:
        If Len(fileExt) > 0 Then
            fileExt = Replace(Replace(fileExt, "*", ""), ".", "")
            fileExt = Concat("*.", fileExt, "*")
            fileExt = "Files (" & fileExt & "), " & fileExt
            If Len(defaultFileName) > 0 Then
                retV = Application.GetSaveAsFilename(InitialFileName:=defaultFileName, FileFilter:=fileExt, title:=savePrompt, ButtonText:="USE THIS NAME")
            Else
                retV = Application.GetSaveAsFilename(FileFilter:=fileExt, title:=savePrompt, ButtonText:="USE THIS NAME")
            End If
            'retV = Application.GetOpenFilename(InitialFileName:=IIf(IsMissing(defaultFileName), "", defaultFileName), FileFilter:=fileExt, title:=choosePrompt, ButtonText:="USE THIS NAME")
        End If
    
    #End If
    
    If Err.number = 0 Then
        SaveFileAs = CStr(retV)
    Else
        Debug.Print "ERROR: pbCommon.ChooseFile "
        Err.Clear
    End If

End Function

Public Function chooseFile(choosePrompt As String, Optional ByVal fileExt As String = vbNullString) As String
'TODO:  Also check out Application.GetSaveAsFileName
'   Get User-Select File Name (MAC and PC Supported)
On Error Resume Next
    ftBeep btMsgBoxChoice
    Dim retV As Variant
        
    #If Mac Then
        If Len(fileExt) > 0 Then
            fileExt = Replace(Replace(fileExt, "*", ""), ".", "")
            retV = Application.GetOpenFilename(FileFilter:=fileExt, ButtonText:="CHOOSE FILE")
        Else
            retV = Application.GetOpenFilename(title:=choosePrompt)
        End If
    #Else
NextCode:
        If Len(fileExt) > 0 Then
            fileExt = Replace(Replace(fileExt, "*", ""), ".", "")
            fileExt = Concat("*.", fileExt, "*")
            fileExt = "Files (" & fileExt & "), " & fileExt
            retV = Application.GetOpenFilename(FileFilter:=fileExt, title:=choosePrompt, ButtonText:="CHOOSE FILE")
        Else
            retV = Application.GetOpenFilename(title:=choosePrompt, ButtonText:="CHOOSE FILE")
        End If
    
    #End If
    
    If Err.number = 0 Then
        chooseFile = CStr(retV)
    Else
        Debug.Print "ERROR pbCommon.ChooseFile "
        Err.Clear
    End If

End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   CREATE THE ** LAST ** DIRECTORY IN 'fullPath'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function CreateDirectory(fullPath As String) As Boolean
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
                If Err.number = 0 Then
                    retV = DirectoryExists(fullPath)
                End If
            End If
        End If
    End If
    CreateDirectory = retV
    If Err.number <> 0 Then Err.Clear
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Returns true if filePth Exists and is not a directory
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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
        If Err.number <> 0 Then Debug.Print "DirectoryExists: Err Getting Path: " & filePth & ", " & Err.number & " - " & Err.Description
    End If
    FileExists = retV
    If Err.number <> 0 Then Err.Clear
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Returns true if DIRECTORY path dirPath) Exists
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function DirectoryExists(dirPath As String) As Boolean
On Error Resume Next
    Dim retV As Boolean
    Dim lastDirName As String, pathArr As Variant, checkFldrName As String
    dirPath = PathCombine(False, dirPath)

    If InStr(dirPath, Application.PathSeparator) > 0 Then
        pathArr = Split(dirPath, Application.PathSeparator)
        checkFldrName = CStr(pathArr(UBound(pathArr)))
        retV = StrComp(Dir(dirPath & "*", vbDirectory), LCase(checkFldrName), vbTextCompare) = 0
        If Err.number <> 0 Then
            Debug.Print "DirectoryExists: Err Getting Path: " & dirPath & ", " & Err.number & " - " & Err.Description
        End If
    End If
    DirectoryExists = retV
    If Err.number <> 0 Then Err.Clear
End Function

Public Function DeleteFolderFiles(folderPath As String, Optional patternMatch As String = vbNullString)
On Error Resume Next
    folderPath = PathCombine(True, folderPath)
    
    If DirectoryFileCount(folderPath) > 0 Then
        Dim MyPath As Variant
        MyPath = PathCombine(True, folderPath)
        ChDir folderPath
        Dim myFile, MyName As String
        MyName = Dir(MyPath, vbNormal)
        Do While MyName <> ""
            If (GetAttr(PathCombine(False, MyPath, MyName)) And vbNormal) = vbNormal Then
                If patternMatch = vbNullString Then
                    Kill PathCombine(False, MyPath, MyName)
                Else
                    If LCase(MyName) Like LCase(patternMatch) Then
                        Kill PathCombine(False, MyPath, MyName)
                    End If
                End If
            End If
            MyName = Dir()
        Loop
    End If
    If Err.number <> 0 Then Err.Clear
End Function



Public Function DirectoryFileCount(tmpDirPath As String) As Long
On Error Resume Next

    Dim myFile, MyPath, MyName As String, retV As Long
    MyPath = PathCombine(True, tmpDirPath)
    MyName = Dir(MyPath, vbNormal)
    Do While MyName <> ""
        If (GetAttr(PathCombine(False, MyPath, MyName)) And vbNormal) = vbNormal Then
            retV = retV + 1
        End If
        MyName = Dir()
    Loop
    DirectoryFileCount = retV
    If Err.number <> 0 Then Err.Clear
End Function

Public Function DirectoryDirectoryCount(tmpDirPath As String) As Long
On Error Resume Next

    Dim myFile, MyPath, MyName As String, retV As Long
    MyPath = PathCombine(True, tmpDirPath)
    MyName = Dir(MyPath, vbDirectory)
    Do While MyName <> ""
        If (GetAttr(PathCombine(False, MyPath, MyName)) And vbDirectory) = vbDirectory Then
            retV = retV + 1
        End If
        MyName = Dir()
    Loop
    DirectoryDirectoryCount = retV
    If Err.number <> 0 Then Err.Clear
End Function

Public Function DeleteFile(filePath As String) As Boolean
    On Error Resume Next
    If FileExists(filePath) Then
        Kill filePath
        DoEvents
    End If
    DeleteFile = FileExists(filePath) = False
End Function



Public Function GetFiles(dirPath As String) As Variant()
On Error Resume Next
    
    Dim cl As New Collection

    Dim myFile, MyPath, MyName As String
    MyPath = PathCombine(True, dirPath)
    MyName = Dir(MyPath, vbNormal)
    Do While MyName <> ""
        If (GetAttr(PathCombine(False, MyPath, MyName)) And vbNormal) = vbNormal Then
            cl.Add MyName
        End If
        MyName = Dir()
    Loop
    
    If cl.Count > 0 Then
        Dim retV() As Variant
        ReDim retV(1 To cl.Count, 1 To 1)
        Dim fidx As Long
        For fidx = 1 To cl.Count
            retV(fidx, 1) = cl(fidx)
        Next fidx
        GetFiles = retV
    End If
    
    If Err.number <> 0 Then Err.Clear
    
End Function

Public Function DirectoryFileCount2(tmpDirPath As String) As Long
On Error Resume Next

    Dim myFile, MyPath, MyName As String, retV As Long
    MyPath = PathCombine(True, tmpDirPath)
    MyName = Dir(MyPath, vbNormal)
    Do While MyName <> ""
        If (GetAttr(PathCombine(False, MyPath, MyName)) And vbNormal) = vbNormal Then
            retV = retV + 1
        End If
        MyName = Dir()
    Loop
    DirectoryFileCount2 = retV
    If Err.number <> 0 Then Err.Clear
End Function

Public Function DirectoryDirectoryCount2(tmpDirPath As String) As Long
On Error Resume Next

    Dim myFile, MyPath, MyName As String, retV As Long
    MyPath = PathCombine(True, tmpDirPath)
    MyName = Dir(MyPath, vbDirectory)
    Do While MyName <> ""
        If (GetAttr(PathCombine(False, MyPath, MyName)) And vbDirectory) = vbDirectory Then
            retV = retV + 1
        End If
        MyName = Dir()
    Loop
    DirectoryDirectoryCount2 = retV
    If Err.number <> 0 Then Err.Clear
End Function


Public Function TempDirName2(Optional dirName As String = vbNullString) As String
    TempDirName2 = IIf(Not dirName = vbNullString, dirName, TEMP_DIRECTORY_NAME2)
End Function

Public Property Get VisibleWorksheets() As Long
    Dim tIdx As Long, retV As Long
    For tIdx = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Worksheets(tIdx).Visible = xlSheetVisible Then
            retV = retV + 1
        End If
    Next tIdx
    VisibleWorksheets = retV
End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   MAKE SURE ALL 'NAMES' ARE VISIBLE IN NAMES MANAGER
'    - ALL WORKSHEETS MUST BE UNPROTECTED
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function DEVMakeAllNamesVisible()
    'This makes all the names visible, then can manually delete from
    '   Formulas --> Name Manager
    Dim nm As Name
    For Each nm In ThisWorkbook.Names
        nm.Visible = True
    Next nm
    MsgBox "Check Formulas --> Name Manager to view names"
End Function


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

Public Function NowWithMS() As String
    NowWithMS = Format(Now, "yyyymmdd hh:mm:ss") & Right(Format(Timer, "0.000"), 4)
End Function

Public Function SysStates() As String
    Dim tEv As String, tSc As String, tIn As String, tCa As String, retV As String
    tEv = IIf(Events, "Evts=ON  ", "")
    tSc = IIf(Application.ScreenUpdating, "Scrn=ON  ", "")
    tIn = IIf(Application.Interactive, "Inter=ON  ", "")
    tCa = IIf(Application.Calculation = xlCalculationAutomatic, "Calc: AUTO  ", "")
    retV = Concat(tEv, tSc, tIn, tCa)
    If Len(retV) = 0 Then
        retV = "SysStates: (ALL OFF)"
    Else
        retV = Concat("SysStates: ( ", retV, ")")
    End If
    SysStates = retV
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Set Zoom Level
'   [zoomLevel] = Zoom Level.  Default is: 100.
'   If [wksht] arguement is missing, will be applied to all
'   worksheets in [wkbk]
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function SetZoom(wkbk As Workbook, Optional zoomLevel, Optional wksht As Worksheet)
    On Error GoTo E:
    Dim evts As Boolean
    evts = Events
    EventsOff
    If IsMissing(zoomLevel) Then
        zoomLevel = 100
    ElseIf Not IsNumeric(zoomLevel) Then
        zoomLevel = 100
    End If
    
    Dim visStatus As Variant, tmpWS As Worksheet, doZoom As Boolean, curWS As Worksheet
    Set curWS = wkbk.ActiveSheet
    Dim shtView As WorksheetView
    For Each shtView In wkbk.Windows(1).SheetViews
        doZoom = False
        If Not wksht Is Nothing Then
            If shtView.Sheet Is wksht Then
                doZoom = True
            End If
        Else
            doZoom = True
        End If
'        If shtView.Sheet Is wsBusy Then
'            doZoom = False
'        End If
        If doZoom Then
            visStatus = shtView.Sheet.Visible
            shtView.Sheet.Visible = xlSheetVisible
            shtView.Sheet.Activate
            shtView.Sheet.Parent.Windows(1).Zoom = zoomLevel
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False
            shtView.Sheet.Visible = visStatus
        End If
    Next shtView
    
Finalize:
    On Error Resume Next
        curWS.Activate
        Events = evts
    If Err.number <> 0 Then Err.Clear
    Exit Function
E:
    ErrorCheck
    Resume Finalize:
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
    If Not ActiveWindow.ActiveSheet Is wksht Then Exit Function
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
            ElseIf .SplitRow > 0 Then
                .Panes(pnIdx).ScrollRow = .SplitRow + 1
            End If
            If pnIdx = 1 Or pnIdx = 2 Then
                .Panes(pnIdx).ScrollColumn = 1
            ElseIf .SplitColumn > 0 Then
                .Panes(pnIdx).ScrollColumn = .SplitColumn + 1
            End If
        Next pnIdx
        'Ensure split panes are in the right place
        If splitOnRow > 0 And Not .SplitRow = splitOnRow Then
            .SplitRow = splitOnRow
        ElseIf splitOnRow = 0 And .SplitRow <> 0 And removeUnspecified Then
            .SplitRow = 0
        End If
        If splitOnColumn > 0 And Not .SplitColumn = splitOnColumn Then
            .SplitColumn = splitOnColumn
        ElseIf splitOnColumn = 0 And .SplitColumn <> 0 And removeUnspecified Then
            .SplitColumn = 0
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

Public Function LastPopulatedRow(wks As Worksheet, Optional column As Variant) As Long
    Dim lpr As Long
    Dim rowOffset As Long, colOffset As Long, urColEnd As Long, urRowEnd As Long
    rowOffset = wks.UsedRange.Row - 1
    colOffset = wks.UsedRange.column - 1
    urColEnd = wks.UsedRange.Columns.Count + colOffset
    urRowEnd = wks.UsedRange.Rows.Count + rowOffset
    Dim exactCol As Long: If Not IsMissing(column) Then exactCol = column
    If exactCol > 0 And exactCol < urColEnd Then urColEnd = exactCol
    '   If asking for Column outside last col with data, then return 0
    If exactCol > urColEnd Then
        lpr = 0
    '   HANDLE EMPTY SHEET ( OR JUST $A$1 HAS DATA)
    ElseIf urRowEnd = 1 And urColEnd = 1 Then
        lpr = IIf(Len(wks.Range("A1").Text) > 0, 1, 0)
        If Not exactCol > 0 And lpr = 1 Then
            If exactCol > 1 Then lpr = 0
        End If
    '   HANDLE SINGLE CELL POPULATED OTHER THAN $A$1
    ElseIf (VarType(wks.UsedRange.Text) And VbVarType.vbArray) = 0 Then
        lpr = urRowEnd
        If exactCol > 0 Then
            ' ONLY ONE CELL POPULATED IIF urColEnd doesn't match Column, then return 0
            If exactCol <> urColEnd Then lpr = 0
        End If
    Else
        lpr = urRowEnd
    End If
   'SHOULD BE GOOD, UNLESS THE ROW [LPR] ISN'T VISIBLE
   '(HIDDEN ROW THAT HAS NOT DATA IS STILL COUNTED IN USED RANGE, SO
   ' NOW WE NEED TO CHECK THAT)
   If lpr > 0 Then
        Dim deepCheck As Boolean
        If wks.Rows(lpr).Hidden Then deepCheck = True
        If Not deepCheck And exactCol > 0 Then
            If Len(wks.Cells(lpr, exactCol).Text) = 0 Then deepCheck = True
        End If
        If deepCheck Then
            Dim rowIdx As Long, colIdx As Long
            Dim hasRowData As Boolean
            For rowIdx = lpr To 1 Step -1
                If exactCol > 0 Then
                    hasRowData = Len(wks.Cells(rowIdx, exactCol + colOffset)) > 0
                    If hasRowData Then
                        lpr = rowIdx
                    ElseIf rowIdx = 1 Then
                        'NO ROWS IN [COLUMN] HAVE ANY DATA
                        lpr = 0
                        Exit For
                    End If
                Else
                    For colIdx = 1 To urColEnd
                        hasRowData = Len(wks.Cells(rowIdx, colIdx + colOffset)) > 0
                        If hasRowData Then
                            lpr = rowIdx
                            Exit For
                        End If
                    Next colIdx
                End If
                DoEvents
                If hasRowData Then Exit For
            Next rowIdx
        End If
    End If
    LastPopulatedRow = lpr
End Function



' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'
'       ERROR HANDLING
'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Get ErrString(Optional customSrc As String, Optional errNUM As Variant, Optional errDESC As Variant, Optional errERL As Variant) As String
'   Format Known Error Information
    If IsMissing(errNUM) Then errNUM = Err.number
    If IsMissing(errDESC) Then errDESC = Err.Description
    If IsMissing(errERL) Then errERL = Erl
    If errNUM <> 0 Then
        Dim msg As String
        msg = "ERROR: " & errNUM & ", Desc: " & errDESC & ", Src: " & Err.Source
        If errERL <> 0 Then msg = msg & " (ERL: " & errERL & ") "
        If Len(customSrc & "") > 0 Then
            msg = Concat(customSrc, vbCrLf, msg)
        End If
        ErrString = msg
    End If
End Property
'Called When Must End
Private Function FatalEnd()
'    RaiseEvent
    Err.Raise 1004, Description:="FatalEnd Not Implemented"
End Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   RAISE ERROR
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Sub RaiseError(errNumber As Long, Optional errorDesc As String = vbNullString, Optional resetUserInterface As Boolean = True)
    'LogError Concat("RaiseError Called: ", errNumber, " - errorDesc: ", errorDesc)
    
    Dim cancelRaise As Variant
    pbCommonEvents.RaiseOnRaiseError errNumber, errorDesc, Application.Caller, cancelRaise
    If cancelRaise = True Then
        Exit Sub
    Else
        If resetUserInterface Then
            Application.ScreenUpdating = True
            Application.Cursor = xlDefault
            Application.EnableEvents = True
        End If
        If Len(errorDesc) > 0 Then
            Err.Raise errNumber, Description:=errorDesc
        Else
            Err.Raise errNumber, Description:=(Error(errNumber))
        End If
    End If

End Sub

Public Function ErrorCheck(Optional Source As String, Optional options As ErrorOptionsEnum, Optional customErrorMsg As String) As Long
    Dim errNumber As Long, errDESC As Variant, errorInfo As String, ignoreError As Boolean, errERL As Long
    errNumber = Err.number
    errDESC = Err.Description
    errERL = Erl
    errorInfo = ErrString(customSrc:=Source, errNUM:=errNumber, errDESC:=errDESC, errERL:=errERL)
    If errNumber = 0 Then Exit Function
    pbCommonEvents.RaiseOnErrorHandlerBegin errNumber, errDESC, Source
    
    pbCommonEvents.RaiseOnErrorHandlerEnd errNumber, errDESC, Source
    
'    LogError Concat("pbError.ErrorCheck: ", errorInfo)
    
''    If ThisWorkbook.ReadOnly Then
''        LogError "Workbook is READ-ONLY - Closing NOW"
''        pbPerf.RestoreDefaultAppSettingsOnly
''        ThisWorkbook.Close SaveChanges:=False
''        Exit Function
''    End If
''
''    ftBeep btError
''
''    If options = 0 Then options = ftDefaults
''    On Error GoTo -1
''    On Error Resume Next
''
''
''    Dim cancelMsg As String
''    Dim okONLY As Boolean
''    okONLY = True
''    If (IsDEV And DebugMode) Or EnumCompare(options, ftERR_ResponseAllowBreak) Then
''        okONLY = False
''    End If
''    If okONLY Then
''        cancelMsg = vbCrLf & "** PRESS OK TO CONTINUE"
''    Else
''        cancelMsg = vbCrLf & "** PRESS OK TO CONTINUE CODE EXECUTION -- WHICH IS NORMALLY WHAT YOU WANT TO DO.  'CANCEL'  WILL STOP CODE FROM RUNNING, AND SHOULD ONLY BE USED IF YOU'RE CONTINUALLY SEEING ERROR MESSAGES"
''    End If
''
''    Dim EventsOn As Boolean: EventsOn = Application.EnableEvents
''    Application.EnableEvents = False
''
''    '***** ftErr_NoBeeper
''    If ignoreError = False And EnumCompare(options, ftERR_NoBeeper) = False Then Beeper
''
''    If (customErrorMsg & "") <> vbNullString Then
''        errorInfo = errorInfo & vbCrLf & customErrorMsg
''    End If
''
''    '*****ftERR_MessageIgnore
''    If ignoreError = False And EnumCompare(options, ftERR_MessageIgnore) = False Then
''        If okONLY = False Then
''            If MsgBox_FT(CStr(errorInfo) & cancelMsg, vbOKCancel + vbCritical + vbSystemModal + vbDefaultButton1, "ERROR LOGGED TO LOG FILE") = vbCancel Then
''                If Err.number <> 0 Then Err.Clear
''                FatalEnd
''                Exit Function
''            End If
''        Else
''            If IsDEV Then
''                If MsgBox_FT(CStr(errorInfo) & cancelMsg, vbOKCancel + vbSystemModal, "DEV MSG: CANCEL TO STOP CODE") = vbCancel Then
''                    Beep
''                    DoEvents
''                    Stop
''                End If
''            Else
''                MsgBox_FT CStr(errorInfo) & cancelMsg, vbOKOnly + vbError, "ERROR LOGGED TO LOG FILE"
''            End If
''            If Err.number <> 0 Then Err.Clear
''            Exit Function
''        End If
''    End If
''
''    If errNumber <> 0 Then Err.Clear
''    If errNumber = 18 Then
''        ResetUI
''        MsgBox_FT "User Cancelled Current Process", title:="Cancelled"
''        End
''    End If
''
''    '*****ftERR_ProtectSheet
''    If ignoreError = False And EnumCompare(options, ftERR_ProtectSheet) Then
''        ProtectSht ThisWorkbook.ActiveSheet
''    End If
''    If Err.number <> 0 Then
''        ErrorCheck
''    End If
''    If ignoreError = False Then
''        ErrorCheck = errNumber
''    Else
''        ErrorCheck = 0
''    End If
''    If Err.number <> 0 Then Err.Clear
End Function
