Attribute VB_Name = "pbCommonUtil"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  Common Methods & Utililities
'  Most Modules/Classes in just-VBA Library are dependent
'  on this common module
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbCommonUtil.bas
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   PUBLIC GLOBAL ENUMS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

Public Enum AppModeEnum
    appStatusUnknown = 0
    appStatusStarting = 1
    appStatusRunning = 2
    appStatusClosing = 3
End Enum

Public Enum SheetSettingEnum
    ssAddNavigation = 1
    sssplitrow = 2
    sssplitcol = 3
    ssDoNotProtect = 4
End Enum

Public Enum LIstColType
    lcEditable
    lcFormula
    lcError
    lcDoubleClick
End Enum
Public Enum ListColAlign
    lcLeft
    lcRight
    lcCenter
End Enum
Public Enum ListColFormat
    lcGeneral
    lcTextForced
    lcLong
    lcLongMaskZero
    lcDecimalOne
    lcDecimalTwo
    lcDecimalThree
    lcDecimalFour
    lcCurrencyWhole
    lcCurrencyDecimal
    lcDateMMDDYYYY
    lcDateDDMMM
    lcDateDDMMMYY
    lcDateMMMYYYY
    lcDateDDMMMYYYY
    lcPercentWhole
    lcPercentDecimalTwo
End Enum


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   used to cache protection settings for specific worksheet
'   Key=WorkbookName.WorksheetCodeName
'   Value = 1-based Array:
'   (1)=password,(2)=SheetProtection enum
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private protSettingCache As New Collection
Private protSessionSheets As New Collection

Private l_pbPackageRunning As Boolean
Private l_appMode As AppModeEnum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   LISTOBJECT CONSTANTS (COMMON ONLY)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Const tblDevNotes As String = "tblDevNotes"

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   LOCAL SETTINGS KEY CONSTANTS (COMMON ONLY)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Const STG_DEV_PAUSELOCKING As String = "DEV_PAUSELOCKING"
Public Const STG_SYSTEM_PREFIX As String = "SYS_"
Public Const STG_USER_PREFIX As String = "USR_"

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   GENERALIZED CONSTANTS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Const CACHED_PROTECT_PWD_POSITION As Long = 1
Public Const CACHED_PROTECT_ENUM_POSITION As Long = 2
Public Const CFG_PROTECT_PASSWORD As String = "00000"
Public Const CFG_PROTECT_PASSWORD_EXPORT As String = "000001"
Public Const CFG_PROTECT_PASSWORD_MISC As String = "0000015"
Public Const CFG_PROTECT_PASSWORD_VBA = "0123210"
Public Const TEMP_DIRECTORY_NAME2 As String = "VBATemp"

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   CUSTOM ERROR CONSTANTS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   *** GENERIC ERRORS ONLY
''   *** STARTS AT 1001 - (513-1000 RESERVED FOR APPLICATION SPECIFIC ERRORS)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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
Public Const ERR_PBCOMMON_LOG_DISABLED = vbObjectError + 1028



' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   COMMON SETTINGS CONSTANTS
''   Setting Keys Ending with '_OS' Automatically Get
''       'PC' or 'MAC' Appended to Setting Key
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

Public Const STG_DO_NOT_HIDE As String = "DO_NOT_HIDE"
Public Const STG_DO_NOT_PROTECT As String = "DO_NOT_PROTECT"
Public Const STG_SPLIT_ROW_DEFAULT As String = "SPLIT_ROW_DEFAULT"
Public Const STG_SPLIT_COL_DEFAULT As String = "SPLIT_COL_DEFAULT"
Public Const STG_DEV_OUTPUT_TRACE As String = "DEV_OUTPUT_TRACE"
Public Const STG_LOGGING_ENABLED As String = "LOGGING_ENABLED"
Public Const STG_LOG_LEVEL As String = "LOG_LEVEL"
Public Const STG_DEV_NAMES = "DEV_NAMES"
Public Const STG_DEV_NO_NETWORK = "DEV_NO_NETWORK"
Public Const STG_DEV_SHOW_LOG_MESSAGES As String = "DEV_LOG_SHOW"
Public Const STG_ENABLE_SPLIT_PANES = "SPLIT_PANES_ENABLED"

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   GENERALIZED TYPES
'' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Type KVP
  key As String
  value As Variant
End Type
Public Type StringSequenceResult
    failed As Boolean
    searchString As String
    failedAtIndex As Long
    ''  Results
    ''  Each results first dimension contains searchedValue, foundAtIndex
    ''  e.g. If searched string was "AABBCC" and search sequence criteria was "AA", "C"
    ''  results() array would contain
    ''  results(1,1) = "AA", results(1,2) = 1
    ''  results(2,1) = "C", results(2,2) = 5
    Results() As Variant
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
    rows As Long
    Columns As Long
    Dimensions As Long
    Ubound_first As Long
    LBound_first As Long
    UBound_second As Long
    LBound_second As Long
    isArray As Boolean
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
    rows As Long
    Columns As Long
    AreasSameRows As Boolean
    AreasSameColumns As Boolean
    Areas As Long
End Type

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   GENERALIZED ENUMS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

Public Enum ModifyCollection
    mcAddUniqueKey
    mcRemoveUniqueKey
End Enum


'Public Enum SettingsType
'    stUNK = 0
'    stFileLevel = 1
'    stComputerLevel = 2
'End Enum

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
    ebTRUE = &H1
    ebFALSE = &H2
    ebPartial = &H4
    ebERROR = &H8
    ebNULL = &H10
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
    mrUnprotect = &H1
    mrClearFormatting = &H2
    mrClearContents = &H4
    mrMergeAcrossOnly = &H8
End Enum

Public Enum MergeFormatEnum
    mfUnknown = 0
    mfMerged = 1
    mfNotMerged = 2
    mfPartialMerged = 3
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
    
    Public Enum SheetProtection
        psOPTIONS_NOT_SET = 0
        psProtectContents = &H1
        psUsePassword = &H2
        psProtectDrawingObjects = &H4
        psProtectScenarios = &H8
        psUserInterfaceOnly = &H10
        psAllowFormattingCells = &H20
        psAllowFormattingColumns = &H40
        psAllowFormattingRows = &H80
        psAllowInsertingColumns = &H100
        psAllowInsertingRows = &H200
        psAllowInsertingHyperlinks = &H400
        psAllowDeletingColumns = &H800
        psAllowDeletingRows = &H1000
        psAllowSorting = &H2000
        psAllowFiltering = &H4000
        psAllowUsingPivotTables = &H8000&
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
        Red = 3
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
    
    Public Enum XMatchMode
    ''   DO NOT EDIT ENUM VALUES
        exactMatch = 0
        ExactMatchOrNextSmaller = -1
        ExactMatchOrNextLarger = 1
        WildcardCharacterMatch = 2
    End Enum
    
    Public Enum XSearchMode
    ''   DO NOT EDIT ENUM VALUES
        searchFirstToLast = 1
        searchLastToFirst = -1
        searchBinaryAsc = 2
        searchBinaryDesc = -2
    End Enum
    
    Public Enum ArrayOptionFlags
        aoNone = 0
        aoUnique = &H1
        aoUniqueNoSort = &H2
        aoAreaSizesMustMatch = &H4
        aoVisibleRangeOnly = &H8
        aoIncludeListObjHeaderRow = &H10
    End Enum

    Public Enum ListReturnType
        lrtArray = 1
        lrtDictionary = 2
        lrtCollection = 3
    End Enum

    Public Enum MinMax
        minValue = 1
        maxValue = 2
    End Enum
    
    Public Enum BeepType
        btDefault = 0
        btMsgBoxOK = &H1
        btMsgBoxChoice = &H2
        btError = &H4
        btBusyWait = &H8
        btButton = &H10
        btForced = &H20
    End Enum
    
    Public Enum PerfStatus
        psNONE = 0
        psEventsOn = &H1
        psEventsOff = &H2
        psScreenUpdateOn = &H4
        psScreenUpdateOff = &H8
        psInteractiveOn = &H10
        psInteractiveOff = &H20
        psAlertsOn = &H40
        psAlertsOff = &H80
        psPrintCommOn = &H100
        psPrintCommOff = &H200
        psPointerDefault = &H400
        psPointerBeam = &H800
        psPointerWait = &H1000
        psPointerNWArrow = &H2000
        psCalcAuto = &H4000
        psCalcManual = &H8000&
        psCalcSemiAuto = &H10000
        psStatusBarOn = &H20000
        psStatusBarOff = &H40000
    End Enum
    
    Public Enum PerfEnum
        
        peEventsOn = &H1
        peScreenUpdatingON = &H2
        peInteractiveON = &H4
        peMousePointerDefault = &H8
        peMousePointerWait = &H10
        peMousePointerBeam = &H20
        peMousePointerNWArrpw = &H40
        peCalcAutomatic = &H80
        peCalcManual = &H100
        peCalcSemiAutomatic = &H200
        peAlertsON = &H400
        pePrintCommunication = &H800
        
    End Enum
    
    Public Enum ErrorOptionsEnum
        ftDefaults = &H1
        ftERR_ProtectSheet = &H2
        ftERR_MessageIgnore = &H4
        ftERR_NoBeeper = &H8
        ftERR_DoNotCloseBusy = &H10
        ftERR_ResponseAllowBreak = &H20
    End Enum

    Public Enum CellErrorCheck
        ceFirstInRange = 1
        ceAnyInRange = 2
        ceAllInRange = 3
    End Enum

    Private lBypassOnCloseCheck As Boolean
    
    Public Property Get AppMode() As AppModeEnum
        AppMode = l_appMode
    End Property
    Public Property Let AppMode(appModeVal As AppModeEnum)
        If l_appMode <> appModeVal Then
            LogFORCED "App Mode Changing From: " & AppModeFriendly(l_appMode) & " to " & AppModeFriendly(appModeVal)
            l_appMode = appModeVal
        End If
    End Property
    Public Function AppModeFriendly(appModeVal As AppModeEnum) As String
        Select Case appModeVal
            Case AppModeEnum.appStatusStarting
                AppModeFriendly = "Starting"
            Case AppModeEnum.appStatusRunning
                AppModeFriendly = "Running"
            Case AppModeEnum.appStatusClosing
                AppModeFriendly = "Closing"
            Case Else
                AppModeFriendly = "Unknown"
        End Select
    End Function
    
    Public Property Get IsClosing() As Boolean
        IsClosing = (AppMode = AppModeEnum.appStatusClosing)
    End Property
    Public Property Let IsClosing(closing As Boolean)
            AppMode = appStatusClosing
    End Property
    Public Property Get pbPackageRunning() As Boolean
        pbPackageRunning = l_pbPackageRunning
    End Property
    Public Property Let pbPackageRunning(vl As Boolean)
        l_pbPackageRunning = vl
    End Property
    
    Public Function ArrayLength(arr, Optional arrayDimension)
        On Error Resume Next
        Dim arrLength As Long
        Dim lowerBound As Long, upperBound As Long
        If Not IsMissing(arrayDimension) Then
            lowerBound = LBound(arr, arrayDimension)
            upperBound = UBound(arr, arrayDimension)
        Else
            lowerBound = LBound(arr)
            upperBound = UBound(arr)
        End If
        arrLength = upperBound - lowerBound + 1
        If Err.number = 0 Then
            ArrayLength = arrLength
        Else
            ArrayLength = -1
        End If
    End Function
    
    Public Function MergeFormat(checkRange As Range) As MergeFormatEnum
        If checkRange Is Nothing Then
            MergeFormat = mfUnknown
        ElseIf IsNull(checkRange.MergeCells) Then
            MergeFormat = mfPartialMerged
        ElseIf checkRange.MergeCells = True Then
            MergeFormat = mfMerged
        ElseIf checkRange.MergeCells = False Then
            MergeFormat = mfNotMerged
        End If
    End Function
    
    

Public Function IsCellError(ByVal srcRng As Range, Optional cellCheck As CellErrorCheck = CellErrorCheck.ceFirstInRange) As Boolean
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Uses VBA.Information.IsError to check relevant cells in [srcRng] for Errors
''   If [cellCheck] = ceFirstInRange, returns True if FIRST cell in range is Error
''   If [cellCheck] = ceAnyInRange, returns True if ANY cell in range is Error
''   If [cellCheck] = ceAllInRange, returns True if ALL CELLS in range
''       areErrors
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

    Dim failed As Boolean
    Dim c
    If cellCheck = ceFirstInRange Then
        IsCellError = isError(srcRng(1, 1))
    ElseIf cellCheck = ceAnyInRange Then
        For Each c In srcRng
            If isError(c) Then
                IsCellError = True
                Exit For
            End If
        Next c
    ElseIf cellCheck = ceAllInRange Then
        Dim errCount As Long
        For Each c In srcRng
            If Not isError(c) Then
                IsCellError = False
                Exit For
            Else
                errCount = errCount + 1
            End If
        Next c
        If errCount = srcRng.Count Then
            IsCellError = True
        End If
    End If
End Function


Public Function SetRaiseEvents(Optional raiseEvts As Object) As Object
    Static raiseCommonEvents As Object
    If raiseCommonEvents Is Nothing Then
        If Not raiseEvts Is Nothing Then
            Set raiseCommonEvents = raiseEvts
        End If
    End If
    Set SetRaiseEvents = raiseCommonEvents
End Function
Public Function RaiseCommonEvent(eventName, ParamArray args() As Variant)
    Dim cRe As Object
    Set cRe = SetRaiseEvents
    If Not cRe Is Nothing Then
'        cRe.Raise eventName, args
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''       PERFORMANCE STATES
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  PERF MODE - DEFAULT
''  IF SETTINGS KEY "PERF_MODE_DEFAULT" IS SET, USE THAT VALUE, OTHERWISE
''  USE VALUE BELOW
''  TO SET IN SETTINGS, GET THE VALUE DESIRED BY:
''      Debug.Print [option] + [option] + ...
''      STG.Setting("PERF_MODE_DEFAULT") = PerfStatus.psAlertsOn + PerfStatus.psCalcAuto + _
''              PerfStatus.psEventsOn  + PerfStatus.psInteractiveOn  + _
''              PerfStatus.psPointerDefault  + PerfStatus.psPrintCommOff + _
''              PerfStatus.psScreenUpdateOn  + PerfStatus.psStatusBarOn
''      Debug.Print STG.Setting("PERF_MODE_DEFAULT")
''      '' > 296533
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function PerformanceDefault(Optional ByVal newDefault As PerfStatus = PerfStatus.psNONE, Optional resetDefault As Boolean = False) As PerfStatus
        Static tPerfStatus As PerfStatus
        Dim perfDefKey As String: perfDefKey = "PERF_MODE_DEFAULT"
        Dim tPerf As PerfStatus, validateEnum As Boolean
        If resetDefault Then
            tPerfStatus = 0
            newDefault = 0
            stg.Delete perfDefKey
        End If
        If newDefault > 0 Then
            tPerf = newDefault
            validateEnum = True
            '' dont forget to save to settings
        ElseIf tPerfStatus > 0 Then
            tPerf = tPerfStatus
        ElseIf stg.Exists(perfDefKey) Then
            If Not stg.SettingType(perfDefKey) = teNumeric Then
                stg.ForceNumericFormat (perfDefKey)
            End If
            tPerf = CLng(stg.Setting(perfDefKey))
            validateEnum = True
        Else
            tPerf = PerfStatus.psAlertsOn + PerfStatus.psCalcAuto + PerfStatus.psEventsOn _
            + PerfStatus.psInteractiveOn + PerfStatus.psPointerDefault + PerfStatus.psPrintCommOff _
            + PerfStatus.psScreenUpdateOn + PerfStatus.psStatusBarOn
        End If
        If validateEnum Then
            Dim perfVal As PerfStatus
            perfVal = tPerf
            '' MAKE SURE A FEW SPECIFIC OPTIONS HAVE 'USER WORKING' VALUES
            '' Events, ScreenUpdate, UserInteraction must all be true for default perf
            perfVal = EnumModify(newDefault, PerfStatus.psEventsOn, feVerifyEnumExists)
            perfVal = EnumModify(newDefault, PerfStatus.psEventsOff, feVerifyEnumRemoved)
            perfVal = EnumModify(newDefault, PerfStatus.psScreenUpdateOn, feVerifyEnumExists)
            perfVal = EnumModify(newDefault, PerfStatus.psScreenUpdateOff, feVerifyEnumRemoved)
            perfVal = EnumModify(newDefault, PerfStatus.psInteractiveOn, feVerifyEnumExists)
            perfVal = EnumModify(newDefault, PerfStatus.psInteractiveOff, feVerifyEnumRemoved)
            If perfVal <> tPerf Then
                stg.Setting(perfDefKey) = perfVal
            End If
            tPerfStatus = perfVal
        Else
            If tPerfStatus <> tPerf Then
                tPerfStatus = tPerf
                stg.Setting(perfDefKey) = tPerf
            End If
        End If
        PerformanceDefault = tPerfStatus
    End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''       CURRENT PERFORANCE SNAPSHOT
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function PerfOn(Optional disableAutoCalc As Boolean = True, _
    Optional mousePointerWait As Boolean = True) As PerfEnum
    Dim tPerf As PerfEnum
    
    If disableAutoCalc Then
        tPerf = EnumModify(tPerf, PerfEnum.peCalcManual, feVerifyEnumExists)
    Else
        tPerf = EnumModify(tPerf, PerfEnum.peCalcAutomatic, feVerifyEnumExists)
    End If
    If mousePointerWait Then
        tPerf = EnumModify(tPerf, PerfEnum.peMousePointerWait, feVerifyEnumExists)
    Else
        tPerf = EnumModify(tPerf, PerfEnum.peMousePointerDefault, feVerifyEnumExists)
    End If
    tPerf = EnumModify(tPerf, PerfEnum.peAlertsON, feVerifyEnumRemoved)
    tPerf = EnumModify(tPerf, PerfEnum.peEventsOn, feVerifyEnumRemoved)
    tPerf = EnumModify(tPerf, PerfEnum.peInteractiveON, feVerifyEnumRemoved)
    tPerf = EnumModify(tPerf, PerfEnum.peScreenUpdatingON, feVerifyEnumRemoved)
    
    Performance = tPerf
    PerfOn = tPerf
End Function
Public Function PerfOff()
    Performance = PerfNormal
End Function
Public Property Get Performance() As PerfEnum
    Dim retPer As PerfEnum
    If Events Then retPer = EnumModify(retPer, PerfEnum.peEventsOn, feVerifyEnumExists)
    If Screen Then retPer = EnumModify(retPer, PerfEnum.peScreenUpdatingON, feVerifyEnumExists)
    If Application.Interactive = True Then retPer = EnumModify(retPer, PerfEnum.peInteractiveON, feVerifyEnumExists)
    If Application.DisplayAlerts = True Then retPer = EnumModify(retPer, PerfEnum.peAlertsON, feVerifyEnumExists)
    
    If Application.Calculation = xlCalculationAutomatic Then
        retPer = EnumModify(retPer, peCalcAutomatic, feVerifyEnumExists)
    ElseIf Application.Calculation = xlCalculationSemiautomatic Then
        retPer = EnumModify(retPer, peCalcSemiAutomatic, feVerifyEnumExists)
    Else
        retPer = EnumModify(retPer, peCalcManual, feVerifyEnumExists)
    End If
    If Application.Cursor = xlDefault Then
        retPer = EnumModify(retPer, peMousePointerDefault, feVerifyEnumExists)
    ElseIf Application.Cursor = xlIBeam Then
        retPer = EnumModify(retPer, peMousePointerBeam, feVerifyEnumExists)
    ElseIf Application.Cursor = xlNorthwestArrow Then
        retPer = EnumModify(retPer, peMousePointerNWArrpw, feVerifyEnumExists)
    Else
        retPer = EnumModify(retPer, peMousePointerWait, feVerifyEnumExists)
    End If
    Performance = retPer
End Property

Public Property Let Performance(perfOptions As PerfEnum)
    On Error Resume Next
    
    Application.EnableEvents = EnumCompare(perfOptions, PerfEnum.peEventsOn)
    Application.ScreenUpdating = EnumCompare(perfOptions, PerfEnum.peScreenUpdatingON)
    Application.Interactive = EnumCompare(perfOptions, PerfEnum.peInteractiveON)
    Application.DisplayAlerts = EnumCompare(perfOptions, PerfEnum.peAlertsON)
    
    If IsDev Then
        Dim perfError As Boolean
        If Application.EnableEvents <> EnumCompare(perfOptions, PerfEnum.peEventsOn) Then perfError = True
        If Application.ScreenUpdating <> EnumCompare(perfOptions, PerfEnum.peScreenUpdatingON) Then perfError = True
        If Application.Interactive <> EnumCompare(perfOptions, PerfEnum.peInteractiveON) Then perfError = True
        If Application.DisplayAlerts <> EnumCompare(perfOptions, PerfEnum.peAlertsON) Then perfError = True
        If perfError Then
            Beep
            Stop
        End If
    End If
    
    If EnumCompare(perfOptions, PerfEnum.peMousePointerDefault) Then
        Application.Cursor = xlDefault
    ElseIf EnumCompare(perfOptions, PerfEnum.peMousePointerBeam) Then
        Application.Cursor = xlIBeam
    ElseIf EnumCompare(perfOptions, PerfEnum.peMousePointerNWArrpw) Then
        Application.Cursor = xlNorthwestArrow
    ElseIf EnumCompare(perfOptions, PerfEnum.peMousePointerWait) Then
        Application.Cursor = xlWait
    Else
        Application.Cursor = xlDefault
    End If
    If EnumCompare(perfOptions, PerfEnum.peCalcAutomatic) Then
        Application.Calculation = xlCalculationAutomatic
    ElseIf EnumCompare(perfOptions, PerfEnum.peCalcManual) Then
        Application.Calculation = xlCalculationManual
    ElseIf EnumCompare(perfOptions, PerfEnum.peCalcSemiAutomatic) Then
        Application.Calculation = xlCalculationSemiautomatic
    End If
    
    If IsDev Then
        If Err.number <> 0 Then
            Beep
            Stop
        End If
        If stg.Setting(STG_DEV_OUTPUT_TRACE) = True Then
            Application.StatusBar = GetPerfText
        End If
    End If
End Property
Public Property Get PerfBlockAll() As PerfEnum
    PerfBlockAll = PerfEnum.peCalcManual + peMousePointerWait
End Property
Public Property Get PerfBlockAllCalcAuto() As PerfEnum
    PerfBlockAllCalcAuto = PerfEnum.peCalcAutomatic + peMousePointerWait
End Property
Public Property Get PerfNormal() As PerfEnum
    PerfNormal = _
        PerfEnum.peAlertsON _
        + PerfEnum.peCalcAutomatic _
        + PerfEnum.peEventsOn _
        + PerfEnum.peInteractiveON _
        + PerfEnum.peMousePointerDefault _
        + PerfEnum.peScreenUpdatingON
End Property
Public Function GetPerfText(Optional notNormalOnly As Boolean = False) As String
    notNormalOnly = False
    
    
    Dim curPerf As PerfEnum: curPerf = Performance
    Dim c As New Collection, v As Variant, resp As Variant
    'CHECK EVENTS
    
    If EnumCompare(curPerf, PerfEnum.peEventsOn) = False Then
        c.Add "evts:N"
    ElseIf notNormalOnly = False Then
        c.Add "evts:Y"
    End If
    'CHECK SCREEN UPDATING
    If EnumCompare(curPerf, PerfEnum.peScreenUpdatingON) = False Then
        c.Add "scrn:N"
    ElseIf notNormalOnly = False Then
        c.Add "scrn:Y"
    End If
    'CHECK INTERACTIVE
    If EnumCompare(curPerf, PerfEnum.peInteractiveON) = False Then
        c.Add "intr:N"
    ElseIf notNormalOnly = False Then
        c.Add "intr:Y"
    End If
    'CHECK ALERTS
    If EnumCompare(curPerf, PerfEnum.peAlertsON) = False Then
        c.Add "alrt:N"
    ElseIf notNormalOnly = False Then
        c.Add "alrt:Y"
    End If
    'CHECK CALC MODE
    If EnumCompare(curPerf, PerfEnum.peCalcAutomatic) = False Then
        c.Add "calc:M"
    ElseIf notNormalOnly = False Then
        c.Add "calc:A"
    End If
    If EnumCompare(curPerf, PerfEnum.peMousePointerDefault) = False Then
        c.Add "curs:W"
    ElseIf notNormalOnly = False Then
        c.Add "curs:D"
    End If
    If notNormalOnly Then
        resp = " ( Abnormal States:"
    Else
        resp = " ( States:"
    End If
    For Each v In c
        resp = Concat(resp, " ", v)
    Next v
    resp = Concat(resp, " )")
    Set c = Nothing
    GetPerfText = resp
End Function
Public Function DEVListPerfStates()
    Dim info
    info = Concat("Events Enabled:" & Application.EnableEvents, vbNewLine)
    info = Concat(info, "Alerts Enabled:" & Application.DisplayAlerts, vbNewLine)
    info = Concat(info, "Screen Updating:" & Application.ScreenUpdating, vbNewLine)
    info = Concat(info, "Screen Interactive:" & Application.Interactive, vbNewLine)
    Dim tMouse, tCalc
    If Application.Cursor = xlDefault Then
        tMouse = "xlDefault"
    ElseIf Application.Cursor = xlWait Then
        tMouse = "xlWait"
    ElseIf Application.Cursor = xlIBeam Then
        tMouse = "xlIBeam"
    ElseIf Application.Cursor = xlNorthwestArrow Then
        tMouse = "xlNorthwestArrow"
    End If
    info = Concat(info, "Mouse: ", tMouse, vbNewLine)
    
    If Application.Calculation = xlCalculationAutomatic Then
        tCalc = "Automatic"
    ElseIf Application.Calculation = xlCalculationManual Then
        tCalc = "Manual"
    ElseIf Application.Calculation = xlCalculationSemiautomatic Then
        tCalc = "Semi-automatic"
    End If
    info = Concat(info, "Calculation: ", tCalc)
    Debug.Print info
    
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''       DATES & TIME UTILITIES
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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

Public Function dtPart(thePart As DateDiffType, dt1 As Variant, _
    Optional ByVal firstDayOfWeek As VbDayOfWeek = vbSunday, _
    Optional ByVal firstWeekOfYear As VbFirstWeekOfYear = VbFirstWeekOfYear.vbFirstJan1) As Variant
    Select Case thePart
        Case DateDiffType.dtDate_NoTime
            dtPart = DateSerial(dtPart(dtYear, dt1), dtPart(dtMonth, dt1), dtPart(dtDay, dt1))
        Case DateDiffType.dtDay
            dtPart = DatePart("d", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtDayOfYear
            dtPart = DatePart("y", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtHour
            dtPart = DatePart("h", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtMinute
            dtPart = DatePart("n", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtMonth
            dtPart = DatePart("m", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtQuarter
            dtPart = DatePart("q", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtSecond
            dtPart = DatePart("s", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtWeek
            dtPart = DatePart("ww", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtWeekday
            dtPart = DatePart("w", dt1, firstDayOfWeek, firstWeekOfYear)
        Case DateDiffType.dtYear
            dtPart = DatePart("yyyy", dt1, firstDayOfWeek, firstWeekOfYear)
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


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''   WORKSHEET PROTECTION
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Return Default 'SheetProtection' Flag Enum
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Get DefaultProtectOptions() As SheetProtection
    DefaultProtectOptions = _
        psAllowFiltering _
        + psAllowFormattingCells _
        + psAllowFormattingColumns _
        + psAllowFormattingRows _
        + psProtectDrawingObjects _
        + psUserInterfaceOnly _
        + psProtectContents _
        + psAllowSorting
End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  ** CurrentProtectionOptions **
''  Builds and Returns a 'SheetProtection' enum based on
''      current protection options of a protected worksheet
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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
            If .protectDrawingObjects Then spEnum = EnumModify(spEnum, SheetProtection.psProtectDrawingObjects, feVerifyEnumExists)
            If .protectContents Then spEnum = EnumModify(spEnum, SheetProtection.psProtectContents, feVerifyEnumExists)
            If .protectScenarios Then spEnum = EnumModify(spEnum, SheetProtection.psProtectScenarios, feVerifyEnumExists)
            
            If .protection.allowFormattingCells Then spEnum = EnumModify(spEnum, SheetProtection.psAllowFormattingCells, feVerifyEnumExists)
            If .protection.allowFormattingColumns Then spEnum = EnumModify(spEnum, SheetProtection.psAllowFormattingColumns, feVerifyEnumExists)
            If .protection.allowFormattingRows Then spEnum = EnumModify(spEnum, SheetProtection.psAllowFormattingRows, feVerifyEnumExists)
            If .protection.allowInsertingColumns Then spEnum = EnumModify(spEnum, SheetProtection.psAllowInsertingColumns, feVerifyEnumExists)
            If .protection.allowInsertingRows Then spEnum = EnumModify(spEnum, SheetProtection.psAllowInsertingRows, feVerifyEnumExists)
            If .protection.allowInsertingHyperlinks Then spEnum = EnumModify(spEnum, SheetProtection.psAllowInsertingHyperlinks, feVerifyEnumExists)
            If .protection.allowDeletingColumns Then spEnum = EnumModify(spEnum, SheetProtection.psAllowDeletingColumns, feVerifyEnumExists)
            If .protection.allowDeletingRows Then spEnum = EnumModify(spEnum, SheetProtection.psAllowDeletingRows, feVerifyEnumExists)
            If .protection.allowSorting Then spEnum = EnumModify(spEnum, SheetProtection.psAllowSorting, feVerifyEnumExists)
            If .protection.allowFiltering Then spEnum = EnumModify(spEnum, SheetProtection.psAllowFiltering, feVerifyEnumExists)
            If .protection.allowUsingPivotTables Then spEnum = EnumModify(spEnum, SheetProtection.psAllowUsingPivotTables, feVerifyEnumExists)
        End If
    End With

    CurrentProtectionOptions = spEnum

End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  ** UnprotectSheet **
''  Will use CFG_PROTECT_PASSWORD Constant if separate
''  Password is not supplied, and if a cached password
''  does not exist
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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
'    If CollectionKeyExists(protSessionSheets, wksht.CodeName) Then
'        protSessionSheets.Remove wksht.CodeName
'        LogTRACE "Removed " & wksht.CodeName & " from protSessionSheets collection"
'    End If
    
End Function
Public Function UnprotectAllSheets(Optional wkbk As Workbook, Optional unhideAll As Boolean = False)
    If wkbk Is Nothing Then
        Set wkbk = ThisWorkbook
    End If
    Dim tws As Worksheet
    For Each tws In wkbk.Worksheets
        UnprotectSheet tws
        If Not tws.visible = xlSheetVisible And unhideAll Then
            tws.visible = xlSheetVisible
        End If
    Next tws
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Returns Arraycached SheetProtection options, otherwise '0' (zero)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function CachedProtectionOptions(ByRef wksht As Worksheet) As Variant
    Dim item() As Variant, tKey
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
    
    protectDrawingObjects = EnumCompare(options, SheetProtection.psProtectDrawingObjects)
    protectContents = EnumCompare(options, SheetProtection.psProtectContents)
    protectScenarios = EnumCompare(options, SheetProtection.psProtectScenarios)
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
    
    '' just force this
    If userInterfaceOnly = False Then
        userInterfaceOnly = True
        If IsDev Then
            Beep
            Stop
        End If
    End If

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
    
'        If Not CollectionKeyExists(protSessionSheets, wksht.CodeName) Then
'            protSessionSheets.Add Now(), key:=wksht.CodeName
'            LogTRACE "Added " & wksht.CodeName & " to protSessionSheets collection"
'        Else
'            protSessionSheets.Remove (wksht.CodeName)
'            protSessionSheets.Add Now(), key:=wksht.CodeName
'            LogTRACE "Re-Added " & wksht.CodeName & " to protSessionSheets collection"
'        End If
        
    End With
    

End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   *** PUBLIC *** IMPLEMENTATION OF COMMON FUNCTIONS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get byPassOnCloseCheck() As Boolean
        byPassOnCloseCheck = lBypassOnCloseCheck
    End Property
    Public Property Let byPassOnCloseCheck(bypassCheck As Boolean)
        lBypassOnCloseCheck = bypassCheck
    End Property
    Public Property Get DevUserNames() As String()
        DevUserNames = stg.Setting(STG_DEV_NAMES)
    End Property
    Public Function DevUserNames_Add(ParamArray names() As Variant)
        stg.ArraySettingAppendTo STG_DEV_NAMES, False, False, names
    End Function
    Public Function DevUserNames_Remove(ParamArray removeNames() As Variant)
        stg.ArraySettingRemoveFrom STG_DEV_NAMES, False, False, removeNames
    End Function
    Public Function ConcatCollectionItems(ByRef col As Collection, Optional delimitBy As String = "|") As String
        Dim resp As String
        Dim colItem
        For Each colItem In col
            If resp = vbNullString Then
                resp = CStr(colItem)
            Else
                resp = ConcatWithDelim(delimitBy, resp, colItem)
            End If
        Next colItem
        ConcatCollectionItems = resp
    End Function
    Public Sub ftBeep(bpType As BeepType)
        Dim doBeep    As Boolean
        If EnumCompare(bpType, BeepType.btMsgBoxOK) Then
            doBeep = stg.SettingWithDefault("MESSAGE_BEEPS_INFO", True)
        ElseIf EnumCompare(bpType, BeepType.btMsgBoxChoice) Then
            doBeep = stg.SettingWithDefault("MESSAGE_BEEPS_CHOICE", True)
        ElseIf EnumCompare(bpType, BeepType.btButton) Then
            doBeep = stg.SettingWithDefault("MESSAGE_BEEPS_BUTTON", False)
        ElseIf EnumCompare(bpType, BeepType.btError) Then
            doBeep = stg.SettingWithDefault("MESSAGE_BEEPS_ERROR", True)
        ElseIf EnumCompare(bpType, BeepType.btForced) Then
            doBeep = True
        ElseIf EnumCompare(bpType, BeepType.btBusyWait) Then
            doBeep = stg.SettingWithDefault("MESSAGE_BEEPS_BUSY", False)
        Else
            doBeep = False
        End If
        If doBeep Then
            Beep
        End If
    End Sub
    Public Property Get AppVersionString() As String
        AppVersionString = Format(AppVersion, "0.00##")
    End Property
    Public Property Get AppVersion() As Double
        AppVersion = CDbl(stg.Setting("Version"))
    End Property
    Public Property Let AppVersion(vrsn As Double)
        stg.Setting("Version") = vrsn
    End Property
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  module pbCommonUtil.bas
''  author (c) Paul Brower https://github.com/lopperman/just-VBA
''  license GNU General Public License v3.0
''  DO NOT REMOVE THIS METHOD, I MEAN SERIOUSLY YOU'RE PROBABLY GETTING A
''   TON OF VALUE FROM ALL THIS FREE CODE I MADE AVAILABLE TO YOU, SO THE LEAST
''   YOU CAN DO IS LEAVE THIS METHOD IN TACT BECAUSE GRIDLINES ARE AWFUL!
''   OPTIONALLY, I ENCOURAGE YOU TO USE THIS METHOD!
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function AllUrGridlineBelongToUs(Optional requireConfirmation As Boolean = True)
    Dim doTheRightThing As Boolean
    doTheRightThing = True
    If requireConfirmation Then
        doTheRightThing = MsgBox_FT("Remove gridlines from all sheets in all open workbooks?" & vbNewLine & "Please say yes, please please say yes", vbCritical + vbYesNo + vbDefaultButton2, "DOING THE WORK OF A GOD") = vbYes
    End If
    If doTheRightThing Then
        RemoveGridLines allOpenWorkbooks:=True
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  module pbCommonUtil.bas
''  author (c) Paul Brower https://github.com/lopperman/just-VBA
''  license GNU General Public License v3.0
''  Change zoom on optional workbook otherwise all workbooks to be [entered]
''  percent for all visible worksheets
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function AllUrZoomTo(newZoom As Long, Optional wkbk As Workbook, Optional useViewName1, Optional useviewName2)
    Dim wb As Workbook, ws As Worksheet
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then
            If (Not wkbk Is Nothing And wb Is wkbk) Or wkbk Is Nothing Then
                For Each ws In wb.Worksheets
                    If ws.visible = xlSheetVisible Then
                        If ws.NamedSheetViews.Count > 0 Then
                            Dim switchToView As NamedSheetView
                            Set switchToView = Nothing
                            Dim nsv As NamedSheetView
                            For Each nsv In ws.NamedSheetViews
                                If Not IsMissing(useViewName1) Then
                                    If StringsMatch(nsv.Name, useViewName1) Then
                                        Set switchToView = nsv
                                    End If
                                End If
                                If Not IsMissing(useviewName2) Then
                                    If StringsMatch(nsv.Name, useviewName2) Then
                                        Set switchToView = nsv
                                    End If
                                End If
                            Next nsv
                            If Not switchToView Is Nothing Then
                                switchToView.Activate
                            End If
                        End If
                        
                        If StringsMatch(newZoom, ws.Parent.Windows(1).Zoom) = False Then
                            Debug.Print "changing zoom for: " & wb.Name & ": " & ws.Name
                            stg.Zoom False, newZoom, , ws
                        End If
                    End If
                Next ws
            End If
        End If
    Next wb
End Function

Public Function RemoveGridLines(Optional wkbkOrWksht, Optional allOpenWorkbooks As Boolean = False)
    On Error Resume Next
    Dim tmpWkbk As Workbook
    Dim tmpWksht As Worksheet
    Dim view As WorksheetView
    Dim tWin As Window
    
    If allOpenWorkbooks Then
        Dim aWB As Workbook
        For Each aWB In Application.Workbooks
            RemoveGridLines aWB
        Next
        Exit Function
    End If
    
    If StringsMatch(TypeName(wkbkOrWksht), "Workbook") Then
        Set tmpWkbk = wkbkOrWksht
    ElseIf StringsMatch(TypeName(wkbkOrWksht), "Worksheet") Then
        Set tmpWksht = wkbkOrWksht
    End If
    If tmpWkbk Is Nothing And tmpWksht Is Nothing Then
        Set tmpWkbk = ThisWorkbook
    End If
    If Not tmpWkbk Is Nothing Then
        For Each tWin In tmpWkbk.Windows
            For Each view In tWin.SheetViews
                If view.DisplayGridlines = True Then
                    view.DisplayGridlines = False
                End If
            Next
        Next
    ElseIf Not tmpWksht Is Nothing Then
        For Each tWin In tmpWksht.Parent.Windows
            For Each view In tWin.SheetViews
                If view.Sheet Is tmpWksht Then
                    If view.DisplayGridlines = True Then
                        view.DisplayGridlines = False
                    End If
                End If
            Next
        Next
    End If
    Set tmpWkbk = Nothing
    Set tmpWksht = Nothing
    If Err.number <> 0 Then Err.Clear
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   COPY ENTIRE WORKSHEET TO A DIFFERENT WORKBOOK
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function CopyWorksheetUtil(sourceWksht As Worksheet, targetWB As Workbook)
On Error GoTo e:
    Dim failed As Boolean
    Dim evts As Boolean: evts = Application.EnableEvents
    Dim srcVis As Variant
    srcVis = sourceWksht.visible
    
    Application.EnableEvents = False
    sourceWksht.visible = xlSheetVisible
    
    If Not targetWB Is Nothing Then
        With sourceWksht
            .Copy After:=targetWB.Worksheets(1)
            DoEvents
        End With
    End If

Finalize:
    On Error Resume Next
    sourceWksht.visible = srcVis
    If Not failed Then
        Beep
    End If
    Application.EnableEvents = evts
    
    Exit Function
e:
    failed = True
    sourceWksht.visible = srcVis
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
    
    If StringsMatch(strIn, "'", smContains) Then
        If padSingleQuote Then
            strIn = CleanSingleTicks(strIn)
        Else
            strIn = Replace(strIn, "'", "")
        End If
    End If
    
    ReplaceIllegalCharacters = strIn
End Function

Public Property Get IsDev() As Boolean
    If Not stg Is Nothing Then
        IsDev = stg.IsDeveloper
    End If
End Property
    
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''
    ''  CHECK IF A STRING CONTAINS 1 OR MORE STRING FOLLOWING EACH OTHER
    ''  @checkString = string that searching applies to (the 'haystack')
    ''  @sequences = ParamArray of strings in order to be searched (e.g. "A", "CD", "J")
    ''
    ''  EXAMPLES
    ''      searchStr = "ABCD(EFGGG) HIXXKAB"
    ''      Returns TRUE: = StringSequence(searchStr,"(",")")
    ''      Returns TRUE: = StringSequence(searchStr,"a","b","xx")
    ''      Returns TRUE: = StringSequence(searchStr,"a","b","b")
    ''      Returns TRUE: = StringSequence(searchStr,"EFG","GG")
    ''      Returns FALSE: = StringSequence(searchStr,"EFGG","GG")
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function StringSequence( _
        ByVal checkString, _
        ParamArray Search() As Variant) As Boolean
        Dim failed As Boolean
        Dim startPosition As Long: startPosition = 1
        Dim findString
        For Each findString In Search
            startPosition = InStr(startPosition, checkString, findString, vbTextCompare)
            If startPosition > 0 Then startPosition = startPosition + Len(findString)
            If startPosition = 0 Then failed = True
            If failed Then Exit For
        Next
        StringSequence = Not failed
    End Function
    
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''
    ''  CHECK IF A STRING CONTAINS 1 OR MORE STRING FOLLOWING EACH OTHER
    ''  @checkString = string that searching applies to (the 'haystack')
    ''  @sequences = ParamArray of strings in order to be searched (e.g. "A", "CD", "J")
    ''
    ''  Returns Custom Type: StringSequenceResult
    ''      : failed (true if any of the [search()] value were not found in sequence
    ''      : searchString (original string to be searched)
    ''      : failedAtIndex (if failed = true, failedAtIndex is the 1-based index for the first
    ''      :   failed search term
    ''      : results() (1-based, 2 dimension  variant array)
    ''      : results(1,1) = first searched term; results(1,2) = index where searched item was found
    ''      : results(2,1) = second searched term; results(2,2) = index where second item was found
    ''      :       etc
    ''      : Note: first searched item to fail get's 0 (zero) in the result(x,2) position
    ''      :   all search terms after the first failed search term, do not get searched,
    ''      :   so results(x,2) for those non-searched items is -1
    ''
    '' EXAMPLE USAGE:
    ''  Dim resp as StringSequenceResult
    ''  resp = StringSequence2("ABCDEDD","A","DD")
    ''  Debug.Print resp.failed (outputs: False)
    ''  Debug.Print resp.results(2,2) (outputs: 6)
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function StringSequence2( _
        ByVal checkString, _
        ParamArray Search() As Variant) As StringSequenceResult
        Dim resp As StringSequenceResult
        Dim startPosition As Long: startPosition = 1
        Dim findString, curIdx As Long
        resp.searchString = checkString
        ReDim resp.Results(1 To UBound(Search) - LBound(Search) + 1, 1 To 2)
        For Each findString In Search
            curIdx = curIdx + 1
            resp.Results(curIdx, 1) = findString
            If Not resp.failed Then
                startPosition = InStr(startPosition, checkString, findString, vbTextCompare)
            Else
                startPosition = -1
            End If
            resp.Results(curIdx, 2) = startPosition
            
            If startPosition > 0 Then
                startPosition = startPosition + Len(findString)
            Else
                If Not resp.failed Then
                    resp.failed = True
                    resp.failedAtIndex = curIdx
                End If
            End If
        Next
        StringSequence2 = resp
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
        For rRow = 1 To rng.Areas(rArea).rows.Count
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
Public Function Concat(ParamArray items() As Variant) As String
    Concat = Join(items, "")
End Function
Public Function Concat_1DArray(dArr, Optional delimiter As String = " | ")
On Error Resume Next
    Concat_1DArray = Join(dArr, delimiter)
End Function

Public Property Get ENV_User() As String
'    ENV_User = VBA.Interaction.Environ("USER")
    ENV_User = ENV_LogName
End Property

Public Function ENV_LogName() As String
    #If Mac Then
        ENV_LogName = ReplaceIllegalCharacters(VBA.Interaction.Environ("LOGNAME"), "", False)
    #Else
        ENV_LogName = ReplaceIllegalCharacters(VBA.Interaction.Environ("USERNAME"), "", False)
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
    If Not ThisWorkbook.activeSheet Is Nothing Then
        If Intersect(ThisWorkbook.Windows(1).VisibleRange, ThisWorkbook.activeSheet.Range(activeSheetAddress).Cells(1, 1)) Is Nothing Then
            InVisibleRange = False
        Else
            InVisibleRange = True
        End If
    End If
    
    If InVisibleRange = False And scrollTo = True Then
        Dim scrn As Boolean: scrn = Application.ScreenUpdating
        Application.ScreenUpdating = True
        Application.GoTo Reference:=ThisWorkbook.activeSheet.Range(activeSheetAddress).Cells(1, 1), Scroll:=True
        DoEvents
        Application.ScreenUpdating = scrn
    End If
    
    If Err.number <> 0 Then
        Err.Clear
    End If
End Function

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

Public Function isMac() As Boolean
'   Returns True If Mac OS
    #If Mac Then
        isMac = True
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  CLEAN SINGLE TICKS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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
    On Error GoTo e:
    wbName = CleanSingleTicks(wbName)
    Application.Run ("'" & wbName & "'!'" & procName & "'")
    Exit Function
e:
    ftBeep btError
    If Not raiseErrorOnFail Then
        Err.Clear
        On Error GoTo 0
    Else
        Err.Raise Err.number, Err.Description
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   FLAG ENUM COMPARE
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function EnumCompare(theEnum As Variant, enumMember As Variant, Optional ByVal iType As ecComparisonType = ecComparisonType.ecOR) As Boolean
    Dim c As Long
    c = theEnum And enumMember
    EnumCompare = IIf(iType = ecOR, c <> 0, c = enumMember)
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   FLAG ENUM - ADD/REMOVE SPECIFIC ENUM MEMBER
''   (Works with any flag enum)
''   e.g. If you have vbMsgBoxStyle enum and want to make sure
''   'DefaultButton1' is included
''   msgBtnOption = vbYesNo + vbQuestion
''   msgBtnOption = EnumModify(msgBtnOption,vbDefaultButton1,feVerifyEnumExists)
''   'now includes vbDefaultButton1, would not modify enum value if it already
''   contained vbDefaultButton1
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   INPUT BOX
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function InputBox_FT(ByVal prompt As String, Optional ByVal title As String = "Input Needed", Optional ByVal Default As Variant, Optional ByVal inputType As ftInputBoxType, Optional ByVal useVBAInput As Boolean = False) As Variant
    Dim evts As Boolean, curWB As Workbook
    Dim screenUpd As Boolean
    ''
    screenUpd = Application.ScreenUpdating
    evts = Application.EnableEvents
    Set curWB = Application.ActiveWorkbook
    ''
    ftBeep btMsgBoxChoice
    Application.EnableEvents = False
    Application.ScreenUpdating = True
    ''
    If IsMissing(inputType) Then inputType = ftibString
    If useVBAInput Then
        InputBox_FT = VBA.InputBox(prompt, title:=title, Default:=Default)
    Else
        InputBox_FT = Application.InputBox(prompt, title:=title, Default:=Default, Type:=inputType)
    End If
    curWB.Activate
    DoEvents
    Application.ScreenUpdating = screenUpd
    Application.EnableEvents = evts
    Set curWB = Nothing
    
End Function



' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   MESSAGE BOX REPLACEMENT
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function MsgBox_FT(prompt As String, Optional buttons As VbMsgBoxStyle = vbOKOnly, Optional title As Variant) As Variant
    Dim beeper As BeepType, resp As Variant
    Dim evts As Boolean, curWB As Workbook
    Dim screenUpd As Boolean
    ''
    screenUpd = Application.ScreenUpdating
    evts = Application.EnableEvents
    Set curWB = Application.ActiveWorkbook
    ''
    If EnumCompare(buttons, vbOKCancel) Then
        beeper = btMsgBoxChoice
    ElseIf EnumCompare(buttons, vbRetryCancel) Then
        beeper = btMsgBoxChoice
    ElseIf EnumCompare(buttons, vbYesNo) Then
        beeper = btMsgBoxChoice
    ElseIf EnumCompare(buttons, vbYesNoCancel) Then
        beeper = btMsgBoxChoice
    Else
        beeper = btMsgBoxOK
    End If
    ftBeep beeper
    Application.EnableEvents = False
    Application.ScreenUpdating = True
    resp = MsgBox(prompt, buttons, title)
    curWB.Activate
    DoEvents
    Application.ScreenUpdating = screenUpd
    Application.EnableEvents = evts
    Set curWB = Nothing
    MsgBox_FT = resp
End Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   ASK YES/NO
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   GET NEXT ID
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function GetNextID(listObj, ByVal ListCol) As Long
'   Use to create next (Long) number for unique ROW id in a Range
On Error Resume Next
    Dim nextId As Long
    Dim lo As ListObject, colIdx As Long, tmpListCol As listColumn
    
    If StringsMatch(TypeName(listObj), "ListObject") Then
        Set lo = listObj
    ElseIf StringsMatch(TypeName(listObj), "String") Then
        Set lo = wt(CStr(listObj))
    End If
    If Not lo Is Nothing Then
        If StringsMatch(TypeName(ListCol), "string") Then
            ListCol = CStr(ListCol)
        End If
        Set tmpListCol = lo.ListColumns(ListCol)
        If Not tmpListCol Is Nothing Then
            If lo.listRows.Count = 0 Then
                nextId = 1
            Else
                nextId = CLng(Application.WorksheetFunction.Max(tmpListCol.DataBodyRange)) + 1
            End If
        End If
    End If
    
    If Err.number = 0 And nextId > 0 Then
        GetNextID = nextId
    ElseIf Err.number <> 0 Or lo Is Nothing Or tmpListCol Is Nothing Then
        If lo Is Nothing Then
            Logger.LogERROR "GetNextID failed for invalid listobject "
        Else
            Logger.LogERROR "GetNextID failed for listobject " & lo.Name & ", column: " & ListCol
        End If
        GetNextID = -1
        If Err.number <> 0 Then Err.Clear
    End If
    
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  REF: https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity
''      Application.AutomationSecurity returns or sets an MsoAutomationSecurity constant
''          that represents the security mode that Microsoft Excel uses when
''          programmatically opening files. Read/write.
''  Excel Automatically Defaults Application.AutomationSecurity to msoAutomationSecurityLow
''  If you are programatically opening a file and you DO NOT want macros / VBA to run
''      in that file, use this method to open workbook
''  NOTE: This method prevents 'auto_open' from running in workbook being opened
''
''  Usage:
''      [fullPath] = fully qualified path to excel file
''          If path contains spaces, and is an http path, spaces are automatically encoded
''      [postOpenSecurity] (Optional) = MsoAutomationSecurity value that will be set AFTER
''          file is opened.  Defaults to Microsoft Defaul Value (msoAutomationSecurityLow)
''      [openReadOnly] (Optional) = Should Workbook be opened as ReadOnly. Default to False
''      [addMRU] (Optional) = Should file be added to recent files list. Default to False
''      Returns Workbook object
Public Function OpenWorkbookDisabled(ByVal fullPath As String, _
    Optional ByVal postOpenSecurity As MsoAutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityLow, _
    Optional ByVal openReadOnly As Boolean = False, _
    Optional ByVal addMRU As Boolean = False, _
    Optional ByVal returnFocusCurWkbk As Boolean = True) As Workbook
    ''
    On Error Resume Next
    Dim currentEventsEnabled As Boolean
    ''  GET CURRENT EVENTS ENABLED STATE
    currentEventsEnabled = Application.EnableEvents
    ''  DISABLE APPLICATION EVENTS
    Application.EnableEvents = False
    ''  ENCODE FILE PATH IF NEEDED
    If InStr(1, fullPath, "http", vbTextCompare) = 1 And InStr(1, fullPath, "//", vbTextCompare) >= 5 Then
        fullPath = Replace(fullPath, " ", "%20", compare:=vbTextCompare)
    End If
    ''  PREVENT MACROS/VBA FROM RUNNING IN FILE THAT IS BEING OPENED
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    ''  OPEN FILE
    Set OpenWorkbookDisabled = Workbooks.Open(fullPath, ReadOnly:=openReadOnly, addToMRU:=addMRU)
    ''  RESTORE EVENTS TO PREVIOUS STATE
    If returnFocusCurWkbk Then
        ThisWorkbook.Activate
    End If
    Application.EnableEvents = currentEventsEnabled
    ''  RESTORE APPLICATION.AUTOMATIONSECURITY TO [postOpenSecurity]
    Application.AutomationSecurity = postOpenSecurity
End Function

Public Function GetWorksheet(findName As String, Optional wkbk As Workbook, Optional findByCodeName As Boolean = False) As Worksheet
    Dim searchWkbk As Workbook, searchWS As Worksheet, isMatch As Boolean
    If Not wkbk Is Nothing Then
        Set searchWkbk = wkbk
    Else
        Set searchWkbk = ThisWorkbook
    End If
    For Each searchWS In searchWkbk.Worksheets
        If findByCodeName = True Then
            If StringsMatch(searchWS.CodeName, findName) Then
                isMatch = True
            End If
        Else
            If StringsMatch(searchWS.Name, findName) Then
                isMatch = True
            End If
        End If
        If isMatch Then
            Set GetWorksheet = searchWS
            Exit For
        End If
    Next searchWS
End Function

Public Function WorkbookIsOpen(ByVal wkbkName As String, Optional checkCodeName As String = vbNullString) As Boolean
On Error Resume Next
    Dim tmpFound As Boolean
    wkbkName = FileNameFromFullPath(wkbkName)
    Dim wkbk As Workbook
    For Each wkbk In Application.Workbooks
        If StringsMatch(wkbk.Name, wkbkName) Then
            tmpFound = True
            Exit For
        End If
    Next wkbk
    If Not tmpFound Then
        Dim tWB As Workbook
        Set tWB = Workbooks(wkbkName)
        If Not tWB Is Nothing Then
            If StringsMatch(tWB.Name, wkbkName) Then
                tmpFound = True
            End If
        End If
        Set tWB = Nothing
    End If
    ''
    If Err.number <> 0 Then
        tmpFound = False
    End If
    If WorkbookIsOpen And Len(checkCodeName) > 0 Then
        If StringsMatch(Workbooks(wkbkName).CodeName, checkCodeName) Then
            tmpFound = True
        Else
            tmpFound = False
        End If
    End If
    If Err.number <> 0 Then Err.Clear
    WorkbookIsOpen = tmpFound
    Exit Function
End Function

Public Function FirstMondayOfMonth(ByVal dtVal As Variant) As Variant
    dtVal = CDate(dtVal)
    Dim firstOfMonth As Variant, tMonday As Variant
    firstOfMonth = DateSerial(DatePart("yyyy", dtVal), DatePart("m", dtVal), 1)
    tMonday = GetMondayOfWeek(firstOfMonth)
    If DatePart("m", firstOfMonth) = DatePart("m", tMonday) Then
        FirstMondayOfMonth = tMonday
    Else
        FirstMondayOfMonth = DateAdd("d", 7, tMonday)
    End If
End Function

Public Function GetSundayOfWeek(ByVal inputDate As Variant) As Date
    inputDate = CDate(inputDate)
    Dim processDt As Variant
    If TypeName(inputDate) = "String" Then
        processDt = DateValue(inputDate)
    Else
        processDt = inputDate
    End If
    If Not dtPart(dtWeekday, processDt, firstDayOfWeek:=vbMonday) = 7 Then
        processDt = DtAdd(dtDay, 7 - dtPart(dtWeekday, processDt, firstDayOfWeek:=vbMonday), processDt)
    End If
    GetSundayOfWeek = processDt
End Function

Public Function GetMondayOfWeek(ByVal inputDate As Variant) As Date
    inputDate = CDate(inputDate)
    Dim processDt As Variant
    If TypeName(inputDate) = "String" Then
        processDt = DateValue(inputDate)
    Else
        processDt = inputDate
    End If
    Dim diffDays As Long
    diffDays = 1 - dtPart(dtWeekday, processDt, firstDayOfWeek:=vbMonday)
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
    If Not ThisWorkbook.activeSheet Is Nothing Then
        ActiveSheetName = ThisWorkbook.activeSheet.Name
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






' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''   FILE SYSTEM FUNCTIONS
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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

Public Function OpenPath(fldrPath As String)
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Create Valid File or Directory Path (for PC or Mac, local,
''   or internet) from 1 or more arguments
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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
    Dim tPath As String, tFileName As String, fNameStarts As Long
    tFileName = FileNameFromFullPath(fullFileName)
    fNameStarts = InStr(fullFileName, tFileName)
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

Public Function FileExtension(ByVal fileName) As String
    If Len(fileName) > 0 Then
        Dim perPos As Long
        perPos = InStrRev(fileName, ".", , vbTextCompare)
        If perPos > 0 Then
            FileExtension = Mid(fileName, perPos + 1)
        End If
    End If
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

Public Function ChooseFile(choosePrompt As String, Optional ByVal fileExt As String = vbNullString) As String
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
        ChooseFile = CStr(retV)
    Else
        Debug.Print "ERROR pbCommon.ChooseFile "
        Err.Clear
    End If

End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   CREATE THE ** LAST ** DIRECTORY IN 'fullPath'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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
        If InStrRev(fullPath, Application.PathSeparator, compare:=vbTextCompare) > InStr(1, fullPath, Application.PathSeparator, vbTextCompare) Then
            lastDirName = Left(fullPath, InStrRev(fullPath, Application.PathSeparator, compare:=vbTextCompare) - 1)
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Returns true if filePth Exists and is not a directory
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Returns true if DIRECTORY path dirPath) Exists
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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

Public Function VisibleWorksheets(Optional wkbk As Workbook) As Long
    Dim w As Worksheet, retV As Long
    If wkbk Is Nothing Then
        Set wkbk = ThisWorkbook
    End If
    For Each w In wkbk.Worksheets
        If w.visible = xlSheetVisible Then
            retV = retV + 1
        End If
    Next
    VisibleWorksheets = retV
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   MAKE SURE ALL 'NAMES' ARE VISIBLE IN NAMES MANAGER
''    - ALL WORKSHEETS MUST BE UNPROTECTED
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function DEVMakeAllNamesVisible()
    'This makes all the names visible, then can manually delete from
    '   Formulas --> Name Manager
    Dim nm As Name
    For Each nm In ThisWorkbook.names
        nm.visible = True
    Next nm
    MsgBox "Check Formulas --> Name Manager to view names"
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Returns item from collection by Key
''   If [key] does not exist in collection, error object with
''   error code 1004 is return
''   suggested use:
''
''   Dim colItem as Variant
''   colItem = CollectionItemByKey([collection], [expectedKey])
''
''   'If expecting object, use 'Set'
''    Set colItem = CollectionItemByKey([collection], [expectedKey])
''
''   If Not IsError(colItem) Then
''       'value was returned
''   Else
''       'error was returned
''   End if
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function CollectionItemByKey(ByRef col As Collection, ByVal key)
On Error Resume Next
    key = CStr(key)
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
Public Function CollectionKeyExists(ByRef col As Collection, ByVal key)
On Error Resume Next
    key = CStr(key)
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   returns delimited string with non-object collection items
''   e.g.
''   Dim c as New Collection
''   c.Add "A"
''   c.Add Now
''   c.add 42.55
''   Debug.Print CollectionToString(c)
''   ''Outputs:  "A|5/28/23 7:24:53 PM|42.55"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function CollectionToString( _
    ByRef coll As Variant, _
    Optional delimiter As String = "|") As String
    Dim retStr As String, colItem As Variant
    Dim evalItem As Variant
    For Each colItem In coll
        evalItem = vbEmpty
        If TypeName(colItem) = "Range" Then
            evalItem = colItem.value
        ElseIf Not IsObject(colItem) Then
            evalItem = colItem
        End If
        If Len(evalItem) > 0 Then
            If Len(retStr) = 0 Then
                retStr = CStr(evalItem)
            Else
                retStr = retStr & delimiter & CStr(evalItem)
            End If
        End If
    Next colItem
    CollectionToString = retStr
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Get All  'Environ' Settings, or
''   Get 'Environ' Settings matches or partially matches 'filter'
''   filter applied to [Key] by default, to apply to value(s), use
''       filterKey:=False
''   Returns Collection of 1-based Arrays (1=Key, 2=Val) with [Key] = Environ Key, and
''       [Item] = Environ Value
''   If using filterKey with exactMatch:=True, will return 1 or no matching items
''   When mutiple items are matched (exactMatch:=False),
''       will return 0 to Many items
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   FYI: TO GET USERNAME:
''       ON PC USE:  =GetEnvironSettings("username")
''       ON MAC USE: =GetEnvironSettings("user")
''
''   (VBA.Interaction.Environ([KEY] is case sensitive, however using this
''       'GetEnvironSettings' is not case sensitive)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   EXAMPLE OF USAGE
''       ** Print all Environ Key/Value values to Immediate Window
''       Dim v
''       For Each v In GetEnvironSettings
''           Debug.Print v(1), v(2) 'Key, Value
''       Next v
''       ** Print all Environ Key/Value values
''       Dim v
''       For Each v In GetEnvironSettings
''           Debug.Print v(1) & ", " & v(2)
''       Next v
''       ** Get Environ setting value for key=USER
''       Dim v, tempCol as Collection
''       Set tempCol = GetEnvironSettings("user", exactMatch:=True)
''           THIS
''       If tempCol.Count = 1 Then
''           Debug.Print tempCol(1)(2)
''       End if
''           OR THIS
''       v = tempCol(1)
''           Debug.Print v(2)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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

Public Function DEV_WriteEnvironSettings()
    Dim stgItem
    For Each stgItem In GetEnvironSettings
        Debug.Print stgItem(LBound(stgItem)) & ": " & stgItem(UBound(stgItem))
    Next
    
End Function

Public Property Get UserName() As String
    UserName = stg.UserNameOrLogin
End Property

Public Function NowWithMS() As String
    NowWithMS = Format(Now, "yyyymmdd hh:mm:ss") & Right(Format(Timer, "0.000"), 4)
End Function

Public Function SysStates() As String
    Dim tEv As String, tSc As String, tIn As String, tCa As String, retV As String, tAl As String
    tEv = "Evt:" & IIf(Events, "Y  ", "N  ")
    tSc = "Scr:" & IIf(Application.ScreenUpdating, "Y  ", "N  ")
    tIn = "Int:" & IIf(Application.Interactive, "Y  ", "N  ")
    tCa = "Clc:" & IIf(Application.Calculation = xlCalculationAutomatic, "A  ", IIf(Application.Calculation = xlCalculationManual, "M  ", "SA  "))
    tAl = "Alr:" & IIf(Application.DisplayAlerts, "Y", "N")
    retV = Concat(tEv, tSc, tIn, tCa, tAl)
    retV = Concat("Sys: (", retV, ")")
    SysStates = retV
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Get any list object in workbook by name
''   by default, will cache reference to list object
''   set [cacheReference] to False if looking for a list object
''   that is known to be temporary
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function wt(lstObjName As String, Optional cacheReference As Boolean = True) As ListObject
On Error GoTo e:
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
    
    Exit Function
e:
    RaiseError Err.number, Err.Description
    Err.Clear
End Function



' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Scroll any active sheet to desired location
''    - Does not change previous worksheet selection
''    - Optionally set selection range, if desired ('selectRng')
''
''   Can use for scrolling only, worksheets do not have to have split panes
''
''   Use 'splitOnRow' and/or 'splitOnColumn' to guarantee split is correct
''    - By default split panes will be frozen.  Pass in arrgument: 'freezePanes:=False'
''      to make sure split panes are not frozen
''
''   By Default, if a splitRow/Column is not specific, but one existrs, it will be
''   left alone.  To remove split panes that should not exist by default,
''   pass in 'removeUnspecified:=True'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'Public Function Scroll(wksht As Worksheet _
'    , Optional SplitOnRow As Long _
'    , Optional splitOnColumn As Long _
'    , Optional freezePanes As Boolean = True _
'    , Optional removeUnspecified As Boolean _
'    , Optional selectRng As Range)
'    ' --- '
'    On Error GoTo E:
'    'The Worksheet you are scrolling must be the ActiveSheet'
'    If Not ActiveWindow.activeSheet Is wksht Then Exit Function
'    ' --- '
'    Dim failed As Boolean
'    Dim evts As Boolean, scrn As Boolean, scrn2 As Boolean
'    evts = Application.EnableEvents
'    scrn = Application.ScreenUpdating
'    scrn2 = Application.Interactive
'    ' --- '
'    Dim pnIdx As Long
'    With ActiveWindow
'        'Scroll All Panes to the left, to the top'
'
'        For pnIdx = 1 To .Panes.Count
''            .SmallScroll ToRight:=1
''            .SmallScroll Down:=1
'            If pnIdx = 1 Or pnIdx = 3 Then
'                .Panes(pnIdx).ScrollRow = 1
'            ElseIf .SplitRow > 0 Then
'                .Panes(pnIdx).ScrollRow = .SplitRow + 1
'            End If
'            If pnIdx = 1 Or pnIdx = 2 Then
'                .Panes(pnIdx).ScrollColumn = 1
'            ElseIf .splitColumn > 0 Then
'                .Panes(pnIdx).ScrollColumn = .splitColumn + 1
'            End If
'        Next pnIdx
'        'Ensure split panes are in the right place
'        If SplitOnRow > 0 And Not .SplitRow = SplitOnRow Then
'            .SplitRow = SplitOnRow
'        ElseIf SplitOnRow = 0 And .SplitRow <> 0 And removeUnspecified Then
'            .SplitRow = 0
'        End If
'        If splitOnColumn > 0 And Not .splitColumn = splitOnColumn Then
'            .splitColumn = splitOnColumn
'        ElseIf splitOnColumn = 0 And .splitColumn <> 0 And removeUnspecified Then
'            .splitColumn = 0
'        End If
'        If splitOnColumn > 0 Or SplitOnRow > 0 Then
'            If Not .freezePanes = freezePanes Then
'                .freezePanes = freezePanes
'            End If
'        End If
'    End With
'    If Not selectRng Is Nothing Then
'        If selectRng.Worksheet Is wksht Then
'            selectRng.Select
'        End If
'    End If
'Finalize:
'    On Error Resume Next
'    Application.EnableEvents = evts
'    Application.ScreenUpdating = scrn
'    Application.Interactive = scrn2
'    Application.StatusBar = False
'    Exit Function
'E:
'    'Implement Own Error Handling'
'    failed = True
'    MsgBox "Error in 'Scroll' Function: " & Err.number & " - " & Err.Description
'    Err.Clear
'    Resume Finalize:
'End Function

Public Function LastPopulatedColumn(wks As Worksheet) As Long
    If wks.UsedRange Is Nothing Then
        Exit Function
    End If
    Dim colOffset As Long, lastCol As Long
    colOffset = wks.UsedRange.column - 1
    lastCol = wks.UsedRange.Columns.Count + colOffset
    If WorksheetFunction.CountA(wks.Columns(lastCol)) > 0 Then
        LastPopulatedColumn = lastCol
    Else
        Do While WorksheetFunction.CountA(wks.Columns(lastCol)) = 0
            If lastCol = 1 Then Exit Do
            lastCol = lastCol - 1
        Loop
        LastPopulatedColumn = lastCol
    End If
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  SAFE CHECK IF 'stg' IS PROPERLY INSTANTIATED
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Get stgVALID() As Boolean
    On Error Resume Next
    If Not stg Is Nothing Then
        stgVALID = stg.ValidConfig
    End If
    If Err.number <> 0 Then
        Err.Clear
    End If
End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  pbSettings Primary Accessor
''
''  NOTE: pbSettings functions as a singleton and cannot be instantiated separately
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function stg() As pbSettings
    On Error Resume Next
    Static stgObj As pbSettings
    If stgObj Is Nothing Then
        Set stgObj = New pbSettings
    End If
    If Err.number = 0 Then
        If Not stgObj Is Nothing Then
            If stgObj.ValidConfig Then
                Set stg = stgObj
            End If
        End If
    Else
        Set stg = Nothing
        Err.Clear
    End If
End Function

Public Property Get Setting(ByVal stgKey) As Variant
    Setting = stg.Setting(stgKey)
End Property
Public Property Let Setting(ByVal stgKey, ByVal stgVal)
    stg.Setting(stgKey) = stgVal
End Property

Public Function pbCheckDefaultSettings()
    '' if missing, add default 'btnExit' action to be: 'QuitApp'
    If Len(stg.stdButtonAction(wsDashboard, "btnExit")) = 0 Then
        stg.stdButtonAction(wsDashboard, "btnExit") = "QuitApp"
    End If
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'
'       ERROR HANDLING
'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function ErrString(Optional customSrc As String, Optional errNUM As Variant, Optional errDESC As Variant, Optional errERL As Variant) As String
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
End Function
'Called When Must End
Private Function FatalEnd()
'    RaiseEvent
    Err.Raise 1004, Description:="FatalEnd Not Implemented"
End Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   RAISE ERROR
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Sub RaiseError(errNumber As Long, Optional errorDesc As String = vbNullString, Optional resetUserInterface As Boolean = True)
    '''''LogERROR Concat("RaiseError Called: ", errNumber, " - errorDesc: ", errorDesc)
''    pbCommonEvents.RaiseOnErrorHandlerBegin errNumber, errorDesc
    
    Dim cancelRaise As Boolean
    If IsDev Then
        
        Beep
        Stop
        
    End If
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
    On Error Resume Next
        If IsDev Then
            Beep
            Stop
'            Resume
'            Exit Function
        End If
        Logger.LogERROR Concat(errNumber, " - ", errDESC, " - ", Source)
    
    On Error GoTo 0
    
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'
'       PB COMMON LOGGING
'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   returns error object if not using pbcommong_log
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'Public Function pbLOGPath(Optional wkbk As Workbook) As Variant
'    If PBCOMMON_LOG = False Then
'        pbLOGPath = CVErr(1028)
'    Else
'        Dim logFileName As String
'        If wkbk Is Nothing Then
'            logFileName = FileNameWithoutExtension(ThisWorkbook.Name)
'        Else
'            logFileName = FileNameWithoutExtension(wkbk.Name)
'        End If
'        pbLOGPath = PathCombine(False, Application.DefaultFilePath _
'            , LOG_DIR _
'            , ConcatWithDelim("_", logFileName, "LOG", Format(Date, "YYYYMMDD") & ".log"))
'    End If
'End Function
'Private Function pbLogDirectory() As String
'    pbLogDirectory = PathCombine(True, Application.DefaultFilePath, LOG_DIR)
'End Function

'Public Function pbLogOpen()
'    If PBCOMMON_LOG = False Then Exit Function
'    'if the log file is open and this method is called, will close and then re-open
'    If val(pbLogFileNumber) > 0 Then
'        Close #pbLogFileNumber
'    End If
'    pbLogFileNumber = FreeFile
'    Dim LogPath As String
'    LogPath = CStr(pbLOGPath)
'    On Error Resume Next
'    Open pbLOGPath For Append As #pbLogFileNumber
'    If Err.number = 75 Then
'        Err.Clear
'        On Error GoTo 0
'        CreateDirectory PathCombine(True, Application.DefaultFilePath, LOG_DIR)
'        Open pbLOGPath For Append As #pbLogFileNumber
'    End If
'End Function
'Public Function pbLogClose()
'    If PBCOMMON_LOG = False Then Exit Function
'    If val(pbLogFileNumber) > 0 Then
'        Close #pbLogFileNumber
'        DoEvents
'        pbLogFileNumber = Empty
'    End If
'End Function

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
'Public Function pbLOG(msg, Optional lType As LogType = LogType.ltTRACE, Optional prependTimeStamp As Boolean = True, Optional closeLog As Boolean = True)
'    If PBCOMMON_LOG = False Then Exit Function
'    If lType = ltWARN Then
'        lwarnCount = lwarnCount + 1
'    ElseIf lType = ltERROR Then
'        lerrorCount = lerrorCount + 1
'    End If
'    msg = CStr(msg)
'    If lType >= LogLevel Then
'        If val(pbLogFileNumber) = 0 Then
'            pbLogOpen
'        End If
'        msg = ConcatWithDelim(" | ", LogTypeDesc(lType), msg)
'        If prependTimeStamp Then
'            msg = ConcatWithDelim(" | ", NowWithMS, msg)
'        End If
'        Write #pbLogFileNumber, msg
'        If closeLog Then
'            pbLogClose
'        End If
'    End If
'End Function
'Public Property Get LogWarningCount() As Long
'    LogWarningCount = lwarnCount
'End Property
'Public Property Get LogErrorCount() As Long
'    LogErrorCount = lerrorCount
'End Property

Public Property Get devMode() As Boolean
    If IsDev Then
        If stg.Exists("DEV_MODE") Then
            devMode = CBool(stg.Setting("DEV_MODE"))
        End If
    End If
End Property
Public Property Let devMode(ByVal dvMode As Boolean)
    stg.Setting("DEV_MODE") = dvMode
    If dvMode = False Then
        DevEventsDisabled = False
        stg.Setting(STG_DEV_PAUSELOCKING) = False
    End If
End Property


Public Function Logger(Optional lgLevel) As pbLOG
Static lg As pbLOG

    If lg Is Nothing Then
        On Error Resume Next
        Set lg = New pbLOG
        If IsMissing(lgLevel) Then
            lgLevel = LogType.ltWARN
        End If
        lg.InitLogging CLng(lgLevel), commonEvts:=pbCommonEvents
        If lg.Initialized = False Or Err.number <> 0 Then
            Set lg = Nothing
        End If
        If Err.number <> 0 Then
            Err.Clear
        End If
    End If
    
    If Not lg Is Nothing Then
        If Not IsMissing(lgLevel) Then
            If CLng(lgLevel) <> lg.logLevel Then
                lg.logLevel = CLng(lgLevel)
            End If
        End If
        Set Logger = lg
    End If
    
End Function
Public Property Get CanLog() As Boolean
    CanLog = (IsClosing = False)
    If CanLog Then
        If Logger Is Nothing Then CanLog = False
    End If
End Property
Public Function LogTRACE(msg, Optional keepLogOpen As Boolean = False)
    If CanLog Then
        Logger.LogTRACE msg, keepLogOpen
    End If
End Function
Public Function LogDEBUG(msg, Optional keepLogOpen As Boolean = False)
    If CanLog Then
        Logger.LogDEBUG msg, keepLogOpen
    End If
End Function
Public Function LogWARN(msg, Optional keepLogOpen As Boolean = False)
    If CanLog Then
        Logger.LogWARN msg, keepLogOpen
    End If
End Function
Public Function LogERROR(msg, Optional keepLogOpen As Boolean = False)
    If CanLog Then
        Logger.LogERROR msg, keepLogOpen
    End If
End Function
Public Function LogFORCED(msg, Optional keepLogOpen As Boolean = False)
    If CanLog Then
        Logger.LogFORCED msg, keepLogOpen
    End If
End Function
Public Property Let logLevel(lgLevel As LogType)
    If CanLog Then
        Logger.logLevel = lgLevel
    End If
End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''       QUIT OR CLOSE APP
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      QUIT APP
''      CALL DIRECTLY OR CALLED VIA WORKBOOK_BEFORECLOSE
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function QuitApp(Optional askUser As Boolean = True)
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
        '' DO NOT DELETE
        ''
        '' BELOW CODE NEEDS TO GO IN CODE-BEHIND AREA FOR WORKBOOK
            'Private Sub Workbook_BeforeClose(Cancel As Boolean)
            '    If AppMode = appStatusRunning Then
            '        Cancel = True
            '        QuitApp
            '    End If
            'End Sub
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Performance = PerfNormal
    Dim wbCount: wbCount = Application.Workbooks.Count
    Dim doClose As Boolean
    If askUser = False Then
        doClose = True
    Else
        If AskYesNo("Close and Save " & ThisWorkbook.Name & "?", "Exit") = vbYes Then
            doClose = True
        End If
    End If
    If doClose Then
        IsClosing = True
        If IsClosing Then
            With ThisWorkbook
                .Save
            End With
            If wbCount = 1 Then
                Application.Quit
            Else
                ThisWorkbook.Close SaveChanges:=True
            End If
        End If
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Hides all sheets that do not meet any of the following
'       conditions:
'   1.  Is not included in the [keepSheetsVisible()] parameter
'           (Can be a Worksheet object, a worksheet code-name,
'           or a worksheet name)
'   2.  Is not the Active Worksheet (ThisWorkbook.ActiveSheet)
'   3.  Is not part of the 'STG_DO_NOT_HIDE' FileSetting
'
'   Usage:
'       HideSheets
'       HideSheets Sheet1, Sheet2
'       HideSheets Sheet1.CodeName, Sheet2.CodeName
'       HideSheets "Sheet1", "Sheet2"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function HideSheets(ParamArray doNotHide() As Variant)
    If stg.stdDEVIsDevMode Then
        Exit Function
    End If
    
    Dim DoNotHideSheets As New Collection
    Dim tmpCodeName
    For Each tmpCodeName In doNotHide
        DoNotHideSheets.Add tmpCodeName, key:=CStr(tmpCodeName)
    Next
    
    Dim p As PerfEnum
    p = Performance
    Performance = PerfBlockAll
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not (ws Is ThisWorkbook.activeSheet) Then
            If Not CollectionKeyExists(DoNotHideSheets, ws.CodeName) Then
                If Not ws.visible = xlSheetVeryHidden Then
                    ws.visible = xlSheetVeryHidden
                End If
            Else
                If Not ws.visible = xlSheetVisible Then
                    ws.visible = xlSheetVisible
                End If
            End If
        End If
    Next ws
    Performance = p
End Function

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
        If Not ws.visible = xlSheetVisible Then
            If Not CollectionKeyExists(hideColl, ws.CodeName) Then
                ws.visible = xlSheetVisible
            End If
        End If
    Next ws
    If hideColl.Count > 0 Then
        For Each ws In wkbk.Worksheets
            If CollectionKeyExists(hideColl, ws.CodeName) Then
                If ws.visible = xlSheetVisible Then
                    If stg.stdDEVIsDevMode = False Then
                        ws.visible = xlSheetVeryHidden
                    End If
                End If
            End If
        Next ws
    End If
    
    Events = evts
    Screen = scrn
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   FORCED WAIT PERIOD for [xSeconds] (max 30 seconds)
''   ** Without DoEvents **
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function ForceWait(xSeconds As Long)
    If xSeconds > CLng(30) Then xSeconds = CLng(30)
    If CLng(xSeconds) > 0 Then
        Application.Wait (Now + TimeValue("0:00:" & Format(xSeconds, "00")))
    End If
End Function

Public Function OpenUrl(webAddress As String) As Boolean
    OpenUrl = FollowUrl(webAddress)
End Function
Public Function FollowUrl(URL As String) As Boolean
On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.FollowHyperlink Address:=URL
    FollowUrl = Err.number = 0
    If Err.number <> 0 Then Err.Clear
    Application.DisplayAlerts = True
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  FORMULAS
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function CheckRangeFormula(rng As Range, theFormula, Optional refStyle As XlReferenceStyle = XlReferenceStyle.xlA1, Optional canModify As Boolean = True) As Boolean
On Error Resume Next
    ''  why do we check so many 'formula' properties?
    ''  Add the following formula to a cell:  =SQRT(A1:A4)
    ''  These would be the property values of that cell:
    ''      .Formula =              "=SQRT(A1:A4)"
    ''      .Formula2 =            "=SQRT(@A1:A4)"
    ''      .FormulaR1C1 =      "=SQRT(R[-1]C[-3]:R[2]C[-3])"
    ''      .Formula2R1C1 =    "=SQRT(@R[-1]C[-3]:R[2]C[-3])"
    Dim fmlaMatch As Boolean, resp As Boolean
    ''''''If EnumCompare(vbvar TypeName(theFormula),
    fmlaMatch = StringsMatch(rng.Formula, theFormula)
    If Not fmlaMatch Then
        If refStyle = xlA1 Then
            fmlaMatch = StringsMatch(rng.Formula2, theFormula)
        Else
            fmlaMatch = StringsMatch(rng.FormulaR1C1, theFormula)
            If Not fmlaMatch Then fmlaMatch = StringsMatch(rng.Formula2R1C1, theFormula)
        End If
    End If
    If canModify = False Or fmlaMatch = True Then
        CheckRangeFormula = fmlaMatch
    ElseIf canModify = True And fmlaMatch = False Then
        If refStyle = xlA1 Then
            rng.Formula = theFormula
        Else
            rng.Formula2R1C1 = theFormula
        End If
        resp = CheckRangeFormula(rng, theFormula, refStyle, False)
        CheckRangeFormula = resp
    End If
    
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  CONVERT LONG TO FRIENDLY RGB
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Function ConvertColorToRgb(ColorValue As Long) As String
    Dim Red As Long, Green As Long, Blue As Long
    Red = ColorValue Mod 256
    Green = ((ColorValue - Red) / 256) Mod 256
    Blue = ((ColorValue - Red - (Green * 256)) / 256 / 256) Mod 256
    
    ConvertColorToRgb = "RGB(" & _
                    Red & ", " & _
                    Green & ", " & _
                    Blue & ")"
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  GENERAL WORKBOOK INFO
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function HiddenWorksheetCount( _
    Optional wkbk As Workbook, _
    Optional includeHidden As Boolean = True, _
    Optional includeVeryHidden As Boolean = True _
    ) As Long
    If wkbk Is Nothing Then Set wkbk = ThisWorkbook
    Dim tws As Worksheet
    Dim resp As Long
    For Each tws In wkbk.Worksheets
        If tws.visible = xlSheetHidden And includeHidden Then
            resp = resp + 1
        ElseIf tws.visible = xlSheetVeryHidden And includeVeryHidden Then
            resp = resp + 1
        End If
    Next tws
    HiddenWorksheetCount = resp
End Function
Public Function HiddenWorksheets( _
    Optional wkbk As Workbook, _
    Optional includeHidden As Boolean = True, _
    Optional includeVeryHidden As Boolean = True _
    ) As Collection
    ''
    If wkbk Is Nothing Then Set wkbk = ThisWorkbook
    Dim tws As Worksheet
    Dim resp As New Collection
    For Each tws In wkbk.Worksheets
        If tws.visible = xlSheetHidden And includeHidden Then
            resp.Add tws, key:=tws.Name
        ElseIf tws.visible = xlSheetVeryHidden And includeVeryHidden Then
            resp.Add tws, key:=tws.Name
        End If
    Next tws
    Set HiddenWorksheets = resp
End Function

Public Function ProtectedWorksheetCount(Optional wkbk As Workbook) As Long
    If wkbk Is Nothing Then Set wkbk = ThisWorkbook
    Dim tws As Worksheet
    Dim resp As Long
    For Each tws In wkbk.Worksheets
        If tws.protectContents Or tws.protectDrawingObjects Then
            resp = resp + 1
        End If
    Next tws
    ProtectedWorksheetCount = resp
End Function
Public Function ProtectedWorksheets( _
    Optional wkbk As Workbook _
    ) As Collection
    ''
    If wkbk Is Nothing Then Set wkbk = ThisWorkbook
    Dim tws As Worksheet
    Dim resp As New Collection
    For Each tws In wkbk.Worksheets
        If tws.protectContents Or tws.protectDrawingObjects Then
            resp.Add tws, key:=tws.Name
        End If
    Next tws
    Set ProtectedWorksheets = resp
End Function


Public Function BusyWait(msg As String, Optional waitSeconds As Long = 0, Optional ignoreBeep As Boolean = True, Optional ignoreIfHidden As Boolean = False, Optional quietMode As Boolean = False, Optional forceByDefault As Boolean = True, Optional LogDEBUG As Boolean = False)
On Error Resume Next
    If AppMode = appStatusStarting Then ignoreIfHidden = False

    If ignoreIfHidden And wsBusy.visible <> xlSheetVisible Then Exit Function
    
    If Not wsBusy Is ThisWorkbook.activeSheet Then
        wsBusy.Show msg
    Else
        wsBusy.PrivBusyWait msg, waitSeconds, ignoreBeep
    End If

End Function

'Conditional Compiler Args
'conLocal=1
Public Property Get LocalMode() As Boolean
    #If conLocal Then
        LocalMode = True
        If IsDev Then
            Debug.Print Concat(Now, " -- ", "LOCAL MODE: TRUE")
        End If
    #Else
        LocalMode = False
    #End If
End Property

