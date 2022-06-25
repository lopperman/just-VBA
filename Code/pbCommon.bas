Attribute VB_Name = "pbCommon"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' pbCommon v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' Enums, Constants, Types, Common Utilities
'
' @module pbCommon
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'   GENERALIZED CONSTANTS
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Public Const CFG_PROTECT_PASSWORD As String = "00000"
Public Const CFG_PROTECT_PASSWORD_EXPORT As String = "000001"
Public Const CFG_PROTECT_PASSWORD_MISC As String = "0000015"
Public Const CFG_P_LOG As String = "0000016"

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'   GENERALIZED TYPES
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Public Enum ecComparisonType
    ecOR = 0 'default
    ecAnd
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

Public Enum StringMatch
    smEqual = 0
    smNotEqualTo = 1
    smContains = 2
    smStartsWithStr = 3
    smEndWithStr = 4
End Enum

Public Type ArrInformation
    Rows As Long
    Columns As Long
    Dimensions As Long
    Ubound_first As Long
    LBound_first As Long
    UBound_second As Long
    LBound_second As Long
End Type
Public Type AreaStruct
    RowStart As Long
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

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'   ENUMS FOR WORKING WITH pbPerf
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'       ___ ___ ___ ___ ___ ___ ___ ___ ___ ___ ___ ___
'           Flag Enum Items that are used when calling
'           SetPerf(options as ftPerStates)
'           FlagEnums can be 'added' (like how the Msgbox arguments work)
Public Enum ftPerfOption
    poINVALID = 0
    ' CLEAR Control -- returns to default user interaction mode
            poClearControl = 2 ^ 0
        ' Use any combination of these option with 'poClearControl'
        ' E.g.:    SetPerf poClearControl + poForeFinalSheet + poKeepTraceQueued
            poIgnoreSheetProtect = 2 ^ 1 ' Do not Protect Active Sheet
            poKeepTraceQueued = 2 ^ 2 ' Do not write out queued Trace Info
            poForceFinalSheet = 2 ^ 3 ' For User to 'Land' On Specific Sheet
            poBypassCloseChecks = 2 ^ 4 ' Disable Any 'OnClosing' Checking (In case of Error 51 -- Internal Error)
    
    ' SUSPEND Control -- Default 'turns everything off, but you can
    ' selective adjust if needed
            poSuspendControl = 2 ^ 5
     ' Use any combination of these options with 'poSuspendControl'
     ' E.g.:    SetPerf poSuspendControl + poCalcModeManual + poDoNotDisable_Alerts
            poCalcModeManual = 2 ^ 6 ' Keep Calculation (Defaults to Automantic) on Manual Mode on next 'poClearControl' is called
            poDoNotDisable_Screen = 2 ^ 7 ' Do Not Disable Screen During 'SuspendControl'
            poDoNotDisable_Interaction = 2 ^ 8 ' Do Not Disable User Interaction During 'SuspendControl'
            poDoNotDisable_Alerts = 2 ^ 9 ' Do Not Disable Alerts During 'Suspend Control
            
    ' CHECK/VALIDATE SuspendControl -- This ensures all the Application 'Performance' properties
    '   are set to the  values applied during the last ' SetPerf poSuspendControl'
    '   It's common to need to change a performance property while code is running -- for example there may
    '   be a screen update you wish to perform. After the code requiring manual adjustments to these properties
    '   rather than needing to remember what to set everything back to , simply call ' SetPerf poCheckControl '
    '   poCheckControl can be used as often as needed while the Default 'SuspendControl' or
    '   custom 'SuspendControl' values are being used. (Returns to 'Default User Control' any time the 'SetPerf poClearControl' is called
            poCheckControl = 2 ^ 10
    
    ' MISC. Control Operation
    '   The 'Performance' State Only has 2 'modes':
    '    - Default User Interaction mode (Events ON, Screen Updating ON, etc)
    '    - SuspendControl mode (Using either the Default SuspendControl configuration (Which allows tweaking some properties), or
    '    - a custom ftPerfStates Type that you have set. (See the 'SetPerfCustom' Function for more into on that
    '  The the current Performance Mode is in 'SuspendControl', an error will be raised if you try to set a new 'SuspendControl' configuration
    '  In the vent you need to change the performance options in the middle of your code, you can call 'SetPerf poOverride' to clear
    '  the current PerfState before applying a new one.
            poOverride = 2 ^ 11
End Enum
        
Public Type ftPerfStates
    IsPerfState As Boolean
    IsDefault As Boolean
    Events As Boolean
    Interactive As Boolean
    Screen As Boolean
    alerts As Boolean
    Cursor As XlMousePointer
    calc As XlCalculation
End Type
Public Enum ModifySuspendState
    EnableEvents = 2 ^ 0
    EnableInteractive = 2 ^ 1
    EnableScreenUpdate = 2 ^ 2
    EnableAlerts = 2 ^ 3
    CalculationAuto = 2 ^ 4
    CalculationManual = 2 ^ 5
End Enum

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'   GENERALIZED ENUMS
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
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

Public Enum ProtectionPWD
    pwStandard = 1
    pwExport = 2
    pwMisc = 3
    pwLog = 4
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

Public Enum RangeFunctionOperator
    Min = 1
    Max = 2
    Sum = 3
    Count = 4
    CountUnique = 5
    countBlank = 6
End Enum

Public Enum btnLocationEnum
    Beneath = 1
    ToTheRight
End Enum

Public Enum LogLevelEnum
    info = 1
    warning = 2
    Error = 3
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
    ftibFormula = 0
    ftibNumber = 1
    ftibString = 2
    ftibLogicalValue = 4
    ftibCellReference = 8
    ftibErrorValue = 16
    ftibArrayOfValues = 64
End Enum

Public Enum BusyState
    bsUnknown = -1
    bsOpening = 1
    bsClosing = 2
    bsRunning = 3
End Enum

Public Enum ListReturnType
    lrtArray = 1
    lrtDictionary = 2
    lrtCollection = 3
End Enum

Public Enum AllocationReportType
    artFirstOfMonth = 1
    artLastOfMonth = 2
    artFirstOrLastOfMonth = 3
End Enum

Public Enum XMatchMode
    ExactMatch = 0
    ExactMatchOrNextSmaller = -1
    ExactMatchOrNextLarger = 1
    WildcardCharacterMatch = 2
End Enum

Public Enum XSearchMode
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
    'implement aoAreaMustMatchRows
    'implement aoAreasMustMatchCols
    aoVisibleRangeOnly = 2 ^ 3
End Enum

Public Enum CHSuppEnum
    chNONE = 0
    chForceFullUpdate = 2 ^ 0
    chUpdateAllColumns = 2 ^ 1
    chSpecificRange = 2 ^ 2
End Enum

Public Enum TempFolderEnum
    tfSettings = 1
    tfDeploymentFiles = 2
    tfProdRelease = 3
    tfBetaRelease = 4
    tfTestRelease = 4 'this is not a mistake!
End Enum
Public Enum ftTrigger
    ftButtonAction = 1
    ftUserEvent = 2
End Enum
Public Enum ftMinMax
    minValue = 1
    maxValue = 2
End Enum
Public Enum HolidayEnum
    holidayName = 1
    holidayDt = 2
End Enum


Private l_preventProtection As Boolean
Private l_OperatingState As ftOperatingState


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

