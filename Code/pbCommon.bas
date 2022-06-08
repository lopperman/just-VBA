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

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'   GENERALIZED TYPES
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
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
Public Type UiState
    Events As Boolean
    Interactive As Boolean
    Screen As Boolean
    alerts As Boolean
    Cursor As XlMousePointer
    calc As XlCalculation
End Type

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'   GENERALIZED ENUMS
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '


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
Public Enum ftOperatingState
    [_ftunknown] = -1
    ftOpening = 0
    ftRunning = 1
    ftClosing = 2
    ftUpgrading = 3
    ftResetting = 4
    ftImporting = 5
End Enum

Private l_preventProtection As Boolean

Public Property Let PreventProtection(preventProtect As Boolean)
    l_preventProtection = preventProtect
End Property
Public Property Get PreventProtection() As Boolean
    PreventProtection = l_preventProtection
End Property


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~    SYSTEM STATE    ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function RestoreStateDefault()
    RestoreState DefaultState
End Function
Private Property Get DefaultState() As UiState
    Dim retV As UiState
    retV.alerts = True
    retV.calc = xlCalculationAutomatic
    retV.Cursor = XlMousePointer.xlDefault
    retV.Events = True
    retV.Interactive = True
    retV.Screen = True
    DefaultState = retV
End Property
Public Function CurrentUIState() As UiState
'   Get Current 'UI' State Settings
    Dim retV As UiState
    retV.alerts = Application.DisplayAlerts
    retV.calc = Application.Calculation
    retV.Cursor = Application.Cursor
    retV.Events = Application.EnableEvents
    retV.Interactive = Application.Interactive
    retV.Screen = Application.ScreenUpdating
    CurrentUIState = retV
End Function
Public Function RestoreState(prevState As UiState)
'   Update Current State Settings to values in 'prevState'
    SetState prevState
End Function
Public Function SetState(updState As UiState)
    With Excel.Application
        If Not .Interactive = updState.Interactive Then .Interactive = updState.Interactive
        If Not .ScreenUpdating = updState.Screen Then .ScreenUpdating = updState.Screen
        If Not .Cursor = updState.Cursor Then .Cursor = updState.Cursor
        If Not .Calculation = updState.calc Then .Calculation = updState.calc
        If Not .DisplayAlerts = updState.alerts Then .DisplayAlerts = updState.alerts
        If Not .EnableEvents = updState.Events Then .EnableEvents = updState.Events
    End With
End Function
Public Function SuspendState( _
    Optional scrnUpd As Boolean = False, _
    Optional scrnInter As Boolean = True, _
    Optional alerts As Boolean = False, _
    Optional calcMode As XlCalculation = xlCalculationAutomatic _
    ) As UiState
'   Returns current UIState (before suspension options updated)
    Dim retV As UiState
    retV = CurrentUIState
    With Excel.Application
        If .EnableEvents Then .EnableEvents = False
        If Not .ScreenUpdating = scrnUpd Then .ScreenUpdating = scrnUpd
        If Not .Interactive = scrnInter Then .Interactive = scrnInter
        If Not .DisplayAlerts = alerts Then .DisplayAlerts = alerts
        If Not .Calculation = calcMode Then .Calculation = calcMode
    End With
    SuspendState = retV
End Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~   (END) SYSTEM STATE    ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~


