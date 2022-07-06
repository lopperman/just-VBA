VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbPerf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' NOTE: THIS CLASS REQUIRES VB_PredeclaredId = True
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' pbPerf v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
'   Manage Application Settings from a single location to improve performance while
'   code is executing. (Includes suspending events)
'
'   If you need something re-enabled (e.g. ScreenUpdate) while your code is running,
'   you don't need to worry about changing the setting back. Just place the following
'   call at the end of your method: 'pbPerf.Check'
'   'pbPerf.Check' will verify all Application Setting are in the Propery 'Suspended' State
'
'       ~~~  USAGE STRATEGY  ~~~
'       There are typically 3 types of actions that can start code running in your app:
'       (1) User Clicking A Control that has a Macro Assign (Or by Running a Macro),
'       (2) Automatic Code Execution, like 'Auto_Run' Macro, 'OnTime' Application
'            Call, Workbook_Open, etc
'       (3) Event triggered by user interaction with Worksheet Objects (like
'            Double-Clicking a range, Changing A Range, Etc)
'       If you are able to have a clear 'Starting' and 'Ending' Path for your code,
'           this Module will enable you to 'turn things off or on' one time for
'           each time code is tirggered and executed.
'       (This means all the hundreds of places you have code doing things like:
'           Application.EnableEvents = True/False, Application.ScreenUpdating = True/False
'           ALL that code can (and should be) delete, and replaced with One Call
'           to this Module when your code starts, and one call to this module when
'           your code ends.
'
'       ~~~ 'EASY MODE' ~~~
'       To Simplify the use of this module, and still get all the benefits ofo having
'       Application states managed in one place, use the 3 methods below
'
'       1. (SUSPEND) SuspendMode(Optional calc As XlCalculation = XlCalculation.xlCalculationAutomatic)
'            - Syntax:  pbPerf.SuspendMode
'            - Place the App in 'Performance' Mode By Turning off EnableEvents, ScreenUpdating,
'              User Interaction, and various Animations
'
'       2. (CHECK) Check
'            - Syntax: pbPerf.Check
'            - If Current Application Mode is 'Default', this will Put the App into Suspend Mode
'            - If Current Application MOde is 'Suspend' then this will verify all the Application
'              Settting are set correct.  (If you had changed something in your code. like ScreenUpdating,
'              this method will change it back
'
'       3. (CHECK) DefaultMode
'            - Syntax: pbPerf.DefaultMode
'                       Optional ignoreProtect As Boolean = False, _
'                       Optional ignoreDumpTrace As Boolean = False, _
'                       Optional forceSheet As Boolean = False, _
'                       Optional enableCloseBypass As Boolean = False
'            - This method return Application to DefaultMode (Events On, User Interacation Restored)

' @module pbPerf
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1

Private lPerfState As ftPerfStates
Private Const PRIVATE_ERR_CLASS_INSTANCE_PROHIBITED = vbObjectError + 1022

'~~~ ~~~ A 'Home Worksheet' Is Required! ~~~ ~~~'
'   * Set the Homesheet Worksheet Name by assigning it to the FORCED_WORKSHEET_NAME constant
'       underneath these comments *
'   This Module (pbPerf) is Configured as a Singleton (Attribute VB_PredeclaredId = True) and may
'       not be manually instantiated.  To use it, you just access by 'pbPerf.[whatever]'
'   The Private 'Configure' Method will Run the first time this module is accessed, and an Error
'       will occur if the homesheet is not able to be set to a valid worksheet.
'   This module will ensure the 'Homesheet' Is Visible And
'   is the ActiveWorksheet under the following conditions:
'   1. When Updating To the DefaultState (by calling 'pbPerf.DefaultMode' and including 'forceSheet:=True')
'   2. When Updating To the DefaultState (by calling 'pbPerf.PerfState' AND include 'poClearControl' and
'       'poForceFinalSheet' in the options)
'   3. If, when finished restoring to Default State, the ActiveWorksheet is a sheet name listed
'       in the l_hideSheetsOnDefault collection.  (a collection of 0 or more worksheets that should
'       only be visible to users while code is running and app is busy (like a wait screen)

Private Const FORCED_WORKSHEET_NAME As String = "Dashboard"
Private l_homeSheet As Worksheet

'   ~~~ ~~~ Hide Sheets On Default State ~~~ ~~~'
'   ** IF ** you have Sheets that should always be hidden when in DefaultMode, Add those
'       ** Worksheet Names As the Key to the l_HideSheets HashList **
'       ** See the 'Configure' Method For Usage Example **
'   When returning to DefaultMode (user can interact with app), if any Worksheets listed in the
'       l_hideSheets HashList Are Visible, they will be automatically hidden. (These would typically
'       be 'please wait' or 'busy' sheets that are show to users when they must wait for processing
'       to complete.
'   If any Worksheet in the l_hideSheets HashList are the ACTIVE Worksheet when restoring to
'       DefaultMode, then the FORCED_WORKSHEET_NAME Worksheet will be made Active

Private l_hideSheets As New cHashList

Private Function Configure()
On Error GoTo E:
    Dim failed As Boolean
    
'   Set the HomeSheet Worksheet
    If Len(FORCED_WORKSHEET_NAME) > 0 Then
        If WorksheetExists(FORCED_WORKSHEET_NAME) Then
            Set l_homeSheet = ThisWorkbook.Worksheets(FORCED_WORKSHEET_NAME)
        End If
    End If
    
    l_hideSheets.CompareMode = vbTextCompare

'   ~~~ ~~~ OPTIONAL ~~~ ~~~
'   Set the name/s of Sheets to Hide when returning to DefaultMode
    'EXAMPLE
    AddHideSheet "FinTool"
    AddHideSheet "Busy"

Finalize:
    On Error Resume Next
    If failed Then
        RestoreDefaultAppSettingsOnly
    End If
            
    Exit Function
E:
    failed = True
    ErrorCheck "pbPerf.Configure"
    Resume Finalize:
    
End Function

Private Function AddHideSheet(wsName As String)
    If Len(wsName) > 0 Then
        If WorksheetExists(wsName) Then
            l_hideSheets.add wsName, key:=wsName
        End If
    End If
End Function

' ________________________________________
'  ~~~~ ~~~~ ~~~ ~~~ EASY MODE ~~~ ~~~ ~~~ ~~~

    Public Function DefaultMode(Optional ignoreProtect As Boolean = False, _
        Optional ignoreDumpTrace As Boolean = False, _
        Optional forceSheet As Boolean = False, _
        Optional enableCloseBypass As Boolean = False)
        
        PerfStateClear doNotProtect:=ignoreProtect, doNotDumpTrace:=ignoreDumpTrace, forceSheet:=forceSheet, byPassCloseChk:=enableCloseBypass
    
    End Function

    Public Function SuspendMode(Optional calc As XlCalculation = XlCalculation.xlCalculationAutomatic)
        PerfState ftPerfOption.poSuspendControl + IIf(calc = xlCalculationManual, poCalcModeManual, 0)
    End Function

    Public Function Check()
        '   If current Perf Mode is Default then will put into 'default' Suspend Mode
        '   If current Perf Mode is PerfState (lPerfState.IsPerfState), then
        '       will validate and update if needed to make application settings match values in lPerfState
        CheckState
    End Function

' ________________________________________

Public Property Get IsInPerfState() As Boolean
'   Returns TRUE if things are locked down for performance, other FALSE
    IsInPerfState = lPerfState.IsPerfState
End Property

'   ~~~ ~~~   PERF STATE   ~~~ ~~~
'   PerfState Function is used to track and manage the following:
'   1. Events, On or Off (Application.EnableEvents)
'   2. Screen Updates, On or Off (Application.ScreenUpdating)
'   3. User Interaction, On or Off (Application.Interactive)
'   4. Mouse Cursor Display, Busy/Wait or Default (Application.Cursor)
'   5. Alerts, On or Off (Application.DisplayAlerts)
'   6. Calculation Mode, Manual or Automatic (Application.Calculation)
'   7. Menu Animation, On or Off (Application.EnableAnimations)
'   8. Print Communication, On or Off (Application.PrintCommunication)
'   9. Macro Animations, On or Off (Application.EnableMacroAnimations)
Public Function PerfState(options As ftPerfOption) As Boolean
''  TODO: ADD ERR RAISE IF INVALID OPTIONS INCLUDED (NOW JUST DROPPED ON FLOOR)

    '4 Key Action Types (23 in this function, 1 goes to the 'PerfStateCustom', 1 goes to "CheckState")
    '    - Clear Control and Return to Default
    '    - Verify Control (will make sure settings match the control -- in case some manually adjusts in their code
    '    - Add Default Control (SuspendControl)
    '    - Add Custom Control
    
        If EnumCompare(options, ftPerfOption.poOverride) Then
            lPerfState.IsPerfState = False
        End If
        
        '   ~~~ ~~~ CLEARING CONTROL ~~ ~~~
        If EnumCompare(options, ftPerfOption.poClearControl) Then
                Dim tmpDisableProt As Boolean, tmpBypassCloseChk As Boolean, tmpDoNotDump As Boolean, tmpForce As Boolean
                'Default calc to Automatic
                If EnumCompare(options, ftPerfOption.poIgnoreSheetProtect) Then tmpDisableProt = True
                If EnumCompare(options, ftPerfOption.poKeepTraceQueued) Then tmpDoNotDump = True
                If EnumCompare(options, ftPerfOption.poBypassCloseChecks) Then tmpBypassCloseChk = True
                If EnumCompare(options, ftPerfOption.poForceFinalSheet) Then tmpForce = True
        
                PerfStateClear doNotProtect:=tmpDisableProt, doNotDumpTrace:=tmpDoNotDump, forceSheet:=tmpForce, byPassCloseChk:=tmpBypassCloseChk
                
        '   ~~~ ~~~ ADDING CONTROL (SUSPEND) ~~ ~~~
        ElseIf EnumCompare(options, ftPerfOption.poSuspendControl) Then
                
                Dim tmpScreen As Boolean, tmpInter As Boolean, tmpAlert As Boolean, tmpCalc As XlCalculation
                'Default calc to Automatic
                tmpCalc = xlCalculationAutomatic
                If EnumCompare(options, ftPerfOption.poCalcModeManual) Then tmpCalc = xlCalculationManual
                If EnumCompare(options, ftPerfOption.poDoNotDisable_Alerts) Then tmpAlert = True
                If EnumCompare(options, ftPerfOption.poDoNotDisable_Screen) Then tmpScreen = True
                If EnumCompare(options, ftPerfOption.poDoNotDisable_Interaction) Then tmpInter = True
                SuspendState scrnUpd:=tmpScreen, scrnInter:=tmpInter, alerts:=tmpAlert, calcMode:=tmpCalc
        
        ElseIf EnumCompare(options, ftPerfOption.poCheckControl) Then
            CheckState
        End If
    
        
    

End Function

Public Function PerfStateCustom(cstmState As ftPerfStates, Optional overRideControl As Boolean = False)
'   Allows to set a custom 'ftPerfStates' as the Current 'Control'
'   'Control' Means that app is doing something that -- for performance reasons or otherwise -- requires disabling
'   Applicatioin features that are typically need for a user to interact with Excel
'   Assigning a Control Implies preventing typical user behavior.  The user will have to wait until the process has completed
'   Before they are able to resume interacting with Excel
'   ** NOTE ** To return the app back to it's normal 'user interaction mode', call the PerfState Function and
'       include the 'poClearControl' ftPerfOption  (e.g.:pbPerf.DefaultMode)

'   This Function should only be used when needing to add a Custom PerfState. Review the Private Function 'SuspendState'
'   And use that instead (by calling 'pbPerf.SuspendMode'). Automatic Workbook Calculation can be suspended
'   with 'poSuspendControl' by adding this addtional ftPerfOption:
'       pbPerf.SuspendMode + poCalcModeManual
'   If addtional customization are needed you can set the Control with a custom ftPerfStates by call this Function.
'   ** WARNING ** A Control cannot be assigned if there is an existing control already in effect. If needed,
'   The previous Control can be overriden by chaning the 'overRideControl' argument to True, like the following example:
'       PerfStateCustom [customftPerfStates], overRideControl:=True

    If lPerfState.IsPerfState And overRideControl = False Then
        RaiseError ERR_PREVIOUS_PerfState_EXISTS, "A Control State is already set!"
    End If
    If overRideControl Then lPerfState.IsPerfState = False
    If lPerfState.IsPerfState Then
        RaiseError ERR_CONTROL_STATE, errorDesc:="Cannot Overwrite Control State with another Control state.  Previous PerfState Must be removed first."
    ElseIf cstmState.IsDefault Then
        RaiseError ERR_CONTROL_STATE, errorDesc:="Cannot Set Control State to be Default State. Default State is achieved by Clearing The State ('pbPerf.DefaultMode')"
    End If
        
    SetState cstmState
    
    

End Function

Private Function CheckState()
    'CheckState ** always ** implies the Fin Tool is doing something and user should not be interacting
    If lPerfState.IsPerfState Then
'        If IsDEV And DebugMode Then
'            'make a note of which properties have been fiddled with somewhere
'            If Not Application.EnableEvents = lPerfState.Events Then DebugPrint "*CheckState, EVENTS did not match*"
'            If Not Application.Interactive = lPerfState.Interactive Then DebugPrint "*CheckState, INTERACTIVE did not match*"
'            If Not Application.ScreenUpdating = lPerfState.Screen Then DebugPrint "*CheckState, SCREEN UPDATING did not match*"
'            If Not Application.Calculation = lPerfState.calc Then DebugPrint "*CheckState, CALCULATION did not match*"
'        End If
        SetState lPerfState
    Else
        SuspendState
    End If
End Function

Private Function PerfStateClear(Optional doNotProtect As Boolean = False, _
    Optional doNotDumpTrace As Boolean = False, _
    Optional forceSheet As Boolean = False, _
    Optional byPassCloseChk As Boolean = False)
    On Error GoTo E:
    
'   ~~~ ~~~ THIS FUNCTION CLEARS THE EXISTING 'SUSPEND' CONTROL
'                   AND THEN APPLIES THE 'DEFAULT (User Mode) ' PERFORMANCE SETTINGS
    '   NOTE, THIS METHOD IS NOT INTENDED TO NAVIGATE SHEETS.
    '   IT IS INTENDED TO ** CLOSE ** any sheets listed in l_HiddenSheets, IF they are visible and the App is not Closing
    '   IT IS INTENDED TO NAVIGATE TO ** homeSheet in the event that **:
    '    - This Function is running after an Error Has been raised, and Sheet desination may no longer be known
    '    - For an abnormal reason, there is no valid visble Sheet to navigate to
    
    Dim failed As Boolean, mustHide As Boolean
    'normally, actSht should be the last sheet that was navigated to, and will be the sheet the user see when the Control State is cleared
    mustHide = MustHideSheets
    CheckState
    
'   ~~~ ~~~ ~~~ ~~~    GENERAL ACTIONS   ~~~ ~~~ ~~~ ~~~
        If Not doNotDumpTrace Then
            If TraceQueueCount > 100 Then DumpTraceIfAvail
        End If
        If ThisWorkbook.Windows(1).DisplayWorkbookTabs = False Then
                ThisWorkbook.Windows(1).DisplayWorkbookTabs = True
        End If
    
    
''   ~~~ ~~~ ~~~ ~~~   TODO - POSSIBLE RAISE ERROR   ~~~ ~~~ ~~~ ~~~
'        ''TODO: NEED TO ABSTRACT PackageMgr. OUT OF pbCommon Modules
'        If PackageMgr.PackageIsExecuting Then
'            '   WE SHOULDN'T EVER BE CALLING THIS IF A PACKAGE IS RUNNING
'            '   TODO: FIGURE OUT HOW TO DO SIMILAR CHECK IF AN EVENT TRIGGERED PROCESS IS RUNNING (E.G. DOUBLE CLICK)
'            If IsDEV Then
'                Beep
'                DebugPrint "PerfStateClear called while package is executing. SHOULD PROBABLY MAKE THIS RAISE ERROR"
'                'Stop
'            End If
'        End If
'
'
    
    '   APP IS ** NOT ** CLOSING
    If Not ftState = ftClosing Then
        '   MAKE SURE WE HAVE A VALID ACTIVE SHEET
        If ThisWorkbook.ActiveSheet Is Nothing Then
            If Not l_homeSheet Is Nothing Then
                If Not HomeSheet.visible = xlSheetVisible Then HomeSheet.visible = xlSheetVisible
                HomeSheet.Activate
            End If
        End If
        If mustHide Then
            HideSheets
        End If
    End If
    
    If Not doNotProtect Then pbProtectSheet ThisWorkbook.ActiveSheet
    
    '   APP ** IS ** CLOSING
    If ftState = ftClosing Then
        If byPassCloseChk Then byPassOnCloseCheck = True
    End If
    
Finalize:
    On Error Resume Next
    
    If forceSheet And Not l_homeSheet Is Nothing Then
        If Not l_homeSheet.visible = xlSheetVisible Then
            l_homeSheet.visible = xlSheetVisible
        End If
        If Not l_homeSheet Is ThisWorkbook.ActiveSheet Then
            l_homeSheet.Activate
        End If
    End If
    
    pbPackageRunning = False
    lPerfState = DefaultState
    If failed Then
        RestoreDefaultAppSettingsOnly
    Else
        SetState lPerfState
    End If
    
   Exit Function
E:
   failed = True
   Debug.Print "PerfStateClear Error: ", Err.Number, Err.Description
   Beep
   Err.Clear
   Resume Finalize:
   'ErrorCheck options:=ftERR_NoBeeper
    
End Function

Public Function CurrentAppliedftPerfStates() As ftPerfStates
      CurrentAppliedftPerfStates = lPerfState
End Function

Public Function CurrentftPerfStates() As ftPerfStates

'   Get Current 'UI' State Settings
'   Informational -- this does not change any settings
    Dim retV As ftPerfStates
    retV.alerts = Application.DisplayAlerts
    retV.calc = Application.Calculation
    retV.Cursor = Application.Cursor
    retV.Events = Application.EnableEvents
    retV.Interactive = Application.Interactive
    retV.Screen = Application.ScreenUpdating
    retV.IsPerfState = False
    CurrentftPerfStates = retV

End Function

Private Property Get DefaultState() As ftPerfStates
'   THIS PROPERTY PROVIDES THE 'ftPerfStates' Values for
'   What is considered the 'Default' operating mode for a user.
'   Do not ever call this direct, as it won't do anything.
'   When your code is one you sould call: pbPerf.DefaultMode -- plus any additional valid enum items

    Dim retV As ftPerfStates
    
    retV.IsDefault = True
    retV.IsPerfState = False
    
    retV.alerts = True
    retV.calc = xlCalculationAutomatic
    retV.Cursor = XlMousePointer.xlDefault
    retV.Events = True
    retV.Interactive = True
    retV.Screen = True
    DefaultState = retV
    
End Property

Private Function SetState(updState As ftPerfStates)
    With Excel.Application
        If Not .Interactive = updState.Interactive Then .Interactive = updState.Interactive
        If Not .ScreenUpdating = updState.Screen Then .ScreenUpdating = updState.Screen
        If Not .Cursor = updState.Cursor Then .Cursor = updState.Cursor
        If Not .Calculation = updState.calc Then .Calculation = updState.calc
        If Not .DisplayAlerts = updState.alerts Then .DisplayAlerts = updState.alerts
        If Not .EnableEvents = updState.Events Then .EnableEvents = updState.Events
        
        If updState.IsPerfState Then
        ' ~~~ SET AS CONTROL STATE ~~~
            .EnableAnimations = False
            .PrintCommunication = False
            .EnableMacroAnimations = False
            ''ButtonPause = True
        Else
        ' ~~~ SET DEFAULT (CLEAR) STATE ~~~
            .EnableAnimations = True
            .PrintCommunication = True
            .EnableMacroAnimations = True
            ''ButtonPause = False
        End If
    End With
    
    lPerfState = updState
End Function
Private Function SuspendState( _
    Optional scrnUpd As Boolean = False, _
    Optional scrnInter As Boolean = False, _
    Optional alerts As Boolean = False, _
    Optional calcMode As XlCalculation = xlCalculationAutomatic _
    )
    Dim susSt As ftPerfStates
    
    ' ~~~ SET AS CONTROL STATE ~~~
    susSt.IsPerfState = True
    susSt.IsDefault = False
    
    ' ~~~ SET CONFIGURABLE PROPERTIES ~~~
    susSt.alerts = alerts
    susSt.calc = calcMode
    susSt.Interactive = scrnInter
    susSt.Screen = scrnUpd
    
     ' ~~~ SET FORCED PROPERTIIES ~~~
    susSt.Events = False
    susSt.Cursor = xlWait
    susSt.IsDefault = False
    
    SetState susSt
End Function

Public Function RestoreDefaultAppSettingsOnly()
        
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.Interactive = True
        Application.Cursor = xlDefault
        Application.Calculation = xlCalculationAutomatic
        Application.EnableAnimations = True
        Application.EnableMacroAnimations = True

End Function

Private Sub Class_Initialize()
    'If you have the pbError Module from https://github.com/lopperman/VBA-pbUtil
    ' Then you can change the const used for the error to: ERR_CLASS_INSTANCE_PROHIBITED
    If Not Me Is pbPerf Then RaiseError PRIVATE_ERR_CLASS_INSTANCE_PROHIBITED, errorDesc:="pbPerf may not be created as a new Class Instance outside it's PredeclaredId 'pbPerf'"
    Configure
End Sub

Private Property Get HomeSheet() As Worksheet
    If l_homeSheet Is Nothing And Len(FORCED_WORKSHEET_NAME) > 0 Then
        If WorksheetExists(FORCED_WORKSHEET_NAME, ThisWorkbook) Then
            Set l_homeSheet = ThisWorkbook.Worksheets(FORCED_WORKSHEET_NAME)
        End If
    End If
End Property

Private Property Get MustHideSheets() As Boolean
    Dim hideIDX As Long
    If l_hideSheets.Count = 0 Then
        MustHideSheets = False
        Exit Property
    End If
    For hideIDX = 1 To l_hideSheets.Count
        If ThisWorkbook.Worksheets(l_hideSheets.KeyByIndex(hideIDX)).visible = xlSheetVisible Then
                MustHideSheets = True
                Exit For
        End If
    Next hideIDX
End Property
Private Property Get MustActivateHomeSheet() As Boolean
    Dim hideIDX As Long
    If Len(FORCED_WORKSHEET_NAME) = 0 Then
        MustActivateHomeSheet = False
        Exit Property
    End If
    If l_hideSheets.Count = 0 Then
        MustActivateHomeSheet = False
        Exit Property
    End If
    For hideIDX = 1 To l_hideSheets.Count
        If ThisWorkbook.Worksheets(l_hideSheets.KeyByIndex(hideIDX)) Is ThisWorkbook.ActiveSheet Then
            MustActivateHomeSheet = True
            Exit For
        End If
    Next hideIDX
End Property
Private Function HideSheets()
    If MustActivateHomeSheet Then
        If Not HomeSheet.visible = xlSheetVisible Then HomeSheet.visible = xlSheetVisible
        HomeSheet.Activate
    End If
        
    Dim hideIDX As Long
    If l_hideSheets.Count > 0 Then
        For hideIDX = 1 To l_hideSheets.Count
            If Not ThisWorkbook.Worksheets(l_hideSheets.KeyByIndex(hideIDX)).visible = _
                xlSheetVeryHidden Then
                ThisWorkbook.Worksheets(l_hideSheets.KeyByIndex(hideIDX)).visible = _
                    xlSheetVeryHidden
            End If
        Next hideIDX
    End If

End Function

