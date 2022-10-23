Attribute VB_Name = "pbPerf"
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



Public Type pfLite
    Events As Boolean
    Interactive As Boolean
    Screen As Boolean
    alerts As Boolean
    Cursor As XlMousePointer
    calc As XlCalculation
End Type

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
        



Public Property Get PLite() As pfLite
'   Used For OneOffs Where need to ensure certain thing are off,
'   but when done, return to whatever the performance states were
    Dim retV As pfLite
    retV.alerts = Application.DisplayAlerts
    retV.calc = Application.Calculation
    retV.Cursor = Application.Cursor
    retV.Events = Application.EnableEvents
    retV.Interactive = Application.Interactive
    retV.Screen = Application.ScreenUpdating
    PLite = retV
End Property
Public Property Let PLite(pLt As pfLite)
    If Not Application.DisplayAlerts = pLt.alerts Then Application.DisplayAlerts = pLt.alerts
    If Not Application.Calculation = pLt.calc Then Application.Calculation = pLt.calc
    If Not Application.Cursor = pLt.Cursor Then Application.Cursor = pLt.Cursor
    If Not Application.EnableEvents = pLt.Events Then Application.EnableEvents = pLt.Events
    If Not Application.Interactive = pLt.Interactive Then Application.Interactive = pLt.Interactive
    If Not Application.ScreenUpdating = pLt.Screen Then Application.ScreenUpdating = pLt.Screen
End Property


' ________________________________________
'  ~~~~ ~~~~ ~~~ ~~~ EASY MODE ~~~ ~~~ ~~~ ~~~

Public Function DefaultMode(Optional ByVal ignoreProtect As Boolean = False, _
    Optional ByVal ignoreDumpTrace As Boolean = False, _
    Optional ByVal forceSheet As Boolean = False, _
    Optional ByVal enableCloseBypass As Boolean = False)
    
    PerfStateClear doNotProtect:=ignoreProtect, doNotDumpTrace:=ignoreDumpTrace, forceSheet:=forceSheet, byPassCloseChk:=enableCloseBypass
    Application.StatusBar = False
    
End Function

Public Function SuspendMode(Optional ByVal calc As XlCalculation = XlCalculation.xlCalculationAutomatic)
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
Public Function PerfState(ByVal options As ftPerfOption) As Boolean
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

Public Function PerfStateCustom(cstmState As ftPerfStates, Optional ByVal overRideControl As Boolean = False)
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
        SetState lPerfState
    Else
        SuspendState
    End If
End Function

Private Function PerfStateClear(Optional ByVal doNotProtect As Boolean = False, _
    Optional ByVal doNotDumpTrace As Boolean = False, _
    Optional ByVal forceSheet As Boolean = False, _
    Optional ByVal byPassCloseChk As Boolean = False)
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
    Application.EnableEvents = False
    
'   ~~~ ~~~ ~~~ ~~~    GENERAL ACTIONS   ~~~ ~~~ ~~~ ~~~
        If ThisWorkbook.Windows(1).DisplayWorkbookTabs = False Then
                ThisWorkbook.Windows(1).DisplayWorkbookTabs = True
        End If
    
'   ~~~ ~~~ ~~~ ~~~    APP IS ** NOT ** CLOSING   ~~~ ~~~ ~~~ ~~~
    If Not ftState = ftClosing Then
        '   MAKE SURE WE HAVE A VALID ACTIVE SHEET
        If ThisWorkbook.ActiveSheet Is Nothing Then
            If Not wsDashboard Is Nothing Then
                If Not wsDashboard.visible = xlSheetVisible Then wsDashboard.visible = xlSheetVisible
                wsDashboard.Activate
            End If
        End If
        If mustHide Then
            HideSheets
        End If
    End If
    
    If Not doNotProtect Then ProtectSht ThisWorkbook.ActiveSheet
    
'   ~~~ ~~~ ~~~ ~~~    APP ** IS ** CLOSING   ~~~ ~~~ ~~~ ~~~
    If ftState = ftClosing Then
        If byPassCloseChk Then byPassOnCloseCheck = True
    End If
    
Finalize:
    On Error Resume Next
    
    If forceSheet And Not wsDashboard Is Nothing Then
        If Not wsDashboard.visible = xlSheetVisible Then
            wsDashboard.visible = xlSheetVisible
        End If
        If Not wsDashboard Is ThisWorkbook.ActiveSheet Then
            wsDashboard.Activate
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
   LogError ConcatWithDelim(", ", "PerfStateClear Error: ", Err.number, Err.Description)
   
   Err.Clear
   Resume Finalize:
    
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
            '.PrintCommunication = True
            .EnableMacroAnimations = True
            ''ButtonPause = False
        End If
    End With
    
    lPerfState = updState
End Function
Private Function SuspendState( _
    Optional ByVal scrnUpd As Boolean = False, _
    Optional ByVal scrnInter As Boolean = False, _
    Optional ByVal alerts As Boolean = False, _
    Optional ByVal calcMode As XlCalculation = xlCalculationAutomatic _
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
        
    With Application
        .Interactive = True
        .Cursor = xlDefault
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .EnableAnimations = True
        .EnableMacroAnimations = True
        .EnableEvents = True
    End With
    Application.StatusBar = False
    lPerfState.IsPerfState = False
End Function

Private Property Get MustHideSheets() As Boolean
    If wsBusy.visible = xlSheetVisible Then MustHideSheets = True
    If wsOpenClose.visible = xlSheetVisible Then MustHideSheets = True
End Property
Private Property Get MustActivateHomeSheet() As Boolean
    Dim hideIDX As Long
    If Not ThisWorkbook.ActiveSheet Is Nothing Then
        If ThisWorkbook.ActiveSheet Is wsBusy Or ThisWorkbook.ActiveSheet Is wsOpenClose Then
            MustActivateHomeSheet = True
        End If
    End If
End Property
Private Function HideSheets()

    If MustActivateHomeSheet Then
        If Not wsDashboard.visible = xlSheetVisible Then wsDashboard.visible = xlSheetVisible
        wsDashboard.Activate
    End If
        
    If wsBusy.visible = xlSheetVisible Then wsBusy.visible = xlSheetVeryHidden
    If wsOpenClose.visible = xlSheetVisible Then wsOpenClose.visible = xlSheetVeryHidden
    


End Function



