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
' Manage the PERFORMANCE STATE (EnableEvents, ScreenUpdating, UserInteraction, etc) of  your Workbook/App
'
'       ~~~ 'EASY MODE' ~~~ (Just Use LockdownOn, LockdownOff)
'
'       LockdownOn
'        - Use anywhere in the app at anytime to turn OFF or DISABLE
'               Events, ScreenUpdating,Interaction,DisplayAlerts,Animations,Etc

'       DefaultMode
'        - Use anywhere in the app at anytime to turn ON or ENABLE:
'               Will Always set Application.Calculation = xlCalculationAutomatic
'               Events, ScreenUpdating,Interaction,DisplayAlerts,Animations,Etc
'
' @module pbDebug
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1

Private lPerfState As ftPerfStates

' ________________________________________
'  ~~~~ ~~~~ ~~~ ~~~ EASY MODE ~~~ ~~~ ~~~ ~~~

    Public Function DefaultMode(Optional ignoreProtect As Boolean = False, _
        Optional ignoreDumpTrace As Boolean = False, _
        Optional forceSheet As Boolean = False, _
        Optional enableCloseBypass As Boolean = False)
        
        PerfStateClear doNotProtect:=ignoreProtect, doNotDumpTrace:=ignoreDumpTrace, forceToDashboard:=forceSheet, byPassCloseChk:=enableCloseBypass
    
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
    
        If PerfEnumCompare(options, ftPerfOption.poOverride) Then
            ''TODO: ERR ON INVALID ADDITIONAL ENUMS
            lPerfState.IsPerfState = False
        End If
        
        '   ~~~ ~~~ CLEARING CONTROL ~~ ~~~
        If PerfEnumCompare(options, ftPerfOption.poClearControl) Then
                Dim tmpDisableProt As Boolean, tmpBypassCloseChk As Boolean, tmpDoNotDump As Boolean, tmpForceDashboard As Boolean
                'Default calc to Automatic
                If PerfEnumCompare(options, ftPerfOption.poIgnoreSheetProtect) Then tmpDisableProt = True
                If PerfEnumCompare(options, ftPerfOption.poKeepTraceQueued) Then tmpDoNotDump = True
                If PerfEnumCompare(options, ftPerfOption.poBypassCloseChecks) Then tmpBypassCloseChk = True
                If PerfEnumCompare(options, ftPerfOption.poForceFinalSheet) Then tmpForceDashboard = True
        
                PerfStateClear doNotProtect:=tmpDisableProt, doNotDumpTrace:=tmpDoNotDump, forceToDashboard:=tmpForceDashboard, byPassCloseChk:=tmpBypassCloseChk
                
        '   ~~~ ~~~ ADDING CONTROL (SUSPEND) ~~ ~~~
        ElseIf PerfEnumCompare(options, ftPerfOption.poSuspendControl) Then
                
                Dim tmpScreen As Boolean, tmpInter As Boolean, tmpAlert As Boolean, tmpCalc As XlCalculation
                'Default calc to Automatic
                tmpCalc = xlCalculationAutomatic
                If PerfEnumCompare(options, ftPerfOption.poCalcModeManual) Then tmpCalc = xlCalculationManual
                If PerfEnumCompare(options, ftPerfOption.poDoNotDisable_Alerts) Then tmpAlert = True
                If PerfEnumCompare(options, ftPerfOption.poDoNotDisable_Screen) Then tmpScreen = True
                If PerfEnumCompare(options, ftPerfOption.poDoNotDisable_Interaction) Then tmpInter = True
                SuspendState scrnUpd:=tmpScreen, scrnInter:=tmpInter, alerts:=tmpAlert, calcMode:=tmpCalc
        
        ElseIf PerfEnumCompare(options, ftPerfOption.poCheckControl) Then
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
        If IsDEV And DebugMode Then
            'make a note of which properties have been fiddled with somewhere
            If Not Application.EnableEvents = lPerfState.Events Then DebugPrint "*CheckState, EVENTS did not match*"
            If Not Application.Interactive = lPerfState.Interactive Then DebugPrint "*CheckState, INTERACTIVE did not match*"
            If Not Application.ScreenUpdating = lPerfState.Screen Then DebugPrint "*CheckState, SCREEN UPDATING did not match*"
            If Not Application.Calculation = lPerfState.calc Then DebugPrint "*CheckState, CALCULATION did not match*"
        End If
        SetState lPerfState
    Else
        SuspendState
    End If
End Function

Private Function PerfStateClear(Optional doNotProtect As Boolean = False, _
    Optional doNotDumpTrace As Boolean = False, _
    Optional forceToDashboard As Boolean = False, _
    Optional byPassCloseChk As Boolean = False)
    On Error GoTo E:
'   ~~~ ~~~ THIS FUNCTION CLEARS THE EXISTING 'SUSPEND' CONTROL
'                   AND THEN APPLIES THE 'DEFAULT (User Mode) ' PERFORMANCE SETTINGS
    '   NOTE, THIS METHOD IS NOT INTENDED TO NAVIGATE SHEETS.
    '   IT IS INTENDED TO ** CLOSE ** wsBusy and wsOpenClose, IF they are visible and the App is not Closing
    '   IT IS INTENDED TO NAVIGATE TO ** wsDashBoard in the event that **:
    '    - This Function is runnning after an Error Has been raised, and Sheet desination may no longer be known
    '    - For an abnormal reason, there is no valid visble Sheet to navigate to
    
    Dim failed As Boolean, actSht As String
    'normally, actSht should be the last sheet that was navigated to, and will be the sheet the user see when the Control State is cleared
    actSht = ActiveSheetName
    CheckState
    
'   ~~~ ~~~ ~~~ ~~~   TODO - POSSIBLE RAISE ERROR   ~~~ ~~~ ~~~ ~~~
        ''TODO: NEED TO ABSTRACT PACKAGEMGR OUT OF pbCommon Modules
        If PackageMgr.PackageIsExecuting Then
            '   WE SHOULDN'T EVER BE CALLING THIS IF A PACKAGE IS RUNNING
            '   TODO: FIGURE OUT HOW TO DO SIMILAR CHECK IF AN EVENT TRIGGERED PROCESS IS RUNNING (E.G. DOUBLE CLICK)
            If IsDEV Then
                Beep
                Debug.Print "PerfStateClear called while package is executing. SHOULD PROBABLY MAKE THIS RAISE ERROR"
                'Stop
            End If
        End If

'   ~~~ ~~~ ~~~ ~~~    GENERAL ACTIONS   ~~~ ~~~ ~~~ ~~~
        If Not doNotDumpTrace Then DumpTraceIfAvail

    
    '   APP IS ** NOT ** CLOSING
    If Not ftState = ftClosing Then
        '   MAKE SURE WE HAVE A VALID ACTIVE SHEET
        If ThisWorkbook.activeSheet Is Nothing Then
            If Not wsDashboard.visible = xlSheetVisible Then wsDashboard.visible = xlSheetVisible
            wsDashboard.Activate
        End If
        If ThisWorkbook.activeSheet Is wsBusy Or ThisWorkbook.activeSheet Is wsOpenClose Then
            If Not wsDashboard.visible = xlSheetVisible Then wsDashboard.visible = xlSheetVisible
            wsDashboard.Activate
            If Not wsBusy.visible = xlSheetVeryHidden Then wsBusy.visible = xlSheetVeryHidden
            If Not wsOpenClose.visible = xlSheetVeryHidden Then wsOpenClose.visible = xlSheetVeryHidden
        End If
'        If Not WOMRunning Then
'            StartWOM
'            GoTo Finalize:
'        End If
    End If
    
    If Not doNotProtect Then ProtectSHT ThisWorkbook.activeSheet
    
    '   APP ** IS ** CLOSING
    If ftState = ftClosing Then
        If byPassCloseChk Then byPassOnCloseCheck = True
    End If
    
Finalize:
    On Error Resume Next
    
    If forceToDashboard Then
        If Not wsDashboard.visible = xlSheetVisible Then wsDashboard.visible = xlSheetVisible
        wsDashboard.Activate
    End If
    
    ButtonPause = False
    PackageRunning = False
    
    lPerfState = DefaultState
    SetState lPerfState
    
   Exit Function
E:
   failed = True
   ErrorCheck options:=ftERR_NoBeeper
    
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
            ButtonPause = True
        Else
        ' ~~~ SET DEFAULT (CLEAR) STATE ~~~
            .EnableAnimations = True
            .PrintCommunication = True
            .EnableMacroAnimations = True
            ButtonPause = False
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
        ButtonPause = False

End Function

'   ___ ___ ___ ___ ___ ___ ___ ___ ___ ___ ___ ___
'   ~~~~~~~~~~   FLAG ENUM COMPARE ~~~~~~~~~~'
    Private Function PerfEnumCompare(theEnum As Variant, enumMember As Variant) As Boolean
        'Use to check Bitwise enums
        PerfEnumCompare = CBool(theEnum And enumMember) = True
    End Function





