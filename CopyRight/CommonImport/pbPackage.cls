VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' pbPackage v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' A locked and loaded sequence of events that
' can be passed around and fired when ready
'
' @module pbPackage
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'   *** Requires the following additional Modules ***
'
'   - MODULES: pbCommon, pbDebug, pbError, pbMiscUtil, pbRangeArray,
'       pbRange, pbListobj
'
'   - PRIVATE MODULES:  pbPerf
'
'   - See Class Initialize Event for Default Settings
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit
Option Compare Text
Option Base 1

'   UNCOMMENT IF NOT USING pbError MODULE
'   Public Const ERR_CANNOT_CHANGE_PACKAGE_PROPERTY = vbObjectError + 1026
'   Public Const ERR_INVALID_PACKAGE_OPERATION = vbObjectError + 1027

'    Set this to Public Function/Sub to be called before any package is run.
'    Set BypassRequiredAction to FALSE to disabled that method call for package
'
'   If Set to True, when RunRequiredAction is called, it will execute the code in the 'RunRequiredAction'
'       function in this class
'   If 'RunRequiredAction' is successful, the package will execute afterwards
'   If 'RunRequiredAction' fails, this will cause the package
'       to stop running (will not execute)
Private Const CanRunReqAction As Boolean = True


'   IDENTIFICATION
Private l_ActionType As InitActionEnum
Private l_taskName As String
Private l_PackageCreated As Date
Private l_ExecuteStarted As Double
Private l_ExecuteCompleted As Double
Private l_IsManaged As Boolean

'   SYSTEM STATES
Private l_SysState As ftPerfStates
Private l_RestoreDefaultState As Boolean

'   EXECUTION
Private l_failed As Boolean
Private l_sourceSheet As Worksheet
Private l_targetSheet As Worksheet
Private l_UnprotectSheets() As Variant
Private l_RedirectTarget As Worksheet
Private l_RedirectOnFail As Boolean
Private l_OnErrorRedirect As Worksheet
Private l_bypassReqAction As Boolean
Private l_packageResult As Variant
Private l_returnResult As Boolean
Private l_protectTarget As Boolean
Private l_protectRedirect As Boolean
Private l_HideSource As Boolean
Private l_DoNotCleanUp As Boolean

'   FUNCTION CALLING
Private l_CallObject As Object
Private l_CallObjectName As String
Private l_CallProcName As String
Private l_CallType As VbCallType
Private l_CallArgs As Collection


'   ~~~ PACKAGE CONFIGURATION ~~~

'   ACTION TYPE
'   Describes the 'initiating event type' in which this
'   package is being created (ButtonClick, App, Workbook, Worksheet Event, etc)
Public Property Get ActionType() As InitActionEnum
    ActionType = l_ActionType
End Property
Public Property Let ActionType(InitAction As InitActionEnum)
    If l_ActionType > 0 Then
        RaiseError ERR_CANNOT_CHANGE_PACKAGE_PROPERTY, errorDesc:="pbPackage Property 'ActionType' cannot be changed after being set"
    End If
    l_ActionType = InitAction
End Property

'   Bypass Required Actions
'   Used to enable an overall process for 'checking things'
'   If you have a process to check the state of certain data/objects (and perform actions if needed), then
'   set the function/sub name in the 'CheckRequiredActions' constant, and that process will
'   run before any package is executed.  To prevent a particular package from running that process,
'   Set 'BypassRequiredAction' to False
Public Property Get BypassRequiredAction() As Boolean
    BypassRequiredAction = l_bypassReqAction
End Property
Public Property Let BypassRequiredAction(byPassReqActn As Boolean)
    l_bypassReqAction = byPassReqActn
End Property

'   CanHideSource
'   If SourceWorksheet has a value, AND Target has a value
'   then after package is executed, and Target sheet is active,
'   Then SourcSheet will become VeryHidden
'   Default = True
Public Property Get CanHideSource() As Boolean
    CanHideSource = l_HideSource
End Property
Public Property Let CanHideSource(vl As Boolean)
    l_HideSource = vl
End Property

Public Property Get ReturnResult() As Boolean
    ReturnResult = l_returnResult
End Property
Public Property Let ReturnResult(rtRslt As Boolean)
    l_returnResult = rtRslt
End Property
Public Property Get PackageResult() As Variant
    PackageResult = l_packageResult
End Property
Public Property Get IgnoreCleanup() As Boolean
    IgnoreCleanup = l_DoNotCleanUp
End Property
Public Property Let IgnoreCleanup(ignore As Boolean)
    l_DoNotCleanUp = ignore
End Property

'   Source Worksheet / Worksheet Active When Package Created
Public Property Get SourceSheet() As Worksheet
    If Not l_sourceSheet Is Nothing Then
        Set SourceSheet = l_sourceSheet
    End If
End Property
Public Function SetSourceSheet(srcSheet As Worksheet)
    If Not l_sourceSheet Is Nothing Then
        RaiseError ERR_CANNOT_CHANGE_PACKAGE_PROPERTY, errorDesc:="pbPackage.SourceSheet cannot be changed once set"
    End If
    Set l_sourceSheet = srcSheet
End Function
Public Function UnprotectSheets(ParamArray unprotSheets() As Variant)
    If UBound(unprotSheets) <> -1 Then
    
        If StrComp(TypeName(unprotSheets(LBound(unprotSheets))), "Worksheet", vbTextCompare) <> 0 Then
            RaiseError ERR_INVALID_PACKAGE_OPERATION, errorDesc:="pbPackage.UnprotectSheets Invalid Parameter: Must Be Type: Worksheet"
        End If
        
        Dim newArr As Variant
        newArr = ArrParams(unprotSheets)
        If ArrayInfo(newArr).Dimensions > 0 Then
            l_UnprotectSheets = newArr
            Erase newArr
        End If
    End If
End Function

Public Function OnErrorRedirect(ErrorSheet As Worksheet)
    Set l_OnErrorRedirect = ErrorSheet
End Function
Public Property Get ErrorSheet() As Worksheet
    If Not l_OnErrorRedirect Is Nothing Then
        Set ErrorSheet = l_OnErrorRedirect
    End If
End Property

'   Target Worksheet (Optional)
'   If set, will make visible and active during 'AfterExecute' Function
Public Property Get TargetSheet() As Worksheet
    If Not l_targetSheet Is Nothing Then
        Set TargetSheet = l_targetSheet
    End If
End Property
Public Function SetTargetSheet(tarSheet As Worksheet)
    If Not l_targetSheet Is Nothing Then
        RaiseError ERR_CANNOT_CHANGE_PACKAGE_PROPERTY, errorDesc:="pbPackage.TargetSheet cannot be changed once set"
    End If
    Set l_targetSheet = tarSheet
End Function
Public Property Let ProtectTargetSheet(protTarget As Boolean)
    l_protectTarget = protTarget
End Property
Public Property Get ProtectTargetSheet() As Boolean
    ProtectTargetSheet = l_protectTarget
End Property

'   Target
'   This is the property that is used to determine final navigation, if applicable
'   If TargetSheet or RedirectTarget is populated, then that worksheet will
'   be made visible and active.
Private Property Get Target() As Worksheet
    If FailedPackage And Not ErrorSheet Is Nothing Then
        Set Target = ErrorSheet
    ElseIf Not l_RedirectTarget Is Nothing Then
        If (FailedPackage And l_RedirectOnFail) Or Not FailedPackage Then
            Set Target = l_RedirectTarget
        End If
    End If
    If Target Is Nothing And Not TargetSheet Is Nothing Then
        Set Target = TargetSheet
    End If
End Property

'   TASK NAME
'   A friendly short desc of what the package is.
'   If the Package is being constructed due to click a Button, then
'   Task Name should be the name of the Button.
Public Property Get TaskName() As String
    TaskName = l_taskName
End Property
Public Property Let TaskName(tskName As String)
    If Len(l_taskName) > 0 And StrComp(l_taskName, tskName, vbTextCompare) <> 0 Then
        RaiseError ERR_CANNOT_CHANGE_PACKAGE_PROPERTY, errorDesc:="pbPackage Property 'TaskName' cannot be changed after being set"
    End If
    l_taskName = tskName
End Property
Public Property Get ButtonName() As String
    If l_ActionType = iaButtonClick Then ButtonName = l_taskName
End Property

'   PACKAGE INFORMATION
'   (Set in Class_Initialize)
Public Property Get PackageCreated() As Date
    PackageCreated = l_PackageCreated
End Property
Public Property Get RunTimeSeconds()
    RunTimeSeconds = Math.Round(l_ExecuteCompleted - l_ExecuteStarted, 3)
End Property

'   Is Package Managed
'   This should only be set by a PackageManager
'   that queues and runs packages. This enables a manager to control
'   A managed package will do little or no 'setup' or 'teardown' activities
Public Property Get IsManaged() As Boolean
    IsManaged = l_IsManaged
End Property
'   For my Reddit Friends, You'll have to wait a bit before I provide my Package Manager :-)
Public Function SetManager(pkgMgr As Object)
    If Not pkgMgr Is Nothing Then
        If TypeName(pkgMgr) = "PackageMgr" Then l_IsManaged = True
    End If
End Function

'   ~~~ MANAGE SYSTEM STATES ~~~
Public Property Get RestoreDefaultState() As Boolean
    RestoreDefaultState = l_RestoreDefaultState
End Property
Public Property Let RestoreDefaultState(restDefault As Boolean)
   If l_IsManaged Then
        RaiseError ERR_CANNOT_CHANGE_PACKAGE_PROPERTY, errorDesc:="pbPackage.RestoreDefaultState cannot be set for managed packages"
   End If
    l_RestoreDefaultState = restDefault
End Property

'   Default Package State.  Unless changed, the Application
'   will set itself to this state while running
Private Property Get DefaultPackageState() As ftPerfStates
    Dim defState As ftPerfStates
    defState.alerts = False
    defState.calc = xlCalculationManual
    defState.Cursor = xlWait
    defState.Events = False
    defState.Interactive = False
    defState.IsDefault = False
    defState.IsPerfState = True
    defState.Screen = False
    DefaultPackageState = defState
End Property
Public Function ModifyExecutionState(adjust As ModifySuspendState)

    If EnumCompare(adjust, ModifySuspendState.CalculationAuto + ModifySuspendState.CalculationManual, ecAnd) Then
        RaiseError ERR_CANNOT_CHANGE_PACKAGE_PROPERTY, errorDesc:="pbPackage.ModifyExecutionState: Cannot Set Calcuation to Manual AND Auto"
    End If
    If EnumCompare(adjust, ModifySuspendState.EnableEvents) Then l_SysState.Events = True
    If EnumCompare(adjust, ModifySuspendState.EnableAlerts) Then l_SysState.alerts = True
    If EnumCompare(adjust, ModifySuspendState.EnableInteractive) Then l_SysState.Interactive = True
    If EnumCompare(adjust, ModifySuspendState.EnableScreenUpdate) Then l_SysState.Screen = True
    If EnumCompare(adjust, ModifySuspendState.CalculationAuto) Then l_SysState.calc = xlCalculationAutomatic
    If EnumCompare(adjust, ModifySuspendState.CalculationManual) Then l_SysState.calc = xlCalculationManual
 
 End Function

'   ~~~ EXECUTION ~~~
'   Failed Property is set after 'DoExecute' Function
Public Property Get FailedPackage() As Boolean
    FailedPackage = l_failed
End Property

' If package has 'Call By' set up, AND 'ReturnResult' = True, then
' the result of the Call By can be obtined from the 'PackageResult' Property
Public Function Execute()
    l_failed = Not BeforeExecute
    pbPackageRunning = True
    If Not l_failed Then
        l_failed = Not DoExecute
    End If
    AfterExecute
    pbPackageRunning = False
End Function
Private Function BeforeExecute()
On Error GoTo E:
    Dim failed  As Boolean
    Dim idx As Long
    
    Dim unpAI  As ArrInformation
    Dim tmpWS As Worksheet
    
    l_ExecuteStarted = Timer
    
    If IsManaged = False Then
        pbPerf.PerfStateCustom l_SysState, overRideControl:=True
    End If
    If Not BypassRequiredAction Then
        failed = Not RunRequiredAction
    End If
    If IsManaged = False Then
        pbPerf.PerfStateCustom l_SysState, overRideControl:=True
    End If
    
    
    unpAI = ArrayInfo(l_UnprotectSheets)
    If unpAI.Dimensions > 0 Then
        For idx = unpAI.LBound_first To unpAI.Ubound_first
            Set tmpWS = l_UnprotectSheets(idx, 1)
            If tmpWS.ProtectContents Then
                pbUnprotectSheet tmpWS
            End If
        Next idx
    End If
    
Finalize:
    On Error Resume Next
     
    Set tmpWS = Nothing
    BeforeExecute = Not failed
    
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize
    
End Function
Private Function AfterExecute()
    l_ExecuteStarted = Timer
    
    If Not Target Is Nothing Then
        If Not Target.visible = xlSheetVisible Then Target.visible = xlSheetVisible
        If Not Target Is ThisWorkbook.ActiveSheet Then Target.Activate
        If Not SourceSheet Is Nothing Then
            If CanHideSource Then SourceSheet.visible = xlSheetVeryHidden
        End If
        If ProtectTargetSheet Then pbProtectSheet Target
    End If

    If IsManaged = False And IgnoreCleanup = False Then
        pbPerf.DefaultMode
    End If
    
    pbPackageRunning = False

End Function
Private Function DoExecute() As Boolean
On Error GoTo E:
    Dim failed As Boolean
    
    pbPackageRunning = True

    If CallBySet Then
        ExecuteCallByName
    End If

Finalize:
    On Error Resume Next
    
    DoExecute = Not failed
    
    If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:
End Function

Private Function RunRequiredAction() As Boolean
On Error GoTo E:
    Dim failed As Boolean
    If CanRunReqAction = False Then
        RunRequiredAction = True
        Exit Function
    End If
    
'   ~~~ INSERT CUSTOM CODE HERE ~~~
'   Should Return True if success, false is fail
'   If fail, package will not be run
    
    If ValidProjectConfig And Setting2(seMustCheckReqActions) Then
        DebugPrint "Starting RunRequiredAction for " & Me.TaskName
        ShowBusy rightSideMessage:="Checking for required actions"
        Wom.PerformRequiredAction
    End If

Finalize:
    On Error Resume Next
    RunRequiredAction = Not failed
    If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:

End Function

'   Redirect Target Sheet
'   If the need arises to change the target location (Active Sheet After Package Executes)
'   You can set the 'Redirect Sheet' to a different worksheet, and optional invalidate the redirect
'   if the package ends with failure

Public Property Get CanRedirect() As Boolean
    'Can only redirect if the 'AfterExecute' has not started
    CanRedirect = l_ExecuteCompleted = 0
End Property

Public Function RedirectTarget(newTargetWS As Worksheet, Optional RedirectOnFail As Boolean = True, Optional protectWkSht As Boolean = False)
    If Not CanRedirect Then
        RaiseError ERR_INVALID_PACKAGE_OPERATION, "pbPackage.REdirectTarget cannot set new target after Execution has finished"
    End If
    'changing this property is allowed as many times as needed
    Set l_RedirectTarget = newTargetWS
    l_RedirectOnFail = RedirectOnFail
    l_protectRedirect = protectWkSht
End Function
Public Property Get ProtectRedirect() As Boolean
    ProtectRedirect = l_protectRedirect
End Property

' ~~~ 'CALL BY' ~~~
Public Property Get CallBySet() As Boolean
    CallBySet = Len(l_CallProcName) > 0
End Property

Public Function SetCallByNamePublicMethod(procName As String)
    Set l_CallObject = Nothing
    l_CallObjectName = "[PUBLIC]"
    l_CallProcName = procName
    l_CallType = VbMethod
End Function

Public Function SetCallByName(obj As Object, procName As String, callType As VbCallType)
    If Not obj Is Nothing Then
        Set l_CallObject = obj
        l_CallObjectName = TypeName(obj)
        l_CallProcName = procName
        l_CallType = callType
    End If
End Function

Public Function UpdateArgs(ParamArray args() As Variant)
    If l_CallArgs Is Nothing Then
        Set l_CallArgs = New Collection
    End If
    Dim aIDX As Long
    For aIDX = LBound(args) To UBound(args)
        l_CallArgs.add args(aIDX)
    Next aIDX
End Function
Public Property Get ArgsCount() As Long
    If Not l_CallArgs Is Nothing Then
        ArgsCount = l_CallArgs.Count
    Else
        ArgsCount = 0
    End If
End Property

Public Function ExecuteCallByName() As Boolean
On Error GoTo E:
    Dim failed As Boolean

    If Len(l_CallProcName & vbNullString) = 0 Then GoTo Finalize:

    If l_CallObject Is Nothing Then
        Trace "* Executing Public Method: (pkg name: " & l_taskName & ") - " & l_CallProcName
        Application.Run l_CallProcName
        GoTo Finalize:
    End If
    
    Trace "* Executing Package: (name: " & l_taskName & ") - " & TypeName(l_CallObject) & "." & l_CallProcName
    
    Select Case ArgsCount
        Case 0
            If ReturnResult = True Then
                l_packageResult = CallByName(l_CallObject, l_CallProcName, l_CallType)
            Else
                CallByName l_CallObject, l_CallProcName, l_CallType
            End If
        Case 1
            If ReturnResult = True Then
                l_packageResult = CallByName(l_CallObject, l_CallProcName, l_CallType, l_CallArgs(1))
            Else
                CallByName l_CallObject, l_CallProcName, l_CallType, l_CallArgs(1)
            End If
        Case 2
            If ReturnResult = True Then
                l_packageResult = CallByName(l_CallObject, l_CallProcName, l_CallType, l_CallArgs(1), l_CallArgs(2))
            Else
                CallByName l_CallObject, l_CallProcName, l_CallType, l_CallArgs(1), l_CallArgs(2)
            End If
        Case 3
            If ReturnResult = True Then
                l_packageResult = CallByName(l_CallObject, l_CallProcName, l_CallType, l_CallArgs(1), l_CallArgs(2), l_CallArgs(3))
            Else
                CallByName l_CallObject, l_CallProcName, l_CallType, l_CallArgs(1), l_CallArgs(2), l_CallArgs(3)
            End If
        Case 4
            If ReturnResult = True Then
                l_packageResult = CallByName(l_CallObject, l_CallProcName, l_CallType, l_CallArgs(1), l_CallArgs(2), l_CallArgs(3), l_CallArgs(4))
            Else
                CallByName l_CallObject, l_CallProcName, l_CallType, l_CallArgs(1), l_CallArgs(2), l_CallArgs(3), l_CallArgs(4)
            End If
        Case 5
            If ReturnResult = True Then
                l_packageResult = CallByName(l_CallObject, l_CallProcName, l_CallType, l_CallArgs(1), l_CallArgs(2), l_CallArgs(3), l_CallArgs(4), l_CallArgs(5))
            Else
                CallByName l_CallObject, l_CallProcName, l_CallType, l_CallArgs(1), l_CallArgs(2), l_CallArgs(3), l_CallArgs(4), l_CallArgs(5)
            End If
        Case 6
            If ReturnResult = True Then
                l_packageResult = CallByName(l_CallObject, l_CallProcName, l_CallType, l_CallArgs(1), l_CallArgs(2), l_CallArgs(3), l_CallArgs(4), l_CallArgs(5), l_CallArgs(6))
            Else
                CallByName l_CallObject, l_CallProcName, l_CallType, l_CallArgs(1), l_CallArgs(2), l_CallArgs(3), l_CallArgs(4), l_CallArgs(5), l_CallArgs(6)
            End If
        Case Else
            failed = True
            Err.Raise 512 + 100, Source:=Me, Description:="Cannot Handle More than 6 Parameters in CallByName"
    End Select
    
    
Finalize:
    On Error Resume Next
    
    If failed Then l_failed = True
    ExecuteCallByName = Not failed
    If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:
End Function

Private Sub Class_Initialize()

    l_PackageCreated = Now
    l_SysState = DefaultPackageState
    l_RestoreDefaultState = True
    l_HideSource = True
    l_protectTarget = True
    
    
End Sub
