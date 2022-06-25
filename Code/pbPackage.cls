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
Option Explicit
Option Compare Text
Option Base 1

'   UNCOMMENT IF NOT USING pbError MODULE
'   Public Const ERR_CANNOT_CHANGE_PACKAGE_PROPERTY = vbObjectError + 1026
'   Public Const ERR_INVALID_PACKAGE_OPERATION = vbObjectError + 1027


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
Private l_bypassReqAction As Boolean
Private l_packageResult As Variant
Private l_returnResult As Boolean

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

'   Source Worksheet / Worksheet Active When Package Created
Public Property Get SourceSheet() As Worksheet
    If Not l_sourceSheet Is Nothing Then
        Set SourceSheet = l_sourceSheet
    End If
End Property
Public Property Set SourceSheet(srcSheet As Worksheet)
    If Not l_sourceSheet Is Nothing Then
        RaiseError ERR_CANNOT_CHANGE_PACKAGE_PROPERTY, errorDesc:="pbPackage.SourceSheet cannot be changed once set"
    End If
    Set l_sourceSheet = srcSheet
End Property
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

'   Target Worksheet (Optional)
'   If set, will make visible and active during 'AfterExecute' Function
Public Property Get TargetSheet() As Worksheet
    If Not l_targetSheet Is Nothing Then
        Set TargetSheet = l_targetSheet
    End If
End Property
Public Property Set TargetSheet(tarSheet As Worksheet)
    If Not l_targetSheet Is Nothing Then
        RaiseError ERR_CANNOT_CHANGE_PACKAGE_PROPERTY, errorDesc:="pbPackage.TargetSheet cannot be changed once set"
    End If
    Set l_targetSheet = tarSheet
End Property

'   Target
'   This is the property that is used to determine final navigation, if applicable
'   If TargetSheet or RedirectTarget is populated, then that worksheet will
'   be made visible and active.
Public Property Get Target() As Worksheet
    If Not l_RedirectTarget Is Nothing Then
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

Public Function Execute()
    l_failed = Not BeforeExecute
    l_failed = Not DoExecute
    AfterExecute
End Function
Private Function BeforeExecute()
On Error GoTo E:
    Dim failed As Boolean
    l_ExecuteStarted = Timer
    
    If IsManaged = False Then
        pbPerf.PerfStateCustom l_SysState, overRideControl:=True
        
    End If
    
Finalize:
    On Error Resume Next
     
    BeforeExecute = Not failed
     
    
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize
    
End Function
Private Function AfterExecute()
    l_ExecuteStarted = Timer
    
    If FailedPackage Then
            
    Else
    
    End If

End Function
Private Function DoExecute() As Boolean

End Function



'   Redirect Target Sheet
'   If the need arises to change the target location (Active Sheet After Package Executes)
'   You can set the 'Redirect Sheet' to a different worksheet, and optional invalidate the redirect
'   if the package ends with failure

Public Property Get CanRedirect() As Boolean
    'Can only redirect if the 'AfterExecute' has not started
    CanRedirect = l_ExecuteCompleted = 0
End Property

Public Function RedirectTarget(newTargetWS As Worksheet, Optional RedirectOnFail As Boolean = True)
    If Not CanRedirect Then
        RaiseError ERR_INVALID_PACKAGE_OPERATION, "pbPackage.REdirectTarget cannot set new target after Execution has finished"
    End If
    'changing this property is allowed as many times as needed
    Set l_RedirectTarget = newTargetWS
    l_RedirectOnFail = RedirectOnFail
End Function


Private Sub Class_Initialize()
    l_PackageCreated = Now
    l_SysState = DefaultPackageState
    l_RestoreDefaultState = True
    
    
End Sub
