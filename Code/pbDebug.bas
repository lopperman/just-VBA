Attribute VB_Name = "pbDebug"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' pbDebug v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' General Utilities for Debugging (Print, Assert, Trace)
'
' @module pbDebug
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'   TO DO
'   - REFACTOR OUT WSBUSY FROM TRACE FUNCTION

Option Explicit
Option Compare Text
Option Base 1

'   used for wildcard compare against current username
Private Const DEV_USERNAME As String = "brower"

'   If this Const has a value, that macro/sub/function will be called when closing a trace
Private Const DUMP_TRACE_CALL_NAME As String = "DumpTraceQueue"

' override: IF conDebug = 0, DebugMode set to true when Override is true
' enables users to dump trace messages to debug screen
Private l_debugReleaseOverride As Boolean
Private l_lastTraceTimer As Single
Private l_lastTraceMsg As String
Private l_lastForcedTimer As Single
Private l_traceSession As Variant
Private lastEventTrace As Single
Private l_trcQueue As Collection

Public Property Get TraceQueue() As Collection
    Set TraceQueue = l_trcQueue
End Property
Public Function ClearTraceQueue()
    Set l_trcQueue = Nothing
End Function

Public Property Get debugReleaseOverride() As Boolean
    If l_debugReleaseOverride Then
        debugReleaseOverride = True
    Else
        debugReleaseOverride = False
    End If
    
End Property
Public Property Let debugReleaseOverride(vl As Boolean)
    l_debugReleaseOverride = vl
End Property

'Conditional Compiler Args
'conDebug=1
Public Property Get DebugMode() As Boolean
    #If conDebug Then
        DebugMode = True
    #Else
        DebugMode = False
    #End If
End Property
Public Property Get CanAssert() As Boolean
    #If conDebug Then
        CanAssert = True
    #Else
        CanAssert = False
    #End If
End Property

Public Property Get DebugInfo() As Boolean
    If debugReleaseOverride Then
        DebugInfo = True
    Else
        DebugInfo = False
    End If

End Property

Public Function ToggleDebugOverride()
    debugReleaseOverride = Not debugReleaseOverride
    
End Function

Public Function FriendlyDebugMode() As String
    Dim retV As String
    If DebugInfo Then
        retV = "TRACE: ON"
    Else
        retV = "TRACE: OFF"
    End If
    If DebugMode Then
        retV = "DEBUG: ON, " & retV
    End If
    FriendlyDebugMode = retV
End Function

Public Function Assert(condition As Boolean)
'   DEV NOTE:  Assert Requires Conditional Compiler Constant 'Debugging = 1'
'   If Assert Failes, Step through to go back to failed caller
    If Not CanAssert Then Exit Function
    Debug.Assert condition
    
End Function

Public Function DebugPrint(Optional stmnt As Variant)
    If IsMissing(stmnt) Then stmnt = ""
    If DebugInfo Or DebugMode Then
        Trace stmnt
    End If
End Function

Public Function ShutdownDebug()
   ClearTraceQueue

End Function

Public Property Get LastTraceMsg() As String
    LastTraceMsg = l_lastTraceMsg
End Property

Private Function QueueTrace(trcInfo As Variant)

    If l_trcQueue Is Nothing Then
        Set l_trcQueue = New Collection
    End If
    
    TraceQueue.add trcInfo
End Function

Public Property Get IsDEV() As Boolean
    IsDEV = LCase(Application.UserName) Like "*" & DEV_USERNAME & "*"
End Property

Public Function TraceSessionStart(Optional sessionName As String = vbNullString)
    l_lastTraceTimer = Timer
    l_traceSession = IIf(Not sessionName = vbNullString, sessionName, CStr("Trace-" & CStr(Timer)))
    Trace "Starting: " & l_traceSession, forceDoEvents:=True
End Function

Public Function DumpTraceIfAvail()
    If Len(DUMP_TRACE_CALL_NAME) > 0 Then
        Application.Run DUMP_TRACE_CALL_NAME
        DoEvents
    End If
End Function

Public Function TraceSessionEnd()
    If Len(l_traceSession & vbNullString) = 0 Then l_traceSession = "UNK Trace Session"
    Trace "Completed: " & l_traceSession & vbNullString, forceDoEvents:=True
    l_traceSession = vbNullString
    DumpTraceIfAvail
End Function

Public Function TrcSysInfo() As String
    TrcSysInfo = " ( E-" & IIf(Events, "*ON*", "off") & ", S-" & IIf(Application.ScreenUpdating, "*ON*", "off") & ", I-" & IIf(Application.Interactive, "*ON*", "off") & ", C-" & IIf(Application.Calculation = xlCalculationAutomatic, "auto", "man") & ")"
End Function

Public Function NowWithMS() As String
    NowWithMS = Format(Now, "yyyymmdd hh:mm:ss") & Right(Format(Timer, "0.000"), 4)
End Function

Public Function Trace(ByVal msg As String, Optional ByVal forceWrite As Boolean = False, Optional ByVal forceDoEvents As Boolean = False, Optional ByVal forceDebug As Boolean = False)
    If ftState = ftClosing Then Exit Function
    
   If msg Like "FALSE" Then Stop
    
    If ThisWorkbook.readOnly Then
        GoTo Finalize:
    End If
    If msg = vbNullString Or Len(msg) = 0 Then
        GoTo Finalize:
    End If
    l_lastTraceMsg = msg
    
    If DebugInfo Or DebugMode Then
        Dim nw As String
        nw = NowWithMS
        Dim trcArr As Variant
        trcArr = Array(nw, msg, IIf(InStr(1, msg, "BusyWait", vbTextCompare) > 0, "", TrcSysInfo))
        If DebugInfo Then QueueTrace trcArr
        If DebugMode Then Debug.Print Join(Array(nw & " -- " & msg, IIf(InStr(1, msg, "BusyWait", vbTextCompare) > 0, "", TrcSysInfo)), ", ")
    End If
    
    
    If wsBusy Is ThisWorkbook.activeSheet Then
        wsBusy.UpdateMsg msg & " " & TrcSysInfo, forceDoEvents
        forceDoEvents = False
    ElseIf forceDoEvents Then
        DoEvents
    End If
    
Finalize:
    If Err.Number <> 0 Then
        Err.Clear
    End If
    
End Function
