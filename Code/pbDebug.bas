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
'Private l_trcQueue As Collection

'Public Property Get TraceQueue() As Collection
'    Set TraceQueue = l_trcQueue
'End Property
'Public Function ClearTraceQueue()
'    Set l_trcQueue = Nothing
'End Function

Public Property Get debugReleaseOverride() As Boolean
    If l_debugReleaseOverride Then
        debugReleaseOverride = True
    Else
        debugReleaseOverride = False
    End If
    
End Property
Public Property Let debugReleaseOverride(vl As Boolean)
    l_debugReleaseOverride = vl
    If vl Then
        If DebugMode Then
            pbLog.SetLogOptions llTrace
        Else
            pbLog.SetLogOptions llWarn
        End If
    End If
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
'Conditional Compiler Args
'conDebug=1
Public Property Get LocalMode() As Boolean
    #If conLocal Then
        LocalMode = True
    #Else
        LocalMode = False
    #End If
End Property



Public Property Get DebugInfo() As Boolean
    If debugReleaseOverride Then
        DebugInfo = True
    Else
        DebugInfo = False
    End If

End Property

Public Function ToggleDebugFromUI()
    ToggleDebugOverride
    MsgBox_FT "Debugging To Log File: " & DebugInfo
End Function

Public Function ToggleDebugOverride()
    debugReleaseOverride = Not debugReleaseOverride
    
End Function



Public Function Assert(condition As Boolean)
'   DEV NOTE:  Assert Requires Conditional Compiler Constant 'Debugging = 1'
'   If Assert Failes, Step through to go back to failed caller
    If Not CanAssert Then Exit Function
    Debug.Assert condition
    
End Function

Public Function DebugPrint(Optional ByVal stmnt As Variant)
    If IsMissing(stmnt) Then stmnt = ""
        Trace stmnt, fromDebugPrint:=True
'    If DebugInfo Or DebugMode Then
'    End If
End Function

Public Function ShutdownDebug()
   'ClearTraceQueue
End Function

Public Property Get LastTraceMsg() As String
    LastTraceMsg = l_lastTraceMsg
End Property



Public Function TraceSessionStart(Optional sessionName As String = vbNullString)
    l_lastTraceTimer = Timer
    l_traceSession = IIf(Not sessionName = vbNullString, sessionName, CStr("Trace-" & CStr(Timer)))
    Trace "Starting: " & l_traceSession, forceDoEvents:=True
End Function

'Public Function DumpTraceIfAvail()
'    If Len(DUMP_TRACE_CALL_NAME) > 0 Then
'        Application.Run DUMP_TRACE_CALL_NAME
'        DoEvents
'    End If
'End Function

Public Function TraceSessionEnd()
    If Len(l_traceSession & vbNullString) = 0 Then l_traceSession = "UNK Trace Session"
    Trace "Completed: " & l_traceSession & vbNullString, forceDoEvents:=True
    l_traceSession = vbNullString
'    DumpTraceIfAvail
End Function

Public Function TrcSysInfo() As String
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
    TrcSysInfo = retV
    
    
    
'    TrcSysInfo =  & ", S-" & IIf(Application.ScreenUpdating, "*ON*", "off") & ", I-" & IIf(Application.Interactive, "*ON*", "off") & ", C-" & IIf(Application.Calculation = xlCalculationAutomatic, "auto", "man") & ")"

End Function

Public Function NowWithMS() As String
    NowWithMS = Format(Now, "yyyymmdd hh:mm:ss") & Right(Format(Timer, "0.000"), 4)
End Function

Public Function Trace(ByVal msg As String, Optional ByVal forceWrite As Boolean = False, Optional ByVal forceDoEvents As Boolean = False, Optional ByVal forceDebug As Boolean = False, Optional fromDebugPrint As Boolean = False)
    
    If msg = vbNullString Or Len(msg) = 0 Then
        GoTo Finalize:
    End If
    If ThisWorkbook.ReadOnly Then
        LogWarn msg, True
        GoTo Finalize:
    End If
    l_lastTraceMsg = msg
    If forceWrite Or forceDebug Then
        LogTrace msg, True
    Else
        LogTrace msg
    End If
    
Finalize:
    If Err.number <> 0 Then
        Err.Clear
    End If
    
End Function

'Public Function TraceQueueCount() As Long
'    Dim tmpCount As Long
'    If Not TraceQueue Is Nothing Then
'        tmpCount = TraceQueue.count
'    End If
'    TraceQueueCount = tmpCount
'End Function

'Public Function DumpTraceQueue()
'On Error GoTo E:
'    Dim failed As Boolean
'
'    pbPerf.Check
'
'    Dim newArr() As Variant
'
'    If TraceQueue Is Nothing Then Exit Function
'    If TraceQueue.count = 0 Then Exit Function
'    Dim newIdx As Long, colIDX As Long
'    Dim ky As Variant, nextRow As Long
'    Dim trc As Variant
'
'    nextRow = LastPopulatedRow(wsDebug, 2) + 1
'    If nextRow < 8 Then nextRow = 8
'
'    Dim checkTraceArray As ArrInformation
'    If TraceQueue.count > 0 Then
'        checkTraceArray = ArrayInfo(TraceQueue(1))
'        If checkTraceArray.Dimensions > 0 Then
'            ReDim newArr(1 To TraceQueue.count, 1 To checkTraceArray.Ubound_first)
'            Dim tcIDX As Long
'            For tcIDX = 1 To TraceQueue.count
'                trc = TraceQueue(tcIDX)
'                For colIDX = 1 To checkTraceArray.Ubound_first
'                    newArr(tcIDX, colIDX) = trc(colIDX)
'                Next colIDX
'            Next tcIDX
'
'            With wsDebug
'                .Range("B" & nextRow & ":D" & nextRow).Resize(rowSize:=TraceQueue.count).Value2 = newArr
'            End With
'        End If
'    End If
'
'Finalize:
'    On Error Resume Next
'
'    ClearTraceQueue
'    If ArrDimensions(newArr) > 0 Then Erase newArr
'
'
'    If Err.Number <> 0 Then Err.Clear
'    Exit Function
'E:
'    failed = True
'    ErrorCheck
'    Resume Finalize
'
'End Function
