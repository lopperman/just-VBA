VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsBusy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Option Base 1

Private lastTmr As Single
Private forceScreenRefreshSeconds As Long
Private lOnCloseGoTo As Worksheet

Private Sub Worksheet_Calculate()
    
End Sub


Public Property Get ForceRefreshSeconds() As Long
    If forceScreenRefreshSeconds <= 0 Then forceScreenRefreshSeconds = 1
    ForceRefreshSeconds = forceScreenRefreshSeconds
End Property

Public Property Let ForceRefreshSeconds(seconds As Long)
    If seconds >= 1 Then
        forceScreenRefreshSeconds = seconds
    End If
End Property

Public Property Get OnCloseGoTo() As Worksheet
    If Not lOnCloseGoTo Is Nothing Then
        Set OnCloseGoTo = lOnCloseGoTo
    End If
End Property
Public Property Set OnCloseGoTo(vl As Worksheet)
    If Not lOnCloseGoTo Is Nothing Then
        Set lOnCloseGoTo = Nothing
    End If
    Set lOnCloseGoTo = vl
End Property

Public Function OnActivate()
'   OnActivate should be called explicitely, as it's likely you've already disabled Application Events
    lastTmr = 0
End Function

Public Function OnDeactivate()
    'IGNORE
End Function

Public Function ClearMessage()
On Error Resume Next
    Dim evts As Boolean: evts = Events
    EventsOff
    wsBusy.Range("busyMessage").value = ""
    wsBusy.Range("busyMessageGeneral").value = ""
    Events = evts
    If Err.number <> 0 Then
        ftBeep btError
        If IsDev Then Stop
        Err.Clear
    End If
End Function

Public Function UpdateMessage(mMsg As String)
On Error Resume Next
    wsBusy.Range("A1:C1").EntireColumn.ColumnWidth = CDbl(2)
    wsBusy.Range("D1:J1").EntireColumn.ColumnWidth = CDbl(10)
    
    Me.Range("D8").value = mMsg
    If Err.number <> 0 Then Err.Clear
End Function

Public Function Show(msg, Optional summaryMsg)
    If Not ThisWorkbook.activeSheet Is Me Then
        If lOnCloseGoTo Is Nothing Then
            Set lOnCloseGoTo = ThisWorkbook.activeSheet
        End If
    End If
    Dim evts As Boolean
    evts = Events
    EventsOff
    UpdateTitle stg.Setting("BusyTitle")
    If Not Me.visible = xlSheetVisible Then
        Me.visible = xlSheetVisible
    End If
    Me.Activate
    Application.GoTo Me.Range("A1"), Scroll:=True
    UpdateMessage CStr(msg)
    If Not IsMissing(summaryMsg) Then
        UpdateSummary CStr(summaryMsg)
    End If
    If Application.ScreenUpdating = False Then
        Application.ScreenUpdating = True
        DoEvents
        Application.ScreenUpdating = False
    Else
        DoEvents
    End If
    Events = evts
End Function
Public Function Hide(Optional ignoreGoToScreen As Boolean = False)
    If Me.visible = xlSheetVisible Then
        
        If lOnCloseGoTo Is Nothing And ignoreGoToScreen = False Then
            Set lOnCloseGoTo = wsDashboard
        End If
        Dim evts As Boolean
        evts = Events
        EventsOff
        If ignoreGoToScreen = False Then
            If Not lOnCloseGoTo.visible = xlSheetVisible Then
                lOnCloseGoTo.visible = xlSheetVisible
            End If
            lOnCloseGoTo.Activate
        End If
        Me.visible = xlSheetVeryHidden
        Events = evts
        ScreenOn
    End If
End Function

Public Function UpdateSummary(sMsg As String)
On Error Resume Next
    Dim evts As Boolean: evts = Events
    EventsOff
    wsBusy.Range("busyMessageGeneral").value = sMsg
    If Err.number <> 0 Then Err.Clear
    Events = evts
End Function
Public Function UpdateTitle(mMsg As String)
On Error Resume Next
    Dim evts As Boolean: evts = Events
    EventsOff
    
    wsBusy.Range("busyTitle").value = mMsg
    If Err.number <> 0 Then Err.Clear
    Events = evts
End Function

'Public Function ShowBusyScreen(msg As String, msgCategory As String, Optional title As String, Optional showAfterClose As Worksheet)
'
'    If Not showAfterClose Is Nothing Then
'        Set lOnCloseGoTo = showAfterClose
'    End If
'
'    wsBusy.Range("busyMessage").value = msg
'    wsBusy.Range("busyMessageGeneral").value = msgCategory
'    If Len(title) > 0 Then wsBusy.Range("busyTitle").value = title
'
'    Dim scrn As Boolean: scrn = Application.ScreenUpdating
'    ScreenOn
'    If Not wsBusy.visible = xlSheetVisible Then wsBusy.visible = xlSheetVisible
'    If Not wsBusy Is ThisWorkbook.ActiveSheet Then wsBusy.Activate
'    DoEvents
'    Application.ScreenUpdating = scrn
'
'End Function

Public Function PrivBusyWait(msg As String, waitSeconds As Long, ignoreBeep As Boolean)

    wsBusy.ForcedWaitMessage msg, waitSeconds, ignoreBeep
    
End Function


Public Function ForcedWaitMessage(msg As String, waitSeconds As Long, Optional ignoreBeep As Boolean = True)
On Error GoTo e:
    
    Dim evts As Boolean, scrn As Boolean
    evts = Events
    scrn = Application.ScreenUpdating
        
    If Not Application.ScreenUpdating Then Application.ScreenUpdating = True
    EventsOff
    
    If Me.visible <> xlSheetVisible Then
        Me.visible = xlSheetVisible
    End If
    If Not wsBusy Is ThisWorkbook.activeSheet Then
        wsBusy.Activate
    End If
    
    UpdateMessage msg
    DoEvents
    
    If waitSeconds > CLng(0) And ignoreBeep = False Then
        ftBeep btBusyWait
    End If
    ForceWait waitSeconds
    
    lastTmr = Timer
Finalize:
    On Error Resume Next
    Application.ScreenUpdating = scrn
    Events = evts
    If Err.number <> 0 Then Err.Clear
    Exit Function
e:
    Beep
    Err.Clear
    Resume Finalize:
End Function






Private Sub Worksheet_Change(ByVal Target As Range)

End Sub
