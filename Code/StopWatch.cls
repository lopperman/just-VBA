VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Private timerStart As Double
Private timerStop As Double
Private lastSplit As Double
Private l_key As Variant

Public Property Get swKey() As Variant
    swKey = l_key
End Property

Public Sub resetTimer()
    timerStart = 0
    timerStop = 0
End Sub
Public Sub startTimer(Optional ByVal swKey As Variant)
    If Not IsMissing(swKey) Then
        If Not IsObject(swKey) Then l_key = swKey
    End If
    resetTimer
    timerStart = Timer
    lastSplit = timerStart
End Sub
Public Sub StopTimer()
    timerStop = Timer
End Sub

Public Function GetSplit() As String
    Dim totalElapsed As Double: totalElapsed = ElapsedSeconds
    Dim prevSplit As Double: prevSplit = lastSplit
    lastSplit = Timer
    Dim splitSeconds As Double
    splitSeconds = Round(lastSplit - prevSplit, 6)
    GetSplit = "This Split: " & splitSeconds & " seconds" & vbNewLine & "(Total Elapsed: " & totalElapsed & " seconds)"
End Function

Public Function NewSplit() As Double
    Dim prevSplit As Double: prevSplit = lastSplit
    lastSplit = Timer
    Dim splitSeconds As Double
    splitSeconds = Round(lastSplit - prevSplit, 6)
    NewSplit = splitSeconds
End Function

Public Property Get ElapsedSeconds() As Double
    ElapsedSeconds = Round(Timer - timerStart, 6)
End Property

Public Property Get totalSeconds() As Double
    totalSeconds = Round(timerStop - timerStart, 6)
End Property
Public Function result() As String
    result = CStr(totalSeconds) & " seconds"
End Function
Public Function ResultAsTime() As String
On Error Resume Next
    Dim retV As String
        Dim startTime As Single, timeElapsed As Single, Hours As Long, mins As Long, seconds As Single
        timeElapsed = totalSeconds
        Hours = timeElapsed / 3600
        timeElapsed = timeElapsed - Hours * 3600
        mins = timeElapsed / 60
        timeElapsed = timeElapsed - mins * 60
        seconds = timeElapsed
        If Hours > 0 Then
            retV = Concat(Format(Hours, "00"), ":", Format(mins, "00"), ":", Format(seconds, "00.00"), " (hours)")
        ElseIf mins > 0 Then
            retV = Concat(Format(mins, "00"), ":", Format(seconds, "00.00"), " (minutes)")
        Else
            retV = Concat(Format(seconds, "00.00"), " (seconds)")
        End If
        ResultAsTime = retV
    If Err.number <> 0 Then
        ResultAsTime = "Error: " & Err.Description
        Err.Clear
    End If
End Function


