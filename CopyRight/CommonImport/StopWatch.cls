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
Private l_key As Variant

Public Property Get swKey() As Variant
    swKey = l_key
End Property

Public Sub resetTimer()
    timerStart = 0
    timerStop = 0
End Sub
Public Sub StartTimer(Optional swKey As Variant)
    If Not IsMissing(swKey) Then
        If Not IsObject(swKey) Then l_key = swKey
    End If
    resetTimer
    timerStart = Timer
End Sub
Public Sub StopTimer()
    timerStop = Timer
End Sub

Public Property Get ElapsedSeconds() As Double
    ElapsedSeconds = Round(Timer - timerStart, 3)
End Property

Public Property Get totalSeconds() As Double
    totalSeconds = Round(timerStop - timerStart, 3)
End Property
Public Function Result() As String
    Result = CStr(totalSeconds) & " seconds"
End Function
Public Function ResultAsTime() As String
    Dim retV As String
        Dim startTime As Single, timeElapsed As Single, hours As Long, mins As Long, seconds As Single
        timeElapsed = totalSeconds
        hours = timeElapsed / 3600
        timeElapsed = timeElapsed - hours * 3600
        mins = timeElapsed / 60
        timeElapsed = timeElapsed - mins * 60
        seconds = timeElapsed
        If hours > 0 Then
            retV = Concat(Format(hours, "00"), ":", Format(mins, "00"), ":", Format(seconds, "00.00"), " (hours)")
        ElseIf mins > 0 Then
            retV = Concat(Format(mins, "00"), ":", Format(seconds, "00.00"), " (minutes)")
        Else
            retV = Concat(Format(seconds, "00.00"), " (seconds)")
        End If
        ResultAsTime = retV
End Function


