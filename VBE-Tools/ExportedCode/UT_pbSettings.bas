Attribute VB_Name = "UT_pbSettings"
Option Explicit
Option Compare Text
Option Base 1



Public Function stgDEV() As pbSettings
    On Error Resume Next
    Static stgObj As pbSettings
    If stgObj Is Nothing Then
        Set stgObj = New pbSettings
        stgObj.DEV_TEST
        stgObj.stdAutoHide = False
        stgObj.pbSettingsSheet.Activate
    End If
    If Err.number = 0 Then
        If Not stgObj Is Nothing Then
            If stgObj.ValidConfig Then
                Set stgDEV = stgObj
            End If
        End If
    Else
        Err.Clear
    End If
End Function

