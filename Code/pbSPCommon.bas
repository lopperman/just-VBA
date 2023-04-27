Attribute VB_Name = "pbSPCommon"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  SHAREPOINT UTILITIES
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbSPCommon.bas
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Option Explicit
Option Compare Text
Option Base 1
Option Private Module

Private l_SPChecked As Boolean
Private l_SPFile As Boolean


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  SHAREPOINT LIST UTILITIES
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function VerifyConnection(spConn As String) As OLEDBConnection
On Error GoTo E:

    If IsDEV Then
        If Len(spConn) < 5 Then
            Beep
            Stop
        End If
    End If

    Dim failed As Boolean
    'make sure connection name we're expecint, exists
    'make sure Connection and OleDbConnection Properties are correct
    'make sure Connection is OleDb Type
    'only return OleDbConn if everything we CAN check, is valid
    Dim tmpWBConn As WorkbookConnection
    Dim tmpOleDBConn As OLEDBConnection
    For Each tmpWBConn In ThisWorkbook.Connections
        If tmpWBConn.Type = xlConnectionTypeOLEDB Then
            If StringsMatch(tmpWBConn.Name, spConn) Then
                BusyWait "Checking Connection Properties for: " & tmpWBConn.Name, ignoreIfHidden:=True
                Set tmpOleDBConn = tmpWBConn.OLEDBConnection
                tmpWBConn.RefreshWithRefreshAll = False
                With tmpOleDBConn
''''TODO:  MOVE TO CONSTANTS
                    If .EnableRefresh = False Then .EnableRefresh = True
                    If .MaintainConnection = True Then .MaintainConnection = False
                    If .BackgroundQuery = True Then .BackgroundQuery = False
                    If .RefreshOnFileOpen = True Then .RefreshOnFileOpen = False
                    If .SourceConnectionFile <> "" Then .SourceConnectionFile = ""
                    If .AlwaysUseConnectionFile = True Then .AlwaysUseConnectionFile = False
                    If .SavePassword = True Then .SavePassword = False
                    If .ServerCredentialsMethod <> xlCredentialsMethodIntegrated Then .ServerCredentialsMethod = xlCredentialsMethodIntegrated
                End With
                Exit For
            End If
        End If
    Next tmpWBConn

Finalize:
    On Error Resume Next
        If Not tmpOleDBConn Is Nothing And Not failed Then
            Set VerifyConnection = tmpOleDBConn
        End If
        Set tmpOleDBConn = Nothing
    Exit Function
E:
    failed = True
    ErrorCheck "ftUpdater.PreflightCheck"
    Resume Finalize:
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   GET WORKBOOK CONNECTION (QUERY)
'   FROM SPConnection Enum
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function GetWkbkConn(ByVal spConn As String) As WorkbookConnection
On Error Resume Next
    Set GetWkbkConn = ThisWorkbook.Connections(spConn)
    If Err.number <> 0 Then
        LogError "GetWkbkConn: spConn = " & spConn
        Err.Clear
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   GET OLE DB CONNECTION FROM SPConnection
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function GetOleDbConn(ByVal spConn As String) As OLEDBConnection
On Error Resume Next
    Dim wbkCn As WorkbookConnection
    Set wbkCn = GetWkbkConn(spConn)
    If Not wbkCn Is Nothing Then
        If wbkCn.Type = xlConnectionTypeOLEDB Then
            Set GetOleDbConn = wbkCn.OLEDBConnection
        End If
    End If
    Set wbkCn = Nothing
    If Err.number <> 0 Then
        LogError "ftUpdater.GetOleDbConn: spConn = " & spConn
        Err.Clear
    End If

End Function

Public Function RefreshSPList(ByVal spConn As String) As Boolean
On Error Resume Next
    If LocalMode Then
        RefreshSPList = True
        Exit Function
    End If
    Dim oleConn As OLEDBConnection
    Dim successUpd As Boolean
    Dim alrt As Boolean, evts As Boolean
    Set oleConn = GetOleDbConn(spConn)
    BusyWait "Refreshing Data: " & oleConn.CommandText, ignoreIfHidden:=True
    LogDEV "Begin SP List Refresh for: " & GetWkbkConnName(spConn)
    alrt = Application.DisplayAlerts
    evts = Application.EnableEvents
    ' ~~~ ALERTS MUST BE ON SO THAT AUTHENTICATION PROMPT CAN POP UP IF SSO TOKEN IS EXPIRED ~~~
    Application.DisplayAlerts = True
    EventsOff
    oleConn.Refresh
    If Err.number = 0 Then
        successUpd = True
        LogDEV "Successfully Updated - " & oleConn.CommandText & " (" & oleConn.Connection & ")"
        BusyWait "Successfully Updated - " & oleConn.CommandText & " (" & oleConn.Connection & ")", ignoreIfHidden:=True
    Else
        LogError "RefreshSPList - " & oleConn.CommandText
        BusyWait "RefreshSPList - " & oleConn.CommandText, ignoreIfHidden:=True
        Err.Clear
        successUpd = False
    End If
    Application.DisplayAlerts = alrt
    Application.EnableEvents = evts
    RefreshSPList = successUpd
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  SHAREPOINT CHECK-IN / CHECK-OUT
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function TrySharePointCheckout()
On Error Resume Next
    ThisWorkbook.Saved = True
    Application.DisplayAlerts = False
    If NeedsCheckOut(ThisWorkbook) Then
        CheckOutFromSharePoint
    Else
        MsgBox_FT "This Workbook currently cannot be checked out.", vbOKOnly + vbInformation, "Cannot Check Out"
        With ThisWorkbook
            
            ThisWorkbook.Close SaveChanges:=False
        End With
    End If
    If Err.number <> 0 Then Err.Clear
End Function

Public Property Get SPFile() As Boolean
    If l_SPChecked = True Then
        SPFile = l_SPFile
    Else
        l_SPChecked = True
        l_SPFile = IsSharePointFileName(FullWbNameCorrected(ThisWorkbook))
        SPFile = l_SPFile
    End If
End Property

Public Function CheckoutQuit()
    Beep
    With ThisWorkbook
        Application.DisplayAlerts = False
        ThisWorkbook.Close SaveChanges:=False
        Application.DisplayAlerts = True
    End With
End Function

Public Function CheckOutFromSharePoint() As Boolean
On Error Resume Next
    If NeedsCheckOut = False Then
       MsgBox_FT "This file cannot be checked out"
       Exit Function
    End If
    Dim fullNameEncoded As String: fullNameEncoded = FullWbNameCorrected(ThisWorkbook)
            
    Application.DisplayAlerts = False
    Workbooks.CheckOut fullNameEncoded
    DoEvents
    MsgBox IIf(NeedsCheckIn, "Checked out successfully - Please Re-open", "Not able to check out file")
    If Err.number <> 0 Then Err.Clear
End Function

Public Function IsSharePointFileName(fullFilePath As String) As Boolean
    If StringsMatch(fullFilePath, "HTTP", smStartsWithStr) Or StringsMatch(fullFilePath, "HTTP", smContains) Then
        If StringsMatch(fullFilePath, "/personal/", smContains) Or StringsMatch(fullFilePath, "_com", smContains) Or StringsMatch(fullFilePath, "onedrive.apx", smContains) Then
            IsSharePointFileName = False
        Else
            IsSharePointFileName = True
        End If
    Else
        IsSharePointFileName = False
    End If
End Function

Public Function IsSharePointFile(Optional wbk As Workbook) As Boolean
    Dim tfileName As String
    If wbk Is Nothing Then
        tfileName = FullWbNameCorrected(ThisWorkbook)
    Else
        tfileName = FullWbNameCorrected(wbk)
    End If
    IsSharePointFile = IsSharePointFileName(tfileName)
    
    If Not wbk Is Nothing Then
        Set wbk = Nothing
    End If
End Function

Public Function NeedsCheckOut(Optional ByVal wbk As Workbook) As Boolean
        If wbk Is Nothing Then
            If SPFile = False Then
                NeedsCheckOut = False
                Exit Function
            End If
            Set wbk = ThisWorkbook
        End If
        If IsSharePointFile(wbk) = False Then
            NeedsCheckOut = False
        Else
            If CanCheckOut(wbk) And NeedsCheckIn(wbk) = False Then
                NeedsCheckOut = True
            End If
        End If
        If Not wbk Is Nothing Then
            Set wbk = Nothing
        End If
End Function

Public Function NeedsCheckIn(Optional ByVal wbk As Workbook) As Boolean
    If wbk Is Nothing Then
        If SPFile = False Then
            NeedsCheckIn = False
            Exit Function
        End If
        Set wbk = ThisWorkbook
    End If
    If IsSharePointFile(wbk) Then
        If wbk.CanCheckIn Then
            NeedsCheckIn = True
        End If
    End If
    If Not wbk Is Nothing Then
        Set wbk = Nothing
    End If
    
End Function

Private Function CanCheckOut(Optional ByVal wbk As Workbook) As Boolean
    If wbk Is Nothing Then
        If SPFile = False Then
            CanCheckOut = False
            Exit Function
        End If
        
        Set wbk = ThisWorkbook
    End If
    If IsSharePointFile(wbk) = False Then
        CanCheckOut = False
    Else
        If Workbooks.CanCheckOut(FullWbNameCorrected(wbk)) Then
            CanCheckOut = True
        Else
            CanCheckOut = False
        End If
    End If
    If Not wbk Is Nothing Then
        Set wbk = Nothing
    End If
End Function

