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

Private Const SP_CONN_RETRY_SECONDS As Single = 300
Private lastSPFail As Single



Private l_SPChecked As Boolean
Private l_SPFile As Boolean


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   DEVELOPER UTILITY TO LIST PROPERTIES OF CONNECTIONS
'   TO SHAREPOINT THAT ARE OLEDB CONNECTIONS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function DEV_ListSharepointConnections(Optional wkbk As Workbook)
    If wkbk Is Nothing Then
        Set wkbk = ThisWorkbook
    End If
    Dim tmpWBConn As WorkbookConnection
    Dim tmpOleDBConn As OLEDBConnection
    Dim tmpCol As New Collection
    tmpCol.Add Array("Sharepoint OLEDB Connections", wkbk.Name)
    For Each tmpWBConn In wkbk.Connections
        If tmpWBConn.Type = xlConnectionTypeOLEDB Then
        
            tmpCol.Add Array("  ", "")
            tmpCol.Add Array("***** CONNECTION *****", "")
            tmpCol.Add Array("CONNECTION NAME", tmpWBConn.Name)
            If Not tmpWBConn.Ranges Is Nothing Then
                If tmpWBConn.Ranges.Count > 0 Then
                    tmpCol.Add Array("TARGET WORKSHEET", tmpWBConn.Ranges(1).Worksheet.Name)
                    tmpCol.Add Array("WORKSHEET RANGE", tmpWBConn.Ranges(1).Address)
                Else
                    tmpCol.Add Array("TARGET WORKSHEET: ", " ** N/A **")
                End If
            End If
            tmpCol.Add Array("REFRESH WITH REFRESH ALL", tmpWBConn.refreshWithRefreshAll)
            Set tmpOleDBConn = tmpWBConn.OLEDBConnection
            tmpCol.Add Array("COMMAND TEXT", tmpOleDBConn.CommandText)
            tmpCol.Add Array("CONNECTION", tmpOleDBConn.Connection)
            tmpCol.Add Array("ENABLE REFRESH", tmpOleDBConn.enableRefresh)
            tmpCol.Add Array("IS CONNECTED", tmpOleDBConn.IsConnected)
            tmpCol.Add Array("MAINTAIN CONNECTION", tmpOleDBConn.maintainConnection)
            tmpCol.Add Array("REFRESH ON FILE OPEN", tmpOleDBConn.refreshOnFileOpen)
            tmpCol.Add Array("REFRESH PERIOD", tmpOleDBConn.RefreshPeriod)
            tmpCol.Add Array("ROBUST CONNECT (xlRobustConnect)", tmpOleDBConn.RobustConnect)
            tmpCol.Add Array("SERVER CREDENTIALS METHOD (xlCredentialsMethod)", tmpOleDBConn.serverCredentialsMethod)
            tmpCol.Add Array("USE LOCAL CONNECTION", tmpOleDBConn.UseLocalConnection)
        End If
    Next tmpWBConn
    Dim cItem, useTab As Boolean
    For Each cItem In tmpCol
        Debug.Print ConcatWithDelim(":  ", UCase(IIf(useTab, vbTab & cItem(1), cItem(1))), cItem(2))
        useTab = True
    Next cItem
End Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  SHAREPOINT LIST UTILITIES
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function VerifyConnection(spConn As String) As OLEDBConnection
On Error GoTo e:

    Dim failed As Boolean
    'make sure connection name we're expecint, exists
    'make sure Connection and OleDbConnection Properties are correct
    'make sure Connection is OleDb Type
    'only return OleDbConn if everything we CAN check, is valid
    Dim tmpWBConn As WorkbookConnection
    Dim tmpOleDBConn As OLEDBConnection
    Dim bw As Boolean: bw = wsBusy Is ThisWorkbook.activeSheet
    For Each tmpWBConn In ThisWorkbook.Connections
        If tmpWBConn.Type = xlConnectionTypeOLEDB Then
            If StringsMatch(tmpWBConn.Name, spConn) Then
                BusyWait "Checking Connection Properties for: " & tmpWBConn.Name, ignoreIfHidden:=True
                Set tmpOleDBConn = tmpWBConn.OLEDBConnection
                tmpWBConn.refreshWithRefreshAll = False
                With tmpOleDBConn
''''TODO:  MOVE TO CONSTANTS
                    If .enableRefresh = False Then .enableRefresh = True
                    If .maintainConnection = True Then .maintainConnection = False
                    If .BackgroundQuery = True Then .BackgroundQuery = False
                    If .refreshOnFileOpen = True Then .refreshOnFileOpen = False
                    If .SourceConnectionFile <> "" Then .SourceConnectionFile = ""
                    If .AlwaysUseConnectionFile = True Then .AlwaysUseConnectionFile = False
                    If .SavePassword = True Then .SavePassword = False
                    If .serverCredentialsMethod <> xlCredentialsMethodIntegrated Then .serverCredentialsMethod = xlCredentialsMethodIntegrated
                    ''If .EnableRefresh = True Then .EnableRefresh = False
                End With
                Exit For
            End If
        End If
    Next tmpWBConn

Finalize:
    On Error Resume Next
        If bw = False Then wsBusy.visible = xlSheetVeryHidden
        If Not tmpOleDBConn Is Nothing And Not failed Then
            Set VerifyConnection = tmpOleDBConn
        End If
        Set tmpOleDBConn = Nothing
    Exit Function
e:
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
        LogERROR "GetWkbkConn: spConn = " & spConn
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
        LogERROR "ftUpdater.GetOleDbConn: spConn = " & spConn
        Err.Clear
    End If

End Function


Public Function RefreshSPList(ByVal spConnName As String, Optional ByVal ignoreRetryPeriod As Boolean = False) As Boolean
On Error Resume Next
    If LocalMode Then
        RefreshSPList = True
        Exit Function
    End If
    If ignoreRetryPeriod = False Then
        If lastSPFail > 0 And (Timer - lastSPFail) <= SP_CONN_RETRY_SECONDS Then
            LogFORCED "Ignoring request to refresh data from SharePoint connection (" & spConnName & ") because last attempt failed " & Timer - lastSPFail & " seconds ago and assume cannot connect to SharePoint"
            RefreshSPList = False
            Exit Function
        End If
    End If
    
    Dim oleConn As OLEDBConnection
    Dim successUpd As Boolean
    Dim alrt As Boolean, evts As Boolean
    Set oleConn = GetOleDbConn(spConnName)
    BusyWait "Refreshing Data: " & oleConn.CommandText, ignoreIfHidden:=True
     
    VerifyConnection spConnName
     
    LogTRACE "Begin SP List Refresh for: " & spConnName
    alrt = Application.DisplayAlerts
    evts = Application.EnableEvents
    ' ~~~ ALERTS MUST BE ON SO THAT AUTHENTICATION PROMPT CAN POP UP IF SSO TOKEN IS EXPIRED ~~~
    Application.DisplayAlerts = True
    EventsOff
    oleConn.Refresh
    If Err.number = 0 Then
        successUpd = True
        LogTRACE "Successfully Updated - " & oleConn.CommandText & " (" & oleConn.Connection & ")"
        BusyWait "Successfully Updated - " & oleConn.CommandText & " (" & oleConn.Connection & ")", ignoreIfHidden:=True
    Else
        LogERROR "RefreshSPList - " & oleConn.CommandText
        BusyWait "RefreshSPList - " & oleConn.CommandText, ignoreIfHidden:=True
        Err.Clear
        successUpd = False
    End If
    If Not successUpd Then
        lastSPFail = Timer
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
    Dim tFileName As String
    If wbk Is Nothing Then
        tFileName = FullWbNameCorrected(ThisWorkbook)
    Else
        tFileName = FullWbNameCorrected(wbk)
    End If
    IsSharePointFile = IsSharePointFileName(tFileName)
    
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


