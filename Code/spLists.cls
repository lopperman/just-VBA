VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "spLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' spLists v1.0.0
' (c) Paul Brower - https://github.com/lopperman/just-VBA
'
' General  Helper Utilities for Working with OleDBConnection objects (SharePoint Lists)
'
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

'   THIS CONTAINS SPECIFIC CONNECTION OBJECT DATA THAT WILL NEED TO BE REPLACED
'   WITH NAMES/PROPERTIES IN YOUR ENVIRONMENT

Option Explicit
Option Compare Text
Option Base 1

Private Const stgSettingUpdateDt As String = "UpdateDate"
Private Const REF_CALENDAR_KEY_COL As String = "CalendarKey"
Private Const REF_CALENDAR_VERSION_COL As String = "HolidayName"
Private Const CONN_NAME_CONFIG As String = "spFTConfig"
Private Const CONN_NAME_SETTINGS As String = "spFTSettings"
Private Const CONN_NAME_CALENDAR As String = "spFTCalendars"
Private Const CONN_NAME_GLPERIODS As String = "spFTGLPeriods"

' MORE THESE ENUMS TO COMMON MODULE IF NEEDED
Public Enum SPConns
    spSettings = 1
    spGLPeriods = 2
    spCalendars = 3
    spConfig = 4
End Enum
Public Enum strMatchEnum
    smEqual = 0
    smNotEqualTo = 1
    smContains = 2
    smStartsWithStr = 3
    smEndWithStr = 4
End Enum



'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'       GET ONLINE SETTINGS UPDATES
'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function UpdateSettings() As Boolean
On Error GoTo E:
    Dim failed As Boolean
    
    Dim tmpVersion As Variant
    Dim ole As OLEDBConnection

'   *** SUPER IMPORTANT ***
'   If Application.DisplayAlerts = False, then user will not be automatically
'   prompted if their Authorization Token has expired
    Application.DisplayAlerts = True
'   *** SUPER IMPORTANT ***
    
    Set ole = PreflightCheck(SPConns.spConfig)
    If Not ole Is Nothing Then
        UpdateSettings = ExecuteRefresh(ole)
    End If
    If UpdateSettings Then
        Set ole = PreflightCheck(SPConns.spSettings)
        If Not ole Is Nothing Then
            UpdateSettings = ExecuteRefresh(ole)
        End If
    End If
    If UpdateSettings Then
        Set ole = PreflightCheck(SPConns.spCalendars)
        If Not ole Is Nothing Then
            UpdateSettings = ExecuteRefresh(ole)
        End If
    End If
    If UpdateSettings Then
        Set ole = PreflightCheck(SPConns.spGLPeriods)
        If Not ole Is Nothing Then
            UpdateSettings = ExecuteRefresh(ole)
        End If
    End If
    
Finalize:
    On Error Resume Next
    
        'ADD YOUR OWN LOGIC FOR FAILURE PATHS
        'If failed Then LogError "ftUpdater.UpdateSettings - FAILED TO UPDATED 1 OR MORE LISTS FROM SHAREPOINT"
        UpdateSettings = Not failed
    
    Exit Function
E:
    failed = True
    'ErrorCheck "ftUpdater.UpdateSettings", ftERR_MessageIgnore
    Resume Finalize:

End Function

Private Function ExecuteRefresh(oleConn As OLEDBConnection) As Boolean
On Error Resume Next
    Dim successUpd As Boolean
    Dim alrt As Boolean, evts As Boolean
    alrt = Application.DisplayAlerts
    evts = Application.EnableEvents
    Application.DisplayAlerts = True
    Application.EnableEvents = False
    oleConn.Refresh
    successUpd = True
    If Err.number = 0 Then
        'UpdateMessage "Successfully Updated - " & oleConn.CommandText & " (" & oleConn.Connection & ")"
    Else
        'LogError "ftUpdater.ExecuteRefresh - " & oleConn.CommandText
        Err.Clear
        successUpd = False
    End If
    Application.DisplayAlerts = alrt
    Application.EnableEvents = evts
    ExecuteRefresh = successUpd
    
End Function

Private Function PreflightCheck(spconn As SPConns) As OLEDBConnection
On Error GoTo E:
    Dim failed As Boolean
    
    'make sure connection name we're expecint, exists
    'make sure Connection and OleDbConnection Properties are correct
    'make sure Connection is OleDb Type
    'only return OleDbConn if everything we CAN check, is valid
    
    Dim tmpWBConn As WorkbookConnection
    Dim tmpOleDBConn As OLEDBConnection
    For Each tmpWBConn In ThisWorkbook.Connections
        If tmpWBConn.Type = xlConnectionTypeOLEDB Then
            If StringsMatch(tmpWBConn.Name, GetWkbkConnName(spconn)) Then
                Set tmpOleDBConn = tmpWBConn.OLEDBConnection
                tmpWBConn.RefreshWithRefreshAll = False
                With tmpOleDBConn
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
            Set PreflightCheck = tmpOleDBConn
        End If
    
        Set tmpOleDBConn = Nothing
    
    Exit Function
E:
    failed = True
    'ErrorCheck "ftUpdater.PreflightCheck"
    Resume Finalize:
    
    
End Function

'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'
'   EXPECTED WORKBOOK CONNECTION NAME (STRING)
'
'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function GetWkbkConnName(ByVal spconn As SPConns) As String
    Select Case spconn
        Case SPConns.spConfig
            GetWkbkConnName = CONN_NAME_CONFIG
        Case SPConns.spSettings
            GetWkbkConnName = CONN_NAME_SETTINGS
        Case SPConns.spCalendars
            GetWkbkConnName = CONN_NAME_CALENDAR
        Case SPConns.spGLPeriods
            GetWkbkConnName = CONN_NAME_GLPERIODS
    End Select
End Function

'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'
'   GET WORKBOOK CONNECTION (QUERY) FROM SPConn Enum
'
'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function GetWkbkConn(ByVal spconn As SPConns) As WorkbookConnection
On Error Resume Next
    
    Set GetWkbkConn = ThisWorkbook.Connections(GetWkbkConnName(spconn))
    
    If Err.number <> 0 Then
        'LogError "ftUpdater.GetWkbkConn: spConn = " & spconn
        Err.Clear
    End If
    
End Function

'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'
'   GET OLE DB CONNECTION FROM SPConns Enum
'
'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function GetOleDbConn(ByVal spconn As SPConns) As OLEDBConnection
On Error Resume Next
    
    Dim wbkCn As WorkbookConnection
    Set wbkCn = GetWkbkConn(spconn)
    If Not wbkCn Is Nothing Then
        If wbkCn.Type = xlConnectionTypeOLEDB Then
            Set GetOleDbConn = wbkCn.OLEDBConnection
        End If
    End If
    Set wbkCn = Nothing

    If Err.number <> 0 Then
        'LogError "ftUpdater.GetOleDbConn: spConn = " & spconn
        Err.Clear
    End If

End Function


Public Function Dev_OleDBConnProperties(Optional ByVal spCon As SPConns = SPConns.spConfig, Optional ByVal saveAsODC As Boolean = False)
On Error Resume Next
    Dim c As WorkbookConnection
    Set c = ThisWorkbook.Connections(GetWkbkConnName(spCon))
    Debug.Print ""
    Debug.Print "'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~"
    Debug.Print "'   ***** CONNECTION QUERY (Expected/Actual Name: " & GetWkbkConnName(spCon) & " / " & c.Name & " ) *****"
    Dim ole As OLEDBConnection
    Set ole = c.OLEDBConnection
    
    '   ~~~ ~~~ OLEDBCONNECTION PROPERTIES ~~~ ~~~
    Debug.Print vbTab & "AlwaysUseConnectionFile - " & ole.AlwaysUseConnectionFile
    Debug.Print vbTab & "BackgroundQuery - " & ole.BackgroundQuery
    Debug.Print vbTab & "CommandText - " & ole.CommandText
    Debug.Print vbTab & "XlCommandType = XlCmdSql - " & ole.CommandType = xlCmdSql
    Debug.Print vbTab & "Connection: " & ole.Connection
    Debug.Print vbTab & "EnableRefresh: " & ole.EnableRefresh & " (by user)"
    Debug.Print vbTab & "IsConnected: " & ole.IsConnected
    Debug.Print vbTab & "MaintainConnection: " & ole.MaintainConnection
    Debug.Print vbTab & "RefreshDate: " & ole.RefreshDate
    Debug.Print vbTab & "Refreshing: " & ole.Refreshing
    Debug.Print vbTab & "RefreshOnFileOpen: " & ole.RefreshOnFileOpen
    Debug.Print vbTab & "RefreshPeriod: " & ole.RefreshPeriod & " (minutes)"
    Debug.Print vbTab & "RobustConnect: " & ole.RobustConnect
    Debug.Print vbTab & "SavePassword: " & ole.SavePassword
    Debug.Print vbTab & "ServerCredentials is xlCredentialsMethodIntegrated: " & ole.ServerCredentialsMethod = xlCredentialsMethodIntegrated
    Debug.Print vbTab & "ServerSSOApplicationID: " & ole.ServerSSOApplicationID
    Debug.Print vbTab & "SourceConnectionFile: " & ole.SourceConnectionFile
    
    Debug.Print "'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~"
    
    If saveAsODC Then
'        Dim savePath As String
'        savePath = pbFileSys.ChooseFolder("Select Save Directory")
'        Debug.Print "Saving To: " & PathCombine(False, savePath, c.Name & ".odc")
'        ole.Refresh
'        ole.saveAsODC PathCombine(False, savePath, c.Name & ".odc")
    End If
    
    
    Exit Function
    
    
    '   ~~~ ~~~ METHODS ~~~ ~~~
    'ThisWorkbook.Connections(CONN_NAME_CONFIG).OLEDBConnection.SaveAsODC PathCombine(False, Application.DefaultFilePath, CONN_NAME_CONFIG & ".odc")
    
    'CancelRefresh - Cancels all refresh operations in progress for the specified OLE DB connection.
    'MakeConnection - Establishes a connection for the specified OLE DB connection.
    'Reconnect - Drops and then reconnects the specified connection.
    'Refresh - Refreshes an OLE DB connection.
    'SaveAsODC - Saves the OLE DB connection as a Microsoft Office Data Connection file.

End Function



' this is usually in a public common module
Public Function StringsMatch( _
    ByVal str1 As Variant, ByVal _
    str2 As Variant, _
    Optional smEnum As strMatchEnum = strMatchEnum.smEqual, _
    Optional compMethod As VbCompareMethod = vbTextCompare) As Boolean
    
'       IF NEEDED, PUT THIS ENUM AT TOP OF A STANDARD MODULE
        'Public Enum strMatchEnum
        '    smEqual = 0
        '    smNotEqualTo = 1
        '    smContains = 2
        '    smStartsWithStr = 3
        '    smEndWithStr = 4
        'End Enum
        
    str1 = CStr(str1)
    str2 = CStr(str2)
    Select Case smEnum
        Case strMatchEnum.smEqual
            StringsMatch = StrComp(str1, str2, compMethod) = 0
        Case strMatchEnum.smNotEqualTo
            StringsMatch = StrComp(str1, str2, compMethod) <> 0
        Case strMatchEnum.smContains
            StringsMatch = InStr(1, str1, str2, compMethod) > 0
        Case strMatchEnum.smStartsWithStr
            StringsMatch = InStr(1, str1, str2, compMethod) = 1
        Case strMatchEnum.smEndWithStr
            If Len(str2) > Len(str1) Then
                StringsMatch = False
            Else
                StringsMatch = InStr(Len(str1) - Len(str2) + 1, str1, str2, compMethod) = Len(str1) - Len(str2) + 1
            End If
    End Select
End Function

