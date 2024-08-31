Attribute VB_Name = "pbDevUtil"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   DEVELOPER UTILITIES
'   (Requires Trust with VBA Object Model)
'   ** DEPENDENCIES **
'       1.  pbCommonUtil
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbDevUtil.bas
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1

Public Function DEVSel() As Range
    If IsDev Then
        Set DEVSel = Selection
    End If
End Function

Public Function DEVListWorkbooks()
    If IsDev Then
        Dim wbIdx As Long
        For wbIdx = 1 To Application.Workbooks.Count
            Debug.Print Concat(Format(wbIdx, "00"), ")", vbTab, Application.Workbooks(wbIdx).Name)
        Next wbIdx
    End If
End Function

Public Function TestArrTypes()

    Dim S() As String, l(1 To 2) As Long, v() As Variant
    
    S = Split("a,b,c", ",")
    l(1) = 1: l(2) = 2
    v = Array("a", 1, True)
    
    Debug.Print TypeName(S)
    Debug.Print TypeName(l)
    Debug.Print TypeName(v)
    
    Debug.Print VarType(S)
    Debug.Print VarType(l)
    Debug.Print VarType(v)
    
    Debug.Print EnumCompare(VarType(S), VbVarType.vbArray)
    Debug.Print EnumCompare(VarType(l), VbVarType.vbArray)
    Debug.Print EnumCompare(VarType(v), VbVarType.vbArray)

End Function

Public Function CheckDevKey()
    DevKeyOn
End Function
Public Function DevKeyOn()
    If stg.IsDeveloper Then
        Application.OnKey "`", "OnDevKey1"
    End If
End Function
Public Function DevKeyOff()
    Application.OnKey "`", ""
End Function
Public Function OnDevKey1()
'   If current user IsDev, pressing the dev key lands here
    If IsDev Then
        Beep
        ToggleDevMode stg
    End If
End Function
Function DEVMode1()
    On Error Resume Next
    If stg.IsDeveloper Then
        stg.stdDEVIsDevMode = True
        Beep
        EventsOff
        ScreenOff
        Dim cWS As Worksheet
        Set cWS = ThisWorkbook.activeSheet
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            ws.visible = xlSheetVisible
        Next ws
        UnprotectAllSheets
        cWS.Activate
        ScreenOn
        EventsOff
    End If
End Function



    Public Function ChangeCodeName(wksht As Worksheet, newCodeName As String, Optional unprotectPassword As String = CFG_PROTECT_PASSWORD)
    ''  EXAMPLE USAGE
    ''  ChangeCodeName ThisWorkbook,Sheet1,"wsNewCodeName"
        On Error Resume Next
        Dim wkbk As Workbook
        Set wkbk = wksht.Parent
        If wkbk.HasVBProject Then
            If wksht.protectContents Then
                wksht.Unprotect unprotectPassword
                If Err.number <> 0 Then
                    MsgBox wksht.CodeName & " needs to be unprotected, and could not be unprotected with password provided!"
                End If
                Exit Function
            End If
            wkbk.VBProject.VBComponents(wksht.CodeName).Properties("_CodeName").value = newCodeName
        End If
    End Function

Public Function ToggleDevMode(stg As pbSettings, Optional unprotectPassword = CFG_PROTECT_PASSWORD, Optional runAutoOpenIfFalse As Boolean = True)
    If IsDev = False Then Exit Function
    Dim wkbk As Workbook
    Set wkbk = stg.pbSettingsSheet.Parent
    
    Dim curDevMode As Boolean
    stg.stdDEVIsDevMode = Not stg.stdDEVIsDevMode
    
    If stg.stdDEVIsDevMode = True Then
        Application.EnableEvents = False
        Dim wks As Worksheet
        For Each wks In wkbk.Worksheets
            wks.visible = xlSheetVisible
            If wks.protectContents Then
                wks.Unprotect CStr(unprotectPassword)
            End If
        Next wks
    Else
        Application.EnableEvents = True
        If runAutoOpenIfFalse Then
            wkbk.RunAutoMacros xlAutoOpen
        End If
    End If
End Function

'Public Property Get DevMode() As Boolean
'    If IsDEV Then
'        If STG.Exists("DEV_MODE") Then
'            DevMode = CBool(STG.Setting("DEV_MODE"))
'        End If
'    End If
'End Property

'Public Property Let DevMode(ByVal dvMode As Boolean)
'    STG.Setting("DEV_MODE") = dvMode
'    If dvMode = False Then
'        DevEventsDisabled = False
'        STG.Setting(STG_DEV_PAUSELOCKING) = False
'    End If
'End Property

Public Property Get DevEventsDisabled() As Boolean
    If stg.Exists("DEV_EVENTS_OFF") = False Then
        DevEventsDisabled = False
    Else
        DevEventsDisabled = stg.Setting("DEV_EVENTS_OFF")
    End If
End Property
Public Property Let DevEventsDisabled(ByVal evDisabled As Boolean)
    stg.Setting("DEV_EVENTS_OFF") = evDisabled
End Property

Public Function CODE_MakeListObjectEnum(lstObj As ListObject, prefix)
    If IsDev = False Then Exit Function
    
    Dim col, i, charVal, newCol
    For Each col In lstObj.HeaderRowRange
        newCol = CStr(col.value)
        For i = 1 To Len(col.value)
            charVal = Asc(Mid(col.value, i, 1))
            If Not (charVal >= 65 And charVal <= 90) And Not (charVal >= 97 And charVal <= 122) Then
                newCol = Replace(newCol, Mid(col.value, i, 1), "")
            End If
        Next i
        Dim fullColName As String
        fullColName = Concat("'", col, "'")
        Debug.Print Concat("    ", prefix, newCol) & IIf(lstObj.ListColumns(CStr(col)).Index = 1, " = 1", "")
    Next col

End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   CLOSE ALL EDITORS/OPEN FILES IN IDE
'   * ARGUMENTS *
'    -visible(): Names of files to Make/Keep Open
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function IDEClear(ParamArray visible() As Variant)
    If IsDev = False Then Exit Function
    Application.ScreenUpdating = False
    Dim visCol As New Collection
    Dim IDX As Long
    If UBound(visible) >= LBound(visible) Then
        For IDX = LBound(visible) To UBound(visible)
            visCol.add visible(IDX), Key:=CStr(visible(IDX))
        Next IDX
    End If
    Dim vbProj As VBProject, vbComp As VBComponent
    For Each vbProj In Application.VBE.VBProjects
        For Each vbComp In vbProj.VBComponents
            If CollectionKeyExists(visCol, vbComp.Name) Then
                vbComp.CodeModule.CodePane.Window.visible = True
            Else
                vbComp.CodeModule.CodePane.Window.visible = False
            End If
        Next vbComp
    Next vbProj
    Application.ScreenUpdating = True
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   OUTPUT FORMULAS TO IMMEDIATE WINDOW FOR [forRange]
'    - Output formatted to be copied and ready to be used in code to set or check
'       Formula with VBA (escapes double quotes)
'    - If Range Not Supplied, Uses ActiveCell.Address
'    - Worksheet Must Be Unlocked
'    - If any Cell is in a ListObject, will show formulat for ListColumn
'    - Show Formula as R1C1 (default), otherwise, A1 style
'    - If [repositionColOffset] or [repositionRowOffset] <> 0, then Active Selection
'       is adjusted at end of function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function FormulaOutput(Optional forRange As Range, Optional r1c1Type As Boolean = True, Optional repositionColOffset As Long = 0, Optional repositionRowOffset As Long = 0) As String
    Dim rng As Range, colName As String, colIdx As Long, firstColIdx As Long
    If forRange Is Nothing Then
        Set rng = activeSheet.Range(ActiveCell.Address)
    Else
        Set rng = forRange
    End If
    If rng.Count > 1 Then
        Dim cRng As Range, cIdx As Long
        For Each cRng In forRange
            cIdx = cIdx + 1
            FormulaOutput cRng, r1c1Type, 0, 0
            If cIdx = forRange.Count Then
                If repositionColOffset + repositionRowOffset <> 0 Then
                    rng.Offset(rowOffset:=repositionRowOffset, ColumnOffset:=repositionColOffset).Select
                End If
            End If
        Next cRng
        Exit Function
    End If
    
    Dim wsName As String
    wsName = rng.Worksheet.Name
    If rng.Worksheet.protectContents Then
        Debug.Print "You need to unprotect " & rng.Worksheet.CodeName & "(" & rng.Worksheet.Name & ")"
        Exit Function
    End If
    If Not rng(1, 1).ListObject Is Nothing Then
        firstColIdx = rng(1, 1).ListObject.ListColumns(1).Range.column
        colIdx = rng(1, 1).column - firstColIdx + 1
        If rng.HasFormula Then
            Debug.Print Concat(wsName, "![", rng(1, 1).ListObject.Name, "].[", rng(1, 1).ListObject.ListColumns(colIdx).Name, "]", " (ListColumn Index: ", colIdx, ")")
        End If
    Else
        If rng.HasFormula Then Debug.Print Concat(wsName, "!", rng(1, 1).Address)
    End If
    Dim F As String
    If rng.HasFormula Then
        If r1c1Type Then
            F = rng.Formula2R1C1
        Else
            F = rng.Formula
        End If
        F = Replace(F, """", """""")
        If StringsMatch(F, vbNewLine, smContains) Then
            F = Replace(F, vbNewLine, " ", compare:=vbTextCompare)
        End If
        Debug.Print """" & F & """"
    End If
    FormulaOutput = F
    If rng.column + repositionColOffset >= 1 And rng.Row + repositionRowOffset >= 1 And (repositionColOffset + repositionRowOffset) <> 0 Then
        rng.Offset(rowOffset:=repositionRowOffset, ColumnOffset:=repositionColOffset).Select
    End If
    Set rng = Nothing
End Function




Public Function DEV_ListListObjects(Optional wkbk As Workbook)
    If wkbk Is Nothing Then
        Set wkbk = ThisWorkbook
    End If
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wkbk.Worksheets
        For Each lo In ws.ListObjects
            Debug.Print ConcatWithDelim(" ", "lstObj: ", lo.Name, "  ( ", ws.CodeName, " )")
        Next lo
    Next ws
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   DEVELOPER UTILITY TO LIST PROPERTIES OF CONNECTIONS
'   TO SHAREPOINT THAT ARE OLEDB CONNECTIONS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' Requires 'StringsMatch' Function and 'strMatchEnum'  from my pbCommon.bas module
'   pbCommon.bas: https://github.com/lopperman/just-VBA/blob/404999e6fa8881a831deaf2c6039ff942f1bb32d/Code_NoDependencies/pbCommon.bas
'   StringsMatch Function: https://github.com/lopperman/just-VBA/blob/404999e6fa8881a831deaf2c6039ff942f1bb32d/Code_NoDependencies/pbCommon.bas#L761C1-L761C1
'   strMatchEnum: https://github.com/lopperman/just-VBA/blob/404999e6fa8881a831deaf2c6039ff942f1bb32d/Code_NoDependencies/pbCommon.bas#L183
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function DEV_ListOLEDBConnections(Optional ByVal targetWorksheet, Optional ByVal connName, Optional ByVal wkbk As Workbook)
    ' if [targetWorksheet] provided is of Type: Worksheet, the worksheet name and code name will be converted to
    '   search criteria
    ' if [connName] is included, matches on 'Name like *connName*'
    ' if [wkbk] is not included, wkbk becomes ThisWorkbook
    Dim searchWorkbook As Workbook
    Dim searchName As Boolean, searchTarget As Boolean
    Dim searchSheetName, searchSheetCodeName, searchConnName As String
    Dim tmpWBConn As WorkbookConnection
    Dim tmpOleDBConn As OLEDBConnection
    Dim tmpCol As New Collection, shouldCheck As Boolean, targetRange As Range
    
    '   SET WORKBOOK TO EVALUATE
    If wkbk Is Nothing Then
        Set searchWorkbook = ThisWorkbook
    Else
        Set searchWorkbook = wkbk
    End If
    
    '   SET SEARCH ON CONN NAME CONDITION
    searchName = Not IsMissing(connName)
    If searchName Then searchConnName = CStr(connName)
        
    '   SET SEARCH ON TARGET SHEET CONDITION
    searchTarget = Not IsMissing(targetWorksheet)
    If searchTarget Then
        If StringsMatch(TypeName(targetWorksheet), "Worksheet") Then
            searchSheetName = targetWorksheet.Name
            searchSheetCodeName = targetWorksheet.CodeName
        Else
            searchSheetName = CStr(targetWorksheet)
            searchSheetCodeName = searchSheetName
        End If
    End If
    tmpCol.add Array(vbTab, "")
    tmpCol.add Array("", "")
    tmpCol.add Array("***** Sharepoint OLEDB Connections *****", searchWorkbook.Name)
    tmpCol.add Array("", "")
    For Each tmpWBConn In searchWorkbook.Connections
        If tmpWBConn.Ranges.Count > 0 Then
            Set targetRange = tmpWBConn.Ranges(1)
        End If
        shouldCheck = True
        If searchName And Not StringsMatch(tmpWBConn.Name, searchConnName, smContains) Then shouldCheck = False
        If shouldCheck And searchTarget Then
            If targetRange Is Nothing Then
                shouldCheck = False
            ElseIf Not StringsMatch(targetRange.Worksheet.Name, searchSheetName, smContains) And Not StringsMatch(targetRange.Worksheet.CodeName, searchSheetCodeName, smContains) Then
                shouldCheck = False
            End If
        End If
        If shouldCheck Then
            If tmpWBConn.Type = xlConnectionTypeOLEDB Then
                tmpCol.add Array("", "")
                tmpCol.add Array("*** CONNECTION NAME ***", tmpWBConn.Name)
                tmpCol.add Array("", "")
                If Not targetRange Is Nothing Then
                    tmpCol.add Array("TARGET WORKSHEET", targetRange.Worksheet.CodeName & "(" & targetRange.Worksheet.Name & ")")
                    tmpCol.add Array("WORKSHEET RANGE", targetRange.Address)
                End If
                tmpCol.add Array("REFRESH WITH REFRESH ALL", tmpWBConn.refreshWithRefreshAll)
                Set tmpOleDBConn = tmpWBConn.OLEDBConnection
                tmpCol.add Array("COMMAND TEXT", tmpOleDBConn.CommandText)
                tmpCol.add Array("CONNECTION", tmpOleDBConn.Connection)
                tmpCol.add Array("ENABLE REFRESH", tmpOleDBConn.enableRefresh)
                tmpCol.add Array("IS CONNECTED", tmpOleDBConn.IsConnected)
                tmpCol.add Array("MAINTAIN CONNECTION", tmpOleDBConn.maintainConnection)
                tmpCol.add Array("REFRESH ON FILE OPEN", tmpOleDBConn.refreshOnFileOpen)
                tmpCol.add Array("REFRESH PERIOD", tmpOleDBConn.RefreshPeriod)
                tmpCol.add Array("ROBUST CONNECT (xlRobustConnect)", tmpOleDBConn.RobustConnect)
                tmpCol.add Array("SERVER CREDENTIALS METHOD (xlCredentialsMethod)", tmpOleDBConn.serverCredentialsMethod)
                tmpCol.add Array("USE LOCAL CONNECTION", tmpOleDBConn.UseLocalConnection)
            End If
        End If
    Next tmpWBConn
    Dim cItem, useTab As Boolean
    For Each cItem In tmpCol
        Debug.Print ConcatWithDelim(":  ", UCase(IIf(useTab, vbTab & cItem(1), cItem(1))), cItem(2))
        useTab = True
    Next cItem
End Function







