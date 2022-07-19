Attribute VB_Name = "pbPivot"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' pbPivot v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' Pivot Table Utilities and Dynamic Pivot Table Management
'
' @module pvPivot
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1

'   ~~~ ~~~ pbPivot Will CREATE THIS WORKSHEET if it does not exist
'   ~~~ ~~~ Name of Worksheet for pbPivot to Create Dynamic Pivot Tables
'   ~~~ ~~~ DO NOT CREATE ANYTHING ON THIS SHEET
'   ~~~ ~~~ pbPivot Will Create And Manage All Control, Layout, Formatting, Etc
Public Const pbPivot_DYNAMIC_SHEET_NAME As String = "Dynamic Pivot"

'~~~ ~~~ ~~~ ~~~~ ~~~ ~~~
'~~~ ~~~ TESTING ~~~ ~~~
Public Function tstpf(fldName As String) As PivotField
    Dim pf As PivotField
    Set pf = devpvt.PivotFields(fldName)
    Set tstpf = pf
    Stop
End Function
Public Function tstpvt()
    Dim pt As PivotTable
    Set pt = devpvt
    Stop
End Function

Public Function tstd()
    If wsDynamicPivot.PivotTables.Count = 1 Then
        devpvt.TableRange2.Clear
    End If
End Function
Public Function pvtS()
    Dim btn As Button
    For Each btn In wsDynamicPivot.buttons
        Debug.Print btn.caption
        Debug.Print "Left, Top", btn.left, btn.Top
    Next btn
End Function

'~~~ ~~~ ~~~ ~~~~ ~~~ ~~~
'~~~ ~~~ ~~~ ~~~~ ~~~ ~~~

' ~~~ ~~~ PIVOT WIZARD BUTTON EVENT HANDLERS ~~~
Public Function pbPivotBtn_ResetLayout()
    Beep
    Debug.Print Application.Caller
End Function
Public Function pbPivotBtn_Subtotals()
    Beep
    Debug.Print Application.Caller

End Function
Public Function pbPivotBtn_FilterFields()
    Beep
    Debug.Print Application.Caller

End Function
Public Function pbPivotBtn_ColumnFields()
    Beep
    Debug.Print Application.Caller

End Function
Public Function pbPivotBtn_RowFields()
    Beep
    Debug.Print Application.Caller

End Function
Public Function pbPivotBtn_ValueFields()
    Beep
    Debug.Print Application.Caller

End Function



Private Function RowFieldCount(pvt As PivotTable) As Long
    RowFieldCount = pvt.rowFields.Count
End Function
Private Function ColumnFieldCount(pvt As PivotTable) As Long
    ColumnFieldCount = pvt.ColumnFields.Count
End Function
Private Function PageFieldCount(pvt As PivotTable) As Long
    PageFieldCount = pvt.pageFields.Count
End Function
Private Function DataFieldCount(pvt As PivotTable) As Long
    DataFieldCount = pvt.dataFields.Count
End Function

' ~~~ ~~~ CREATE NEW PIVOT TABLE ~~~ ~~~
Public Function CreatePivot(srcListobj As ListObject, pvtName As String, Optional destRng As Range) As PivotTable
    
     
    
    'Create a new empty PivotTable
    ThisWorkbook.PivotCaches.Create( _
        sourceType:=xlDatabase, SourceData:=srcListobj.Name, Version:=6) _
        .CreatePivotTable TableDestination:=destRng, _
        tableName:=pvtName, DefaultVersion:=6
        
    Dim pvt As PivotTable
    Set pvt = destRng.Worksheet.PivotTables(pvtName)
        
        'set the various pivottable options you want
        With pvt
            .AllowMultipleFilters = True
            .ShowTableStyleColumnHeaders = False
            .ShowTableStyleRowHeaders = False
            .LayoutRowDefault = xlTabularRow
            .EnableDrilldown = False
            .EnableFieldList = True
            .EnableWizard = True
            .EnableWriteback = False
            .RepeatItemsOnEachPrintedPage = True
            .ShowPageMultipleItemLabel = False
            .RowAxisLayout xlTabularRow
            .SubtotalHiddenPageItems = False
            .SaveData = False
    
            .TableStyle2 = "PivotStyleLight2"
            .ShowTableStyleColumnHeaders = True
            .ShowTableStyleRowHeaders = True
            .ShowTableStyleColumnStripes = False
            .ShowTableStyleRowStripes = True
            .InGridDropZones = False
            .NullString = "Blank"
            .DisplayFieldCaptions = True
            .ShowDrillIndicators = False
            .RepeatAllLabels xlRepeatLabels
            .RowAxisLayout xlTabularRow
            .ShowTableStyleRowStripes = True
        
        End With
        
        Set CreatePivot = pvt
        Set pvt = Nothing
        
End Function

' ~~ Detemine Pivot Field Number Format by looking at source ListObject ~~
Private Function FindNumberFormat(ByRef pvt As PivotTable, fldName As Variant) As Variant
    Dim lo As ListObject
    Set lo = wt(pvt.SourceData)
    If Not lo Is Nothing Then
        If lo.listRows.Count > 0 Then
            If lo.Range.Worksheet.ProtectContents Then pbUnprotectSheet lo.Range.Worksheet
            FindNumberFormat = lo.ListColumns(fldName).DataBodyRange(1, 1).numberFormat
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function UpdateRowField( _
      ByRef pvt As PivotTable _
    , ByVal fldName As Variant _
    , ByVal visible As Boolean _
    , Optional ByVal position As Variant _
    , Optional ByVal showSubTotal As Variant _
    , Optional numberFormat As Variant)

    Dim pf As PivotField
    Set pf = pvt.PivotFields(fldName)
    If Not pf Is Nothing Then
        With pf
            If .orientation <> xlRowField Then
                position = pvt.rowFields.Count + 1
                .orientation = xlRowField
                .position = position
            Else
                If Not IsMissing(position) Then
                    If position <> .position And position <= pvt.rowFields.Count Then
                        .position = position
                    Else
                        .position = pvt.rowFields.Count
                    End If
                End If
            End If
            If visible = False Then
                .orientation = xlHidden
            End If
            Dim nbrFormat As Variant: nbrFormat = FindNumberFormat(pvt, fldName)
            If Len(nbrFormat & vbNullString) > 0 Then
                .dataRange.numberFormat = nbrFormat
            End If
        End With
        pf.Subtotals(1) = showSubTotal
    Else
        Beeper
        Trace "Error 'UpdateRowField': was unable to find pivotfield: " & fldName
    End If

End Function

Public Function VerifyDynamicPivotSheet() As Boolean
    If Len(pbPivot_DYNAMIC_SHEET_NAME) = 0 Then
        Err.Raise 17, "pbPivot.pbPivot_DYNAMIC_SHEET_NAME Cannot Be Empty"
    End If
On Error Resume Next
    Dim ws                                As Worksheet
    Dim failed                           As Boolean
    
    Set ws = Worksheets(pbPivot_DYNAMIC_SHEET_NAME)
    If Err.Number <> 0 Then
        Err.Clear
        Set ws = CreateNewDynamicPivotSheet
    End If
On Error GoTo E:
    With ws
        .Range("A1").EntireColumn.ColumnWidth = 1.17
        .Range("A8:A100").Interior.color = 1137094
        .Range("A8").EntireRow.RowHeight = 10
        .Range("A8:Z8").Interior.color = 1137094
        .Range("A1:A7").EntireRow.RowHeight = 16
        .Range("B4").EntireColumn.ColumnWidth = 18
        
        With .Range("B5:B6")
            .Merge
            .Cells(1, 1).value = "pbPivot"
            .Interior.color = 14395790
            .Font.color = 6740479
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        .Hyperlinks.Add Anchor:=.Range("B5:B6"), Address:= _
            "https://github.com/lopperman/VBA-pbUtil", ScreenTip:= _
            "View VBA-pbUtil on GitHub", TextToDisplay:="pbPivot"
        
    End With
    VerifyButton "btnPvtResetLayout", "Reset Layout", "pbPivotBtn_ResetLayout", 134, 45
    VerifyButton "btnPvtSubtotals", "Sub-totals", "pbPivotBtn_Subtotals", 134, 78
    VerifyButton "btnPvtFilterFields", "FILTER Fields", "pbPivotBtn_FilterFields", 257, 45
    VerifyButton "btnPvtColFields", "COLUMN Fields", "pbPivotBtn_ColumnFields", 257, 78
    VerifyButton "btnPvtRowFields", "ROW Fields", "pbPivotBtn_RowFields", 380, 45
    VerifyButton "btnPvtValueFields", "VALUE Fields", "pbPivotBtn_ValueFields", 380, 78

Finalize:
    On Error Resume Next
    
    VerifyDynamicPivotSheet = Not failed
    
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:

End Function
Private Function CreateNewDynamicPivotSheet() As Worksheet
    Dim tWS As Worksheet
    Set tWS = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    tWS.Name = pbPivot_DYNAMIC_SHEET_NAME
    Set CreateNewDynamicPivotSheet = tWS
    Set tWS = Nothing
End Function
Private Function VerifyButton(btnName As String, btnCaption As String, onActn As String, btnLeft As Variant, btnTop As Variant)
On Error Resume Next
    Dim btn As Button
    Set btn = Worksheets(pbPivot_DYNAMIC_SHEET_NAME).buttons(btnName)
    If Err.Number <> 0 Then
        Err.Clear
        Set btn = CreateButton(btnName, btnLeft, btnTop)
    End If
    With btn
        .Locked = True
        .LockedText = True
        .Name = btnName
        .caption = btnCaption
        .onAction = onActn
        .Font.Size = 14
        .Font.color = 9851952
        .Font.Bold = True
        .AutoScaleFont = False
        .Placement = XlPlacement.xlFreeFloating
        .visible = True
        .Enabled = True
        .left = btnLeft
        .Top = btnTop
        .width = 120
        .height = 30
    End With
End Function
Private Function CreateButton(btnName As String, btnLeft As Variant, btnTop As Variant) As Button
    Dim newButton As Button
    Set newButton = Worksheets(pbPivot_DYNAMIC_SHEET_NAME).buttons.Add(left:=btnLeft, Top:=btnTop, width:=120, height:=30)
    Set CreateButton = newButton
    Set newButton = Nothing
End Function

