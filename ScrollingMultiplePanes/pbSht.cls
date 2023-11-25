VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  Manage Layout, Structure, Custom Properties of Worksheet
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbSht.cls - DEMO
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

'Public Function TestScrollPanes(wksName as string)
'
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Worksheets(wksName)
'
'    Dim pbs As New pbSht
'    Dim splitOnRow as Long, splitOnCol as Long
'    splitOnRow = 8
'    splitOnCol = 0 '(meaning, no column split)
'    pbs.Initialize ws, splitOnRow, splitOnCol
'    'If ignoreInactive = False, the ScrollDefault will force sheet to be visible and active
'    pbs.ScrollDefault ignoreInactive:=False
'
'End Function

Option Explicit
Option Compare Text
Option Base 1

Private wksht As Worksheet
Private splitR As Long
Private splitC As Long


Public Property Get Self() As pbSht
    Set Self = Me
End Property

Public Property Get Sheet() As Worksheet
    Set Sheet = wksht
End Property

' Note: Zero ('0') is a valid value for spRow or spCol -- and means 'remove the split if it exists
Public Function Initialize(ws As Worksheet, Optional spRow, Optional spCol)
    
    If Not wksht Is Nothing Then
        Err.Raise 1004, Description:="Sheet can only be set once!"
        Exit Function
    End If
    
    splitR = -1
    splitC = -1
    
    If Not IsMissing(spRow) Then splitR = CLng(spRow)
    If Not IsMissing(spCol) Then splitC = CLng(spCol)
    
    Set wksht = ws
 
End Function

Public Function Hide()
    If wksht.visible = xlSheetVisible And VisibleSheets >= 2 Then
        wksht.visible = xlSheetVeryHidden
    End If
End Function

Public Function Show()
    If wksht.visible <> xlSheetVisible Then
        wksht.visible = xlSheetVisible
    End If
End Function

Public Function Activate()
    Dim evts As Boolean: evts = Application.EnableEvents
    Application.EnableEvents = False
    Show
    If Not ActiveWindow.activeSheet Is wksht Then
        wksht.Activate
    End If
    Application.EnableEvents = evts
End Function

Public Function FirstVisibleRow(Optional startingAt As Long) As Long
    Dim visRow As Long, rIdx As Long
    If startingAt > 1 Then rIdx = startingAt - 1
    Do While visRow = 0
        rIdx = rIdx + 1
        If wksht.Cells(rIdx, 1).EntireRow.Hidden = False Then
            visRow = rIdx
        End If
    Loop
    FirstVisibleRow = rIdx
End Function

Public Function FirstVisibleCol(Optional startingAt As Long) As Long
    Dim visCol As Long, cIdx As Long
    If startingAt > 1 Then cIdx = startingAt - 1
    Do While visCol = 0
        cIdx = cIdx + 1
        If wksht.Cells(1, cIdx).EntireColumn.Hidden = False Then
            visCol = cIdx
        End If
    Loop
    FirstVisibleCol = cIdx
End Function

Public Function ScrollDefault(Optional ignoreInactive As Boolean = True)
    On Error Resume Next
    If ignoreInactive = True And Not ActiveWindow.activeSheet Is wksht Then
        Exit Function
    End If
    
    Dim evts As Boolean: evts = Application.EnableEvents
    Dim scrn As Boolean: scrn = Application.ScreenUpdating
    Dim intr As Boolean: intr = Application.Interactive
    Dim paneIdx As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Interactive = False
    
    Activate
    
    With ActiveWindow
        If splitR = -1 Then splitR = .splitRow
        If splitC = -1 Then splitC = .splitColumn
    
        If splitR <> .splitRow Then
            .splitRow = splitR
        End If
        If splitC <> .splitColumn Then
            .splitColumn = splitC
        End If
        If splitC > 0 Or splitR > 0 Then
            If Not .freezePanes Then
                .freezePanes = True
            End If
        End If
        If .Panes.Count = 1 Then
            .ScrollColumn = FirstVisibleCol
            .ScrollRow = FirstVisibleRow
        ElseIf .Panes.Count = 2 And .splitRow > 0 And .freezePanes = True Then
            .Panes(1).ScrollColumn = FirstVisibleCol
            .Panes(1).ScrollRow = FirstVisibleRow
            .Panes(2).ScrollRow = FirstVisibleRow(.splitRow) + 1
        ElseIf .Panes.Count = 2 And .splitColumn > 0 And .freezePanes = True Then
            .Panes(1).ScrollColumn = FirstVisibleCol
            .Panes(1).ScrollRow = FirstVisibleRow
            .Panes(2).ScrollColumn = FirstVisibleCol(.splitColumn) + 1
        ElseIf .Panes.Count = 4 And .freezePanes = True Then
            .Panes(1).ScrollRow = FirstVisibleRow
            .Panes(3).ScrollRow = FirstVisibleRow(.splitRow) + 1
            .Panes(2).ScrollColumn = FirstVisibleCol(.splitColumn) + 1
        End If
    End With
    
    Application.EnableEvents = evts
    Application.ScreenUpdating = scrn
    Application.Interactive = intr
End Function

Private Property Get VisibleSheets() As Long
    Dim w As Worksheet, retV As Long
    For Each w In ThisWorkbook.Worksheets
        If w.visible = xlSheetVisible Then
            retV = retV + 1
        End If
    Next
    VisibleSheets = retV
End Property



