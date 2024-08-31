VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  Manage Collection of VB Components
'  Intended Use: Comparing VBA Code
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module VBItems.cls
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1

Public Function CompareVBComp(wkbk1 As Workbook, vbComp1 As String, wkbk2 As Workbook, vbComp2 As String)
    On Error Resume Next
    Dim wbOut As Workbook, wsOut As Worksheet
    Dim vb1 As VBItemProps, vb2 As VBItemProps
    Set vb1 = New VBItemProps
    vb1.Populate wkbk1, wkbk1.VBProject.VBComponents(vbComp1)
    Set vb2 = New VBItemProps
    vb2.Populate wkbk2, wkbk2.VBProject.VBComponents(vbComp2)
    
    Set wbOut = Workbooks.add
    Set wsOut = wbOut.Worksheets(1)
    
    With wsOut
        .Range("B2") = "FILE 1: " & wkbk1.Name
        .Range("B3") = "Object 1: " & vbComp1
        .Range("D2") = "FILE 2: " & wkbk2.Name
        .Range("D3") = "Object 2: " & vbComp2
    End With
    
    Dim s1, s2, i1 As Long, i2 As Long, lines1 As Collection, lines2 As Collection
    Dim writeRow As Long, arr() As Variant
    writeRow = 5
    Set lines1 = vb1.GetLines()
    Set lines2 = vb2.GetLines()
    Dim firstMismatchRow As Long
    For i1 = 1 To lines1.Count
        If lines2.Count >= i1 Then
            ReDim arr(1 To 1, 1 To 4)
            arr(1, 1) = "File1: " & Format(i1, "0000")
            arr(1, 2) = lines1(i1)
            arr(1, 3) = "File2: " & Format(i1, "0000")
            arr(1, 4) = lines2(i1)
            With wsOut.Range("B" & writeRow).Resize(ColumnSize:=4)
                .value = arr
                If StringsMatch(arr(1, 2), arr(1, 4)) = True Then
                    .Interior.color = 13168833
                Else
                    .Interior.color = 65535
                    If firstMismatchRow = 0 Then firstMismatchRow = writeRow
                End If
            End With
            If IsMethodStart(arr(1, 2)) Then
                wsOut.Range("C" & writeRow).Font.Bold = True
                wsOut.Range("C" & writeRow).Font.color = COLOR_BLUEBERRY
            End If
            If IsMethodEnd(arr(1, 2)) Then
                wsOut.Range("C" & writeRow).Font.Bold = False
                wsOut.Range("C" & writeRow).Font.color = 192
            End If
            If IsMethodStart(arr(1, 4)) Then
                wsOut.Range("E" & writeRow).Font.Bold = True
                wsOut.Range("E" & writeRow).Font.color = COLOR_BLUEBERRY
            End If
            If IsMethodEnd(arr(1, 4)) Then
                wsOut.Range("E" & writeRow).Font.Bold = False
                wsOut.Range("E" & writeRow).Font.color = 192
            End If
            writeRow = writeRow + 1
        Else
            ReDim arr(1 To 1, 1 To 4)
            arr(1, 1) = "File1: " & Format(i1, "0000")
            arr(1, 2) = lines1(i1)
            arr(1, 3) = "File2: " & Format(i1, "0000")
            arr(1, 4) = "' #MISSING# "
            With wsOut.Range("B" & writeRow).Resize(ColumnSize:=4)
                .value = arr
                If StringsMatch(arr(1, 2), arr(1, 4)) = True Then
                    .Interior.color = 13168833
                Else
                    .Interior.color = 8628721
                End If
            End With
            If IsMethodStart(arr(1, 2)) Then
                wsOut.Range("C" & writeRow).Font.Bold = True
                wsOut.Range("C" & writeRow).Font.color = COLOR_BLUEBERRY
            End If
            If IsMethodEnd(arr(1, 2)) Then
                wsOut.Range("C" & writeRow).Font.Bold = False
                wsOut.Range("C" & writeRow).Font.color = 192
            End If
            If IsMethodStart(arr(1, 4)) Then
                wsOut.Range("E" & writeRow).Font.Bold = True
                wsOut.Range("E" & writeRow).Font.color = COLOR_BLUEBERRY
            End If
            If IsMethodEnd(arr(1, 4)) Then
                wsOut.Range("E" & writeRow).Font.Bold = False
                wsOut.Range("E" & writeRow).Font.color = 192
            End If
            
            writeRow = writeRow + 1
        End If
    Next
    
    wbOut.Activate
    wsOut.Activate
    wbOut.Windows(1).Left = 120
    wbOut.Windows(1).Top = 110
    wbOut.Windows(1).Width = 1855
    wbOut.Windows(1).Height = 1012
    wbOut.Windows(1).DisplayGridlines = False
    
    wsOut.UsedRange.EntireColumn.AutoFit
    wsOut.UsedRange.WrapText = True
    wsOut.Range("C:C").ColumnWidth = 70
    wsOut.Range("E:E").ColumnWidth = 70
    wsOut.UsedRange.WrapText = True
    If firstMismatchRow > 0 Then
        Application.GoTo wsOut.Range("A" & firstMismatchRow), Scroll:=True
    Else
        Application.GoTo wsOut.Range("A1"), Scroll:=True
    
    End If
    Set wsOut = Nothing
    Set wbOut = Nothing

End Function

Public Function IsMethodStart(sLine) As Boolean
    Dim terms As New Collection
    terms.add "Public "
    terms.add "Private "
    terms.add "Internal "
    Dim chkLine As String
    chkLine = Trim(sLine)
    Dim chkTerm
    For Each chkTerm In terms
        If StringsMatch(chkLine, chkTerm, smStartsWithStr) Then
            IsMethodStart = True
            Exit For
        End If
    Next chkTerm
End Function
Public Function IsMethodEnd(sLine) As Boolean
    Dim terms As New Collection
    terms.add "End Function"
    terms.add "End Sub"
    terms.add "End Property"
    terms.add "End Function"
    Dim chkLine As String
    chkLine = Trim(sLine)
    Dim chkTerm
    For Each chkTerm In terms
        If StringsMatch(chkLine, chkTerm, smStartsWithStr) Then
            IsMethodEnd = True
            Exit For
        End If
    Next chkTerm
End Function
