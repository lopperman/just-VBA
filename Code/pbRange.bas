Attribute VB_Name = "pbRange"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' pbLRange v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' General  Helper Utilities for Working with Ranges
'
' @module pbRange
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1



Public Function HasValidation(ByVal rng As Range, Optional isType As XlDVType = -1) As Boolean
'   Return True if selected Range has Custom Validation for Input
On Error Resume Next
    Dim retV As Boolean
    Dim vType As Long
    vType = rng.Validation.Type
    If Err.Number <> 0 Then
        retV = False
        GoTo Finalize:
    End If
    If isType > -1 Then
        If vType = isType Then
            retV = True
        End If
    Else
        retV = True
    End If
    If Err.Number <> 0 Then Err.Clear
Finalize:
    HasValidation = retV
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function
Public Function UniqueRowNumberInRange(ByVal Target As Range) As Long()

    Dim tmpD As New Dictionary
    tmpD.CompareMode = BinaryCompare
    Dim areaIDX As Long, rwIDX As Long
    Dim realRow As Long
    For areaIDX = 1 To Target.Areas.Count
        For rwIDX = 1 To Target.Areas(areaIDX).Rows.Count
            realRow = Target.Areas(areaIDX).Rows(rwIDX).Row
            If Not tmpD.Exists(realRow) Then
                tmpD(realRow) = realRow
            End If
        Next rwIDX
    Next areaIDX
    
    If tmpD.Count > 0 Then
        Dim retV() As Long, rwCount As Long
        ReDim retV(1 To tmpD.Count)
        Dim ky As Variant
        For Each ky In tmpD.Keys
            retV(rwCount + 1) = ky
            rwCount = rwCount + 1
        Next ky
        UniqueRowNumberInRange = retV
    End If

End Function
Public Function IsListObjectHeader(ByVal Target As Range) As Boolean
'   Returns True if target(1,1) intersects with any ListObject HeaderRowRange
    If Target Is Nothing Then Exit Function
    If Not Target.ListObject Is Nothing Then
        If Not Intersect(Target(1, 1), Target.ListObject.HeaderRowRange) Is Nothing Then
            IsListObjectHeader = True
        End If
    End If
End Function
Public Property Get LastColumnWithData(wks As Worksheet) As Long
    If Not wks.usedRange Is Nothing Then
        LastColumnWithData = wks.usedRange.Columns.Count + (wks.usedRange.column - 1)
    End If
End Property

Public Property Get LastRowWithData(wks As Worksheet, Optional column As Variant) As Long
    Dim ret As Long
    ret = -1
    If Not IsMissing(column) Then
        If IsNumeric(column) Then
            ret = wks.Cells(wks.Rows.Count, CLng(column)).End(xlUp).Row
        Else
            ret = wks.Cells(wks.Rows.Count, CStr(column)).End(xlUp).Row
        End If
    Else
        ret = wks.usedRange.Rows.Count + (wks.usedRange.Row - 1)
    End If
    LastRowWithData = ret
End Property
Public Function GetA1CellRef(fromRng As Range, Optional colOffset As Long = 0, Optional rowCount As Long = 1, Optional colCount As Long = 1, Optional rowOffset As Long = 0, Optional fixedRef As Boolean = False, Optional visibleCellsOnly As Boolean = False) As String
'   return A1 style reference (e.g. "A10:A116") from selection
'   Optional offsets, resized ranges supported
    Dim tmpRng As Range
    Set tmpRng = fromRng.offset(rowOffset, colOffset)
    If colCount > 1 Or rowCount > 1 Then
        Set tmpRng = tmpRng.Resize(rowCount, colCount)
    End If
    If visibleCellsOnly Then
        Set tmpRng = tmpRng.SpecialCells(xlCellTypeVisible)
    End If
    GetA1CellRef = tmpRng.Address(fixedRef, fixedRef)
    Set tmpRng = Nothing
End Function
