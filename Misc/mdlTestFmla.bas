Attribute VB_Name = "mdlTestFmla"
Option Explicit
Option Base 1
Option Compare Text

Private fmlaRanges As New Collection

Public Function UpdateFmlaRanges(fmlaCell As Range, refCells As Range)
    Dim fKey As String
    fKey = BuildKey(fmlaCell)
    Debug.Print fKey
    If CollKeyExists(fmlaRanges, fKey) Then
        fmlaRanges.remove fKey
    End If
    fmlaRanges.Add Array(fmlaCell, refCells), Key:=fKey
End Function

Public Function FmlaRangeExists(fmlaCell As Range) As Boolean
    FmlaRangeExists = CollKeyExists(fmlaRanges, BuildKey(fmlaCell))
End Function

Private Function BuildKey(fmlaCell As Range)
    'allow fmlaCell to include more than one cell, but base key off first cell in range
    If Not fmlaCell Is Nothing Then
        BuildKey = Join(Array(fmlaCell.Worksheet.Parent.Name, fmlaCell.Worksheet.CodeName, fmlaCell(1, 1).Address), "_")
    Else
        BuildKey = CVErr(1004)
    End If
End Function

Private Function GetRefRange(fmlaCell As Range) As Range
    Dim storedRanges As Variant
    If CollKeyExists(fmlaRanges, BuildKey(fmlaCell)) Then
        storedRanges = CollItemByKey(fmlaRanges, BuildKey(fmlaCell))
        Set GetRefRange = storedRanges(UBound(storedRanges))
    End If
End Function

Private Function BuildFmla(fmlaCell As Range) As String
    '' this example will only work within the same workbook where [fmlaCell] exists
    Dim storedRanges As Variant
    Dim refRng As Range
    Set refRng = GetRefRange(fmlaCell)
    If Not refRng Is Nothing Then
        BuildFmla = "='" & refRng.Worksheet.Name & "'!" & refRng.Address(RowAbsolute:=True, ColumnAbsolute:=True)
        fmlaCell.Formula2 = BuildFmla
    End If
End Function

Private Function ReplaceWithFormula(fmlaCell As Range)
    If FmlaRangeExists(fmlaCell) Then
        With fmlaCell
            .ClearContents
            If IsNull(fmlaCell.PrefixCharacter) Then
                fmlaCell.PrefixCharacter = ""
                fmlaCell.numberFormat = "General"
            ElseIf Not fmlaCell.PrefixCharacter = "" Then
                fmlaCell.PrefixCharacter = ""
                fmlaCell.numberFormat = "General"
            End If
            fmlaCell.Formula2 = BuildFmla(fmlaCell)
        End With
    End If
End Function

Public Function TestIt()
    Dim wb As Workbook, sht1 As Worksheet, sht2 As Worksheet
    Set wb = Workbooks.Add
    Set sht1 = wb.Worksheets(1)
    Set sht2 = wb.Worksheets.Add(After:=sht1)
    
    sht1.Name = "Ref Ranges"
    sht2.Name = "Update Formulas"
    
    Dim i As Long
    With sht1
        For i = 1 To 100
            .Range("A" & i) = "Value: " & i
            .Range("B" & i) = "Date: " & CStr(DateAdd("h", i, Now()))
        Next i
    End With
    
    UpdateFmlaRanges sht2.Range("A1"), sht1.Range("A1:A5")
    UpdateFmlaRanges sht2.Range("C1:D1"), sht1.Range("B1:B4")

    ReplaceWithFormula sht2.Range("A1")
    ReplaceWithFormula sht2.Range("C1:D1")
    
    sht1.UsedRange.EntireColumn.AutoFit
    sht2.UsedRange.EntireColumn.AutoFit
    
    
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Returns item from collection by Key
''   If [key] does not exist in collection, error object with
''   error code 1004 is return
''   suggested use:
''
''   Dim colItem as Variant
''   colItem = CollectionItemByKey([collection], [expectedKey])
''
''   'If expecting object, use 'Set'
''    Set colItem = CollectionItemByKey([collection], [expectedKey])
''
''   If Not IsError(colItem) Then
''       'value was returned
''   Else
''       'error was returned
''   End if
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function CollItemByKey(ByRef col As Collection, ByVal Key)
On Error Resume Next
    Key = CStr(Key)
    If IsObject(col(Key)) Then
        If Err.number = 0 Then
            Set CollItemByKey = col(Key)
        End If
    Else
        If Err.number = 0 Then
            CollItemByKey = col(Key)
        End If
    End If
    If Err.number <> 0 Then
        Err.Clear
        CollItemByKey = CVErr(1004)
    End If
End Function
Private Function CollKeyExists(ByRef col As Collection, ByVal Key)
On Error Resume Next
    Key = CStr(Key)
    If IsObject(col(Key)) Then
        If Err.number = 0 Then
            CollKeyExists = True
        Else
            CollKeyExists = False
        End If
    Else
        If Err.number = 0 Then
            CollKeyExists = True
        Else
            CollKeyExists = False
        End If
    End If
    If Err.number <> 0 Then
        Err.Clear
    End If
End Function
