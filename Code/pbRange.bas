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



Public Function FindDuplicateRows(rng As Range, ParamArray checkRangeCols() As Variant) As Dictionary
    'EXAMPLE CALL:   set  [myDictionary] = FindDuplicateRows(Worksheets(1).Range("B5:C100"))
    'EXAMPLE CALL:   set  [myDictionary] = FindDuplicateRows(Worksheets(1).Range("B5:H100"), 1,3,4)
        'Since Range Start on Column B, the columns that would be used to check duplicates would be B, D, E (the 1st, 3rd, and 4th columns in the range)

    'RETURNS DICTIONARY WHERE KEY=WORKSHEET ROW AND VALUE = NUMBER OF DUPLICATES
    'If No Value is passed in for 'checkRangeCols', then the entire row in the ranges will be compared to find duplicates
    'If  'rng' contains multiple areas (for example you passed in something like [range].SpecialCells(xlCellTypeVisible),
        'Then All Areas will be checked for Column Consistency (i.e. All areas in Range Must have identical Column Number and total Columns)
        'If all areas in Range to match column structure, error is raised
On Error GoTo E:
    Dim failed As Boolean
      
    ' ~~~ ~~~ check for mismatched range area columns ~~~ ~~~
    Dim firstCol As Long, totCols As Long
    Dim areaIDX As Long
    If rng.Areas.Count >= 1 Then
        firstCol = rng.Areas(1).column
        totCols = rng.Areas(1).Columns.Count
        For areaIDX = 1 To rng.Areas.Count
            If Not rng.Areas(areaIDX).column = firstCol _
                Or Not rng.Areas(areaIDX).Columns.Count = totCols Then
                Err.Raise 17, Description:="FindDuplicateRows can not support mismatched columns for multiple Range Areas"
            End If
        Next areaIDX
    End If
    
    Dim retDict As New Dictionary, tmpDict As New Dictionary, compareColCount As Long, tmpIdx As Long
    Dim checkCols() As Long
    retDict.CompareMode = TextCompare
    tmpDict.CompareMode = TextCompare
    
    If rng.Areas.Count = 1 And rng.Rows.Count = 1 Then
        GoTo Finalize:
    End If
    ' ~~~ ~~~ Determine Number of columns being compared for each row  ~~~ ~~~
    If UBound(checkRangeCols) = -1 Then
        compareColCount = rng.Areas(1).Columns.Count
        ReDim checkCols(1 To compareColCount)
        For tmpIdx = 1 To compareColCount
            checkCols(tmpIdx) = tmpIdx
        Next tmpIdx
    Else
        compareColCount = (UBound(checkRangeCols) - LBound(checkRangeCols)) + 1
        ReDim checkCols(1 To compareColCount)
        For tmpIdx = LBound(checkRangeCols) To UBound(checkRangeCols)
            checkCols(tmpIdx + 1) = checkRangeCols(tmpIdx)
        Next tmpIdx
    End If
    
    For areaIDX = 1 To rng.Areas.Count
        Dim rowIDX As Long, checkCol As Long, compareArr As Variant, curKey As String
        For rowIDX = 1 To rng.Areas(areaIDX).Rows.Count
            compareArr = GetCompareValues(rng.Areas(areaIDX), rowIDX, checkCols)
            curKey = Join(compareArr, ", ")
            If Not tmpDict.Exists(curKey) Then
                tmpDict(curKey) = rng.Rows(rowIDX).Row
            Else
                Dim keyFirstRow As Long
                keyFirstRow = CLng(tmpDict(curKey))
                'if it exists, then it's a duplicate
                If Not retDict.Exists(keyFirstRow) Then
                    'the first worksheet row with this values is Value from tmpDict
                    retDict(keyFirstRow) = 2
                Else
                    retDict(keyFirstRow) = CLng(retDict(keyFirstRow)) + 1
                End If
            End If
        Next rowIDX
    Next areaIDX
    
Finalize:
    If Not failed Then
        Set FindDuplicateRows = retDict
        
        'For Fun, List the Rows and How Many Duplicates Exist
       Dim dKey As Variant
       For Each dKey In retDict.Keys
            Debug.Print "Worksheet Row: " & dKey & ", has " & retDict(dKey) & " duplicates"
       Next dKey
        
    End If

    Exit Function
E:
    failed = True
    MsgBox "FindDuplicateRows failed. (Error: " & Err.Number & ", " & Err.Description & ")"
    Err.Clear
    Resume Finalize:

End Function

Private Function GetCompareValues(rngArea As Range, rngRow As Long, compCols() As Long) As Variant
    Dim valsArr As Variant
    Dim colcount As Long
    Dim idx As Long, curCol As Long, valCount As Long
    colcount = UBound(compCols) - LBound(compCols) + 1
    ReDim valsArr(1 To colcount)
    For idx = LBound(compCols) To UBound(compCols)
        valCount = valCount + 1
        curCol = compCols(idx)
        valsArr(valCount) = CStr(rngArea(rngRow, curCol).Value2)
    Next idx
    GetCompareValues = valsArr
End Function


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
Public Function GetA1CellRef(fromRng As Range, Optional colOffset As Long = 0, Optional rowCount As Long = 1, Optional colcount As Long = 1, Optional rowOffset As Long = 0, Optional fixedRef As Boolean = False, Optional visibleCellsOnly As Boolean = False) As String
'   return A1 style reference (e.g. "A10:A116") from selection
'   Optional offsets, resized ranges supported
    Dim tmpRng As Range
    Set tmpRng = fromRng.offset(rowOffset, colOffset)
    If colcount > 1 Or rowCount > 1 Then
        Set tmpRng = tmpRng.Resize(rowCount, colcount)
    End If
    If visibleCellsOnly Then
        Set tmpRng = tmpRng.SpecialCells(xlCellTypeVisible)
    End If
    GetA1CellRef = tmpRng.Address(fixedRef, fixedRef)
    Set tmpRng = Nothing
End Function
