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

Public Enum MergeRangeEnum
    mrDefault_MergeAll = 0
    mrUnprotect = 2 ^ 0
    mrClearFormatting = 2 ^ 1
    mrClearContents = 2 ^ 2
    mrMergeAcrossOnly = 2 ^ 3
End Enum

Public Type ftFound
    matchExactFirstIDX As Long
    matchExactLastIDX As Long
    matchSmallerIDX As Long
    matchLargerIDX As Long
    realRowFirst As Long
    realRowLast As Long
    realRowSmaller As Long
    realRowLarger As Long
End Type


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


'This will return RANGE object for value like 'MySheet!$C$4:$D$10'
'(Does not use Application 'Range' object, so workbook does not ave to be active workbook
Public Function RangeBySheetAdd(sheetBangAddress As String) As Range
On Error GoTo E:
    Dim failed As Boolean
    
    Dim shtName As String, rADd As String
    Dim posBang As Long
    posBang = InStr(1, sheetBangAddress, "!", vbTextCompare)
    'next line validate sheet name is at least 1 char, and address is at least 2 chars (e.g. "A1")
    If posBang > 1 And Len(sheetBangAddress) > posBang + 1 Then
        shtName = left(sheetBangAddress, posBang - 1)
        rADd = Mid(sheetBangAddress, posBang + 1)
        Set RangeBySheetAdd = ThisWorkbook.Worksheets(shtName).Range(rADd)
    End If
    
Finalize:
    On Error Resume Next
     
    If failed Then
        Set RangeBySheetAdd = Nothing
    End If
    If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:
    
End Function

'   RETURN TRUE IF RANGE CONTAINS 1 AREA AND ALL CELLS IN RANGE ARE MERGED
Public Function IsMerged(checkRng As Range) As Boolean
    If checkRng.Areas.Count > 1 Then
        IsMerged = False
    ElseIf IsNull(checkRng.MergeCells) Then
        IsMerged = False
    Else
        IsMerged = checkRng.MergeCells
    End If
End Function

Public Function MergeRange(mRng As Range, Optional options As MergeRangeEnum = MergeRangeEnum.mrDefault_MergeAll, _
    Optional vertAlign As XlVAlign, Optional horAlign As XlHAlign) As Boolean
On Error GoTo E:
    Dim failed As Boolean
    
    
    
    
    If mRng.Areas.Count <> 1 Then
        RaiseError ERR_INVALID_RANGE_SIZE, "pbRange.MergeRange 'mRng' argument allows for only 1 area! ('" & mRng.Worksheet.Name & "!" & mRng.Address & " contains * " & mRng.Areas.Count & " * areas!"
    End If
    
    If EnumCompare(options, MergeRangeEnum.mrUnprotect) Then
        If mRng.Worksheet.ProtectContents Then pbUnprotectSheet mRng.Worksheet
    End If
    If EnumCompare(options, MergeRangeEnum.mrClearContents + MergeRangeEnum.mrClearFormatting, ecAnd) Then
        mRng.Clear
    ElseIf EnumCompare(options, MergeRangeEnum.mrClearContents) Then
        mRng.ClearContents
    ElseIf EnumCompare(options, MergeRangeEnum.mrClearFormatting) Then
        mRng.ClearFormats
    End If
        
    If EnumCompare(options, MergeRangeEnum.mrMergeAcrossOnly) Then
        mRng.Merge Across:=True
    Else
        mRng.Merge Across:=False
    End If
    
    If Not IsMissing(vertAlign) Then mRng.VerticalAlignment = vertAlign
    If Not IsMissing(horAlign) Then mRng.HorizontalAlignment = horAlign

Finalize:
    On Error Resume Next
        
    MergeRange = Not failed
    
    If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:
    
End Function

Private Function NewFtFound() As ftFound
    Dim retV As ftFound
    NewFtFound = retV
End Function

Private Function ShouldFormatCriteria(crt As Variant) As Boolean
    Dim retV As Boolean
    Select Case TypeName(crt)
        Case "String"
            retV = False
        Case "Boolean"
            retV = False
        Case Else
            If IsNumeric(crt) Or IsDate(crt) Then
                retV = True
            Else
                retV = False
            End If
    End Select
End Function

Public Function HasAnyOverlappingValue(rng1 As Range, rng2 As Range) As Boolean

    If Not rng1 Is Nothing And Not rng2 Is Nothing Then
        HasAnyOverlappingValue = Not rng1.find(rng2.value) Is Nothing
    End If

End Function

Public Function CountUniqueInRange(rng As Range, Optional includeNonNumeric As Boolean = True) As Long

'    Dim tmpArr As Variant
'    tmpArr = ArrRange(rng, aoUnique)
'
'
    Dim cnt As Variant
    
    If rng.Areas.Count = 1 And rng.Areas(1).Count = 1 Then
        If includeNonNumeric = False And IsNumeric(rng.Cells(1, 1).value) Then
            cnt = 1
        Else
            If Len(rng.Cells(1, 1).value & vbNullString) > 0 Then
                cnt = 1
            End If
        End If
        GoTo Finalize:
        Exit Function
    End If
    
    On Error Resume Next
    If includeNonNumeric Then
        cnt = Application.WorksheetFunction.CountA(Application.WorksheetFunction.unique(rng))
    Else
        cnt = Application.WorksheetFunction.Count(Application.WorksheetFunction.unique(rng))
    End If
    If Err.Number <> 0 Then
        If IsDEV Then
            Beep
            MsgBox_FT "ERROR in pbRange.CoountUniqueInRange"
        End If
        Err.Clear
        cnt = -1
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

Finalize:
    CountUniqueInRange = cnt
    
End Function

Public Function CheckSort(lstObj As ListObject, col As Variant, sortPosition As Long, sortOrder As XlSortOrder) As Boolean

    Dim retV As Boolean
    Dim colcount As Long
    Dim sidx As Long
    Dim tmpIdx As Long
    If lstObj.Sort.SortFields.Count >= sortPosition Then
        retV = True
        Dim sortFld As SortField
        Set sortFld = lstObj.Sort.SortFields(sortPosition)
        If sortFld.key.Columns.Count <> 1 Then
            retV = False
            Exit Function
        End If
        If StrComp(sortFld.key.Address, lstObj.ListColumns(col).DataBodyRange.Address, vbTextCompare) <> 0 Then
            retV = False
            Exit Function
        End If
        If sortFld.Order <> sortOrder Then
            retV = False
            Exit Function
        End If
    End If

    CheckSort = retV

End Function

Public Function GetRangeMultipleCrit(lstObj As ListObject, Columns As Variant, crit As Variant) As Range
    
    
    
    Dim sortcols As Boolean
    Dim colcount As Long
    colcount = UBound(Columns) - LBound(Columns) + 1
    If lstObj.Sort.SortFields.Count < colcount Then sortcols = True
    Dim sidx As Long
    Dim tmpIdx As Long
    If sortcols = False Then
        For sidx = LBound(Columns) To UBound(Columns)
            tmpIdx = tmpIdx + 1
            If StrComp(lstObj.Sort.SortFields(tmpIdx).key.Address, lstObj.ListColumns(Columns(sidx)).DataBodyRange.Address, vbTextCompare) <> 0 Then
                sortcols = True
                Exit For
            End If
        Next sidx
    End If
    If sortcols Then
        For sidx = LBound(Columns) To UBound(Columns)
            If sidx = LBound(Columns) Then
                AddSort lstObj, Columns(sidx), xlAscending, True, True
            Else
                AddSort lstObj, Columns(sidx), xlAscending, False, False
            End If
        Next sidx
    End If

    Dim curRng As Range, currCol As Variant, currCrit As Variant
    
    Dim firstOuter As Long, lastOuter As Long
    
    Dim cIdx As Long
    For cIdx = LBound(Columns) To UBound(Columns)
        currCol = Columns(cIdx)
        currCrit = crit(cIdx)
        If ShouldFormatCriteria(currCrit) Then
            currCrit = Format(crit(cIdx), lstObj.ListColumns(currCol).DataBodyRange(1, 1).numberFormat)
        End If
        If cIdx = LBound(Columns) Then
            firstOuter = FirstRowInRange(lstObj.ListColumns(currCol).DataBodyRange, currCrit)
            lastOuter = LastRowInRange(lstObj.ListColumns(currCol).DataBodyRange, currCrit)
            If firstOuter = 0 Then
                Set curRng = Nothing
                Exit For
            End If
            Set curRng = lstObj.ListColumns(currCol).Range.offset(rowOffset:=firstOuter).Resize(rowSize:=lastOuter - firstOuter + 1)
        Else
            Set curRng = curRng.offset(ColumnOffset:=lstObj.ListColumns(currCol).Range.EntireColumn.column - curRng.column)
            firstOuter = FirstRowInRange(curRng, currCrit)
            lastOuter = LastRowInRange(curRng, currCrit)
            If firstOuter = 0 Then
                Set curRng = Nothing
                Exit For
            End If
            Dim rwOffset As Long
            If firstOuter > 1 Then
                rwOffset = firstOuter - 1
            Else
                rwOffset = 0
            End If
            Set curRng = curRng.offset(rowOffset:=rwOffset).Resize(rowSize:=lastOuter - firstOuter + 1)
        End If
    Next cIdx
         
    If Not curRng Is Nothing Then
        Dim lstRowIdxStart As Long
        lstRowIdxStart = curRng.Row - lstObj.HeaderRowRange.Row
        Set curRng = lstObj.listRows(lstRowIdxStart).Range.Resize(rowSize:=curRng.Rows.Count)
    End If

    

    If Not curRng Is Nothing Then
        Set GetRangeMultipleCrit = curRng
    End If
    
    

End Function

Public Function FirstRowInRange(rng As Range, crit As Variant) As Long
    
    If ShouldFormatCriteria(crit) Then
        crit = Format(crit, rng(1, 1).numberFormat)
    End If
    If TypeName(crit) = "String" Then
        If StrComp(rng(1, 1).value, crit, vbTextCompare) = 0 Then
            FirstRowInRange = 1
        Else
            FirstRowInRange = MatchFirst(crit, rng, ExactMatch)
        End If
    Else
        If rng(1, 1).value = crit Then
            FirstRowInRange = 1
        Else
            FirstRowInRange = MatchFirst(crit, rng, ExactMatch)
        End If
    End If

End Function

Public Function LastRowInRange(rng As Range, crit As Variant) As Long
    If ShouldFormatCriteria(crit) Then
        crit = Format(crit, rng(1, 1).numberFormat)
    End If
    If TypeName(crit) = "String" Then
        If StrComp(rng(rng.Rows.Count, 1).value, crit, vbTextCompare) = 0 Then
            LastRowInRange = rng.Rows.Count
        Else
            LastRowInRange = MatchLast(crit, rng, ExactMatch)
        End If
    Else
        If rng(rng.Rows.Count, 1).value = crit Then
            LastRowInRange = rng.Rows.Count
        Else
            LastRowInRange = MatchLast(crit, rng, ExactMatch)
        End If
    End If

End Function

Private Function FormatSearchCriteria(lstObj As ListObject, colArray As Variant, critArray As Variant) As Variant

    If lstObj.listRows.Count = 0 Then
        GoTo Finalize:
    End If
    
    'FORMAT CRIT
    Dim critIdx As Long
    For critIdx = LBound(critArray) To UBound(critArray)
        If TypeName(critArray(critIdx)) <> "Boolean" Then
            critArray(critIdx) = Format(critArray(critIdx), lstObj.ListColumns(colArray(critIdx)).DataBodyRange(1, 1).numberFormat)
        End If
    Next critIdx

Finalize:

    FormatSearchCriteria = critArray

End Function

Public Function FindFirstListRowMatchingMultCrit(lstObj As ListObject, Columns As Variant, crit As Variant) As Long
On Error GoTo E:

    If lstObj.listRows.Count = 0 Then Exit Function
    
    Dim startLooking As Range
    Dim lastCheckedRow As Long
    Dim lstObjRowIdx As Long
    
    crit = FormatSearchCriteria(lstObj, Columns, crit)
    
    Set startLooking = ListColumnRange(lstObj, Columns(1), includeHeaderRow:=True).find(crit(1), LookIn:=xlValues, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If Not startLooking Is Nothing Then
        Dim matched As Boolean
        Do
            lastCheckedRow = startLooking.Row
            lstObjRowIdx = lastCheckedRow - lstObj.HeaderRowRange.Row
            Dim colArrAI As ArrInformation
            colArrAI = ArrayInfo(Columns)
            If colArrAI.Rows > 0 Then
                Dim idx As Long
                For idx = LBound(Columns) To UBound(Columns)
                    If TypeName(crit(idx)) = "String" Then
                        matched = (StrComp(CStr(lstObj.ListColumns(Columns(idx)).DataBodyRange(RowIndex:=lstObjRowIdx).value), crit(idx), vbTextCompare) = 0)
                    Else
                        If TypeName(crit(idx)) = "Date" Then
                            matched = CLng(lstObj.ListColumns(Columns(idx)).DataBodyRange(RowIndex:=lstObjRowIdx).value) = CLng(crit(idx))
                        Else
                            If IsNumeric(crit(idx)) Then
                                matched = CDbl(lstObj.ListColumns(Columns(idx)).DataBodyRange(RowIndex:=lstObjRowIdx).value) = CDbl(crit(idx))
                            Else
                                matched = lstObj.ListColumns(Columns(idx)).DataBodyRange(RowIndex:=lstObjRowIdx).value = crit(idx)
                            End If
                        End If
                    End If
                    If matched = False Then Exit For
                Next idx
                If matched = True Then
                    'we matched everything
                    FindFirstListRowMatchingMultCrit = lstObjRowIdx
                    GoTo Finalize:
                   Exit Do
                End If
            Else
                FindFirstListRowMatchingMultCrit = lstObjRowIdx
                Exit Do
            End If
            
            Set startLooking = ListColumnRange(lstObj, Columns(1), includeHeaderRow:=True).FindNext(startLooking)
            If startLooking.Row <= lastCheckedRow Then
                Exit Do
            End If
        
        Loop While Not startLooking Is Nothing
    End If
    
Finalize:
    On Error Resume Next
    
    Set startLooking = Nothing
If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    ErrorCheck
    Resume Finalize:

End Function

Public Function GetOffsetRange(ByRef listObjCell As Range, returnCol As Variant) As Range

    'Check to make sure [listObjCell] is:
    '  - In a ListObject
    ' -  has only 1 area
    ' - has only 1 column
    Dim retRng As Range
    If listObjCell.ListObject Is Nothing Then
        Err.Raise 512 - 101, Description:="Range provided must be in the range of a ListObject"
    End If
    If listObjCell.Areas.Count > 1 Then
        Err.Raise 512 - 101, Description:="Range provided must contain only 1 area"
    End If
    If listObjCell.Columns.Count > 1 Then
        Err.Raise 512 - 101, Description:="Range provided must contain only 1 column"
    End If

    Dim firstLOCol As Long: firstLOCol = listObjCell.ListObject.ListColumns(1).Range.column
    Dim fromLOCol As Long: fromLOCol = (listObjCell.column - firstLOCol) + 1
    Dim toFldIdx As Long: toFldIdx = GetFieldIndex(listObjCell.ListObject, returnCol)
    If fromLOCol > 0 And toFldIdx > 0 Then
        Set retRng = listObjCell.offset(ColumnOffset:=(toFldIdx - fromLOCol))
    End If

    Set GetOffsetRange = retRng

End Function

Public Function CountInRange(srcRng As Range, criteria As Variant) As Long
On Error GoTo E:

    Dim cnt As Long
    If TypeName(criteria) = "Date" Then criteria = CDbl(criteria)
    cnt = Application.WorksheetFunction.CountIfs(srcRng, criteria)
    
    CountInRange = cnt
    Exit Function
E:
    Beeper
    Trace "Error in RangeMonger.CountInRange: Range: " & srcRng.Worksheet.Name & "." & srcRng.Address & ", for Criteria: " & CStr(criteria & "")

End Function

Public Function MinMaxNumber(lstObj As ListObject, field As Variant, operator As ftMinMax) As Variant

    If lstObj.listRows.Count = 0 Then Exit Function

    If Not IsNumeric(field) Then
        field = GetFieldIndex(lstObj, field)
    End If
    
    Dim retV As Variant
    Select Case operator
        Case ftMinMax.minValue
            retV = WorksheetFunction.Min(lstObj.ListColumns(field).DataBodyRange)
        Case ftMinMax.maxValue
            retV = WorksheetFunction.Max(lstObj.ListColumns(field).DataBodyRange)
    End Select
        
    If IsNumeric(retV) Then
        MinMaxNumber = retV
    End If

End Function



Public Function ReplaceBlanks(rng As Range, valueIfBlank As Variant)
On Error Resume Next
    
    If Not rng Is Nothing Then
        If Not rng.SpecialCells(xlCellTypeBlanks) Is Nothing Then
            rng.SpecialCells(xlCellTypeBlanks).value = valueIfBlank
        End If
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Function



'The reason behind this stupid magic is that values that are numbers and are ** PASTED ** into a range formatted as text, cannot be searched as text.
'Thanks Mr. Gates
Public Function MatchFirst(crit As Variant, ByRef rng As Range, matchmode As XMatchMode, Optional secondPass As Boolean = False, Optional srchMode As XSearchMode = searchFirstToLast) As Long
On Error GoTo E:
    Dim fnd As Variant
    If Len(crit & vbNullString) = 0 Then
        'We're looking for Blanks, use RangeSys("BlankCell")
        crit = RangeSYS(srBlankCell)
        'crit now equals EMPTY, which will pass numeric validation.  we don't want to look for a '0' value, so set secondPass = true
        secondPass = True
    End If
     
    If IsDate(rng(1, 1)) Then crit = CDbl(crit)
    
    fnd = WorksheetFunction.XMatch(crit, rng, matchmode, srchMode)
      
Finalize:
    On Error Resume Next

    If fnd = 0 And secondPass = False And IsNumeric(crit) Then
        fnd = MatchFirst(CDbl(crit), rng, ExactMatch, True, srchMode)
        Exit Function
    End If
      
    MatchFirst = fnd
   
     If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    Err.Clear
    fnd = 0
    Resume Finalize:
    
End Function


'The reason behind this stupid magic is that values that are numbers and are ** PASTED ** into a range formatted as text, cannot be searched as text.
'Thanks Mr. Gates
Public Function MatchLast(crit As Variant, ByRef rng As Range, matchmode As XMatchMode, Optional secondPass As Boolean = False, Optional srchMode As XSearchMode = searchLastToFirst) As Long
On Error GoTo E:
    Dim fnd As Variant
    If Len(crit & vbNullString) = 0 Then
        'We're looking for Blanks, use RangeSys("BlankCell")
        crit = RangeSYS(srBlankCell)
        'crit now equals EMPTY, which will pass numeric validation.  we don't want to look for a '0' value, so set secondPass = true
        secondPass = True
    End If
    
    If IsDate(rng(1, 1)) Then crit = CDbl(crit)

    
    fnd = WorksheetFunction.XMatch(crit, rng, matchmode, srchMode)
      
Finalize:
    On Error Resume Next

    If fnd = 0 And secondPass = False And IsNumeric(crit) Then
        fnd = MatchLast(CDbl(crit), rng, ExactMatch, True, srchMode)
    End If
     
     MatchLast = fnd
     If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    Err.Clear
    fnd = 0
    Resume Finalize:
    
End Function

Public Function DeleteListRowsRange_ShiftUp(lstObj As ListObject, startListRowIndex As Long, delRowCount As Long) As Boolean
On Error GoTo E:
    Dim failed As Boolean
    
   If lstObj.listRows.Count = 0 Then Exit Function
    
    If lstObj.Range.Worksheet.ProtectContents And lstObj.Range.Worksheet.Protection.AllowDeletingRows = False Then
        If pbUnprotectSheet(lstObj.Range.Worksheet) = False Then
            Err.Raise 419, Description:="Protected sheet does not allow deleting rows"
        End If
    End If
    
    If (startListRowIndex + delRowCount) - 1 > lstObj.listRows.Count Then
        Err.Raise 419, Description:="delete listRowStart + Row Count is > than total listrows"
    End If
    

    Dim delRng As Range
    Set delRng = lstObj.listRows(startListRowIndex).Range.Resize(rowSize:=delRowCount)
    
    delRng.Delete xlShiftUp
Finalize:
    On Error Resume Next
    
    DeleteListRowsRange_ShiftUp = Not failed
    
    If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:
End Function

Public Function DeleteListRows_UsesSort(listObj As ListObject, field As Variant, criteria As Variant, matchmode As XMatchMode) As Long
On Error GoTo E:
    Dim failed As Boolean

    Dim evts As Boolean, scrn As Boolean

   If listObj.listRows.Count = 0 Then Exit Function
    
    If listObj.Range.Worksheet.ProtectContents And listObj.Range.Worksheet.Protection.AllowDeletingRows = False Then
        If pbUnprotectSheet(listObj.Range.Worksheet) = False Then
            Err.Raise 419, Description:="Protected sheet does not allow deleting rows"
        End If
    End If
    
    
    Dim firstIDX As Long, lastIDX As Long, toDeleteCount As Long, fldIDX As Long
    fldIDX = GetFieldIndex(listObj, field)
    If GetFirstRowIndex(listObj, fldIDX, criteria, ExactMatch, False, False) > 0 Then
        If AddSort(listObj, fldIDX, xlAscending, True, True) Then
            firstIDX = GetFirstRowIndex(listObj, fldIDX, criteria, ExactMatch, False, False)
            lastIDX = GetLastRowIndex(listObj, fldIDX, criteria, ExactMatch, False, False)
            toDeleteCount = (lastIDX - firstIDX) + 1
            
            Dim delRng As Range
            Set delRng = listObj.listRows(firstIDX).Range
            Set delRng = delRng.Resize(rowSize:=toDeleteCount)
            delRng.Delete (xlShiftUp)
            Set delRng = Nothing
            
            Trace "Deleted " & toDeleteCount & " rows from " & listObj.Name & " where [" & field & "] = " & criteria
        Else
            failed = True
        End If
    End If
    
Finalize:
    On Error Resume Next
    
    If failed Then
        DeleteListRows_UsesSort = -1
    Else
        DeleteListRows_UsesSort = toDeleteCount
    End If
    
    If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:

End Function

'Remove All List Rows Where [field] matches [criteria]
'can specify full or partial match
'Throws error if Worksheet is Protected Against Deleting ListRows
'Returns Number of Deleted ListRows
Public Function DeleteFoundListRows(listObj As ListObject, field As Variant, criteria As Variant, matchmode As XMatchMode) As Long
   If listObj.listRows.Count = 0 Then Exit Function
   
    
    If listObj.Range.Worksheet.ProtectContents And listObj.Range.Worksheet.Protection.AllowDeletingRows = False Then
        If pbUnprotectSheet(listObj.Range.Worksheet) = False Then
            Err.Raise 419, Description:="Protected sheet does not allow deleting rows"
        End If
    End If
    
    
    ClearFilter listObj
    
    Dim fieldIdx As Integer, foundRow As Long
    fieldIdx = GetFieldIndex(listObj, field)
    
    'ensure all rows visible
    
    Dim sanityCount As Long, rCount As Long, nextRow As Long, deletedCount As Long
        
    With listObj
        rCount = .listRows.Count
        Do While True
            nextRow = 0
            If .listRows.Count = 0 Then
                Exit Do
            End If
            Dim fndRow As Long
            fndRow = MatchFirst(criteria, .ListColumns(fieldIdx).DataBodyRange, matchmode)
            If fndRow > 0 Then
                nextRow = fndRow
                If nextRow > 0 Then
                    .listRows(nextRow).Delete
                    deletedCount = deletedCount + 1
                End If
            Else
                Exit Do
            End If
            
            sanityCount = sanityCount + 1
            If sanityCount > rCount Then
                Exit Do
            End If
        Loop
    End With

    DeleteFoundListRows = deletedCount
    
    

End Function

Private Function GetFieldIndex(ByRef listObj As ListObject, field As Variant) As Long
    Dim fieldIdx As Integer
    If IsNumeric(field) Then
        fieldIdx = CLng(field)
    Else
        fieldIdx = listObj.ListColumns(field).Index
    End If
    GetFieldIndex = fieldIdx
End Function

'Get range of [field] - this will force sort the list object so a range with a single area can be returned
Public Function GetFoundRangeBetweenSortedRange(listObj As ListObject, field As Variant, greaterThanOrEqual As Variant, lessThanOrEqual As Variant) As Range

    '0 = exact match
    '-1 = exact or next smaller
    '1 = exact or next larger
    '2 = wildcard character match
    
    '1 = first to last
    '-1 = last to first
    
    If listObj.listRows.Count = 0 Then Exit Function



    AddSort listObj, field, xlAscending, True
    
    Dim foundRange As Range
    Dim fieldIdx As Long
    fieldIdx = GetFieldIndex(listObj, field)

    
    Dim firstRowIndex As Long, lastRowIndex As Long
    
    If TypeName(greaterThanOrEqual) = "Date" Then
        greaterThanOrEqual = CLng(greaterThanOrEqual)
        lessThanOrEqual = CLng(lessThanOrEqual)
    End If
    
    Dim rng As Range: Set rng = listObj.ListColumns(fieldIdx).DataBodyRange
    Dim countBlank As Long: countBlank = WorksheetFunction.countBlank(listObj.ListColumns(fieldIdx).DataBodyRange)
    If countBlank > 0 Then
        Set rng = listObj.ListColumns(fieldIdx).DataBodyRange.Resize(rowSize:=listObj.ListColumns(fieldIdx).DataBodyRange.Rows.Count - countBlank)
    End If
    
    On Error Resume Next
    firstRowIndex = WorksheetFunction.XMatch(greaterThanOrEqual, rng, 1, 1)
    lastRowIndex = WorksheetFunction.XMatch(lessThanOrEqual, rng, -1, -1)
    
    
    If firstRowIndex > 0 And lastRowIndex > 0 And lastRowIndex >= firstRowIndex Then
        Set foundRange = listObj.ListColumns(fieldIdx).DataBodyRange(RowIndex:=firstRowIndex).Resize(rowSize:=(lastRowIndex - firstRowIndex) + 1)
    End If

    Set GetFoundRangeBetweenSortedRange = foundRange
If Err.Number <> 0 Then Err.Clear
End Function

''This it 20 times SLOWER than the better way (see GetFoundSheetRowsArray)
'Public Function GetFoundSheetRowsArray2(ByVal srchRng As Range, criteria As Variant) As Long()
'On Error GoTo E:
'    Dim failed As Boolean
'
'    Dim dicItems As New Dictionary
'    Dim startLooking As Range
'    Dim foundRng As Range
'    Dim retv() As Long
'
'    Dim rInfo As RngInfo, ai as ArrInformation
'    Dim fieldIdx As Long
'    Dim rowIDX As Variant, realRow As Long, lastCheckedRow  As Long
'
'    rInfo = RangeInfo(srchRng)
'    If rInfo.Columns > 1 Then
'        RaiseError ERR_INVALID_RANGE_SIZE, "pbRange.GetFoundRangeRowsArray 'srchRng' must only contain 1 column"
'    End If
'
'    fieldIdx = 1
'    Set startLooking = srchRng
'    Set foundRng = startLooking.find(criteria, LookIn:=xlValues, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
'    Do While Not foundRng Is Nothing
'        lastCheckedRow = foundRng.Row
'        dicItems(lastCheckedRow) = 1
'
'        Set foundRng = startLooking.FindNext(foundRng)
'        If foundRng.Row <= lastCheckedRow Then
'            Exit Do
'        End If
'    Loop
'
'    If dicItems.Count > 0 Then
'        ReDim retv(1 To dicItems.Count, 1 To 1)
'        Dim dKey As Variant, cntr As Long
'        For Each dKey In dicItems.Keys
'            cntr = cntr + 1
'            retv(cntr, 1) = CLng(dKey)
'        Next dKey
'    End If
'
'Finalize:
'    On Error Resume Next
'    If Not failed Then
'        GetFoundSheetRowsArray2 = retv
'    End If
'    If ArrDimensions(retv) > 0 Then Erase retv
'    Set dicItems = Nothing
'    Set startLooking = Nothing
'    Set foundRng = Nothing
'
'    Exit Function
'E:
'    failed = True
'    ErrorCheck
'    Resume Finalize:
'
'End Function

Public Function GetFoundSheetRowsArray(ByVal srchRng As Range, criteria As Variant) As Long()
On Error GoTo E:
    Dim failed As Boolean
    
    Dim dicItems As New Dictionary
    Dim startLooking As Range
    Dim retV() As Long
    
    Dim rInfo As RngInfo, ai As ArrInformation
    Dim fieldIdx As Long
    Dim rowIDX As Variant, realRow As Long
    
    rInfo = RangeInfo(srchRng)
    If rInfo.Columns > 1 Then
        RaiseError ERR_INVALID_RANGE_SIZE, "pbRange.GetFoundRangeRowsArray 'srchRng' must only contain 1 column"
    End If
    
    fieldIdx = 1
    Set startLooking = srchRng
    rowIDX = MatchFirst(criteria, startLooking, ExactMatch)
    Do While rowIDX > 0
        realRow = startLooking(RowIndex:=rowIDX).Row
        dicItems(realRow) = 1
        If (startLooking.Rows.Count - rowIDX) = 0 Then
            Exit Do
        End If
        Set startLooking = startLooking.offset(rowIDX).Resize(startLooking.Rows.Count - rowIDX)
        rowIDX = MatchFirst(criteria, startLooking, ExactMatch)
    Loop
    
    If dicItems.Count > 0 Then
        ReDim retV(1 To dicItems.Count, 1 To 1)
        Dim dKey As Variant, cntr As Long
        For Each dKey In dicItems.Keys
            cntr = cntr + 1
            retV(cntr, 1) = CLng(dKey)
        Next dKey
    End If

Finalize:
    On Error Resume Next
    If Not failed Then
        GetFoundSheetRowsArray = retV
    End If
    If ArrDimensions(retV) > 0 Then Erase retV
    Set dicItems = Nothing
    Set startLooking = Nothing
    If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:
    
End Function

Public Function GetFoundListIndexArray(lstObj As ListObject, field As Variant, criteria As Variant) As Variant

'Will return Array(1 to 0) if no rows found
On Error GoTo E:

    Dim dicItems As New Dictionary
    If lstObj.listRows.Count = 0 Then Exit Function
    
    Dim fieldIdx As Long
    Dim startLooking As Range
    Dim rowIDX As Variant
    Dim lstObjRowIdx As Long
    
    fieldIdx = GetFieldIndex(lstObj, field)
    Set startLooking = lstObj.ListColumns(fieldIdx).DataBodyRange
    rowIDX = MatchFirst(criteria, startLooking, ExactMatch)
    Do While rowIDX > 0
        lstObjRowIdx = startLooking.Rows(rowIDX).Row - lstObj.HeaderRowRange.Row
        dicItems(lstObjRowIdx) = 1
        
        If (startLooking.Rows.Count - rowIDX) = 0 Then
            Exit Do
        End If
        
        Set startLooking = startLooking.offset(rowIDX).Resize(startLooking.Rows.Count - rowIDX)
        rowIDX = MatchFirst(criteria, startLooking, ExactMatch)
    Loop

Finalize:
    On Error Resume Next

    If dicItems.Count > 0 Then
        Dim k As Variant, retV() As Variant
        ReDim retV(1 To dicItems.Count)
        Dim cnt As Long
        cnt = 1
        For Each k In dicItems.Keys
            retV(cnt) = k
            cnt = cnt + 1
        Next k
        GetFoundListIndexArray = retV
    Else
        GetFoundListIndexArray = Array()
    End If

    Set dicItems = Nothing
    If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    ErrorCheck

End Function

'Return the range in a ListObject of all the matching values in a single ListColumn
Public Function GetFoundRange(listObj As ListObject, field As Variant, criteria As Variant, matchmode As XMatchMode, Optional returnColumn As Variant) As Range

    Dim foundRange As Range
    If TypeName(criteria) = "Date" Then criteria = CDbl(criteria)

    Dim fieldIdx As Integer, foundRow As Long
    fieldIdx = GetFieldIndex(listObj, field)

    Dim firstRow As Long, lastRow As Long
    firstRow = GetFirstRowIndex(listObj, fieldIdx, criteria, matchmode, True, True)
    If firstRow > 0 Then
        lastRow = GetLastRowIndex(listObj, fieldIdx, criteria, matchmode, True, True)
        Set foundRange = listObj.ListColumns(fieldIdx).DataBodyRange(RowIndex:=firstRow)
        Set foundRange = foundRange.Resize(rowSize:=(lastRow - firstRow) + 1)
    End If

    If Not foundRange Is Nothing Then
        If IsMissing(returnColumn) = False Then
            Dim retColIdx As Long
            retColIdx = GetFieldIndex(listObj, returnColumn)
            Dim offsetBy As Long
            offsetBy = retColIdx - fieldIdx
            Set foundRange = foundRange.offset(ColumnOffset:=offsetBy)
        End If
    End If

    Set GetFoundRange = foundRange

End Function

Public Function GetFirstRowIndex(ByRef listObj As ListObject, field As Variant, criteria As Variant, matchmode As XMatchMode, sortIfNeeded As Boolean, ClearFilter As Boolean) As Long
    GetFirstRowIndex = GetRowIndex(listObj, field, criteria, True, matchmode, sortIfNeeded, ClearFilter)
End Function
Public Function GetLastRowIndex(ByRef listObj As ListObject, field As Variant, criteria As Variant, matchmode As XMatchMode, sortIfNeeded As Boolean, ClearFilter As Boolean) As Long
    GetLastRowIndex = GetRowIndex(listObj, field, criteria, False, matchmode, sortIfNeeded, ClearFilter)
End Function



Public Function FindFirstInRange(ByVal rng As Range, criteria As Variant, Optional matchmode As XMatchMode = XMatchMode.ExactMatch) As Double
On Error GoTo E:
    Dim foundIDX As Double
    
    If TypeName(criteria) = "Date" Then criteria = CDbl(criteria)
 
    foundIDX = WorksheetFunction.XMatch(criteria, rng, matchmode, 1)
    
    FindFirstInRange = foundIDX
    
    Exit Function
E:
    'eat the error, return 0
    Err.Clear
    FindFirstInRange = 0
End Function

Public Function FindLastInRange(ByVal rng As Range, criteria As Variant, Optional matchmode As XMatchMode = XMatchMode.ExactMatch) As Double
On Error GoTo E:
    Dim foundIDX As Double
    
    If TypeName(criteria) = "Date" Then criteria = CDbl(criteria)
    
    foundIDX = WorksheetFunction.XMatch(criteria, rng, matchmode, -1)
    
    FindLastInRange = foundIDX
    
    Exit Function
E:
    'eat the error, return 0
    Err.Clear
    FindLastInRange = 0
End Function
Private Function GetRowIndex(ByRef listObj As ListObject, field As Variant, criteria As Variant, firstRowIndex As Boolean, matchmode As XMatchMode, sortIfNeeded As Boolean, clearFilters As Boolean) As Long
    
    
    
    Dim fieldIdx As Integer, foundRow As Long
    fieldIdx = GetFieldIndex(listObj, field)
    
    With listObj
        If .listRows.Count = 0 Then
            GetRowIndex = 0
            Exit Function
        End If
        
        If TypeName(criteria) = "Date" Then criteria = CDbl(criteria)
        If IsDate(listObj.ListColumns(field).DataBodyRange(1, 1)) Then criteria = CDbl(criteria)
        
        'before wasting time sorting and filtering, see if a matching value exists in the range
        If MatchFirst(criteria, .ListColumns(fieldIdx).DataBodyRange, matchmode) = 0 Then
            Exit Function
        End If
        
        If clearFilters Then
            ClearFilter listObj
        End If
        
        If sortIfNeeded Then
            If .Sort.SortFields.Count = 0 Then
                AddSort listObj, field, clearPreviousSorts:=True
            Else
                Dim srtField As SortField
                Set srtField = .Sort.SortFields(1)
                If srtField.key.column <> .ListColumns(fieldIdx).Range.column Then
                    AddSort listObj, field, clearPreviousSorts:=True
                End If
            End If
        End If
        
        If firstRowIndex = True Then
            GetRowIndex = MatchFirst(criteria, .ListColumns(fieldIdx).DataBodyRange, matchmode)
        Else
            GetRowIndex = MatchLast(criteria, .ListColumns(fieldIdx).DataBodyRange, matchmode)
        End If
        
        
    End With
    
    
End Function

Public Function GetRangeMultipleCriteria(lo As ListObject, Columns As Variant, criteria As Variant, returnColumn As Variant, Optional setRangeValue As Variant) As Range

    If UBound(Columns) <> UBound(criteria) Then Exit Function
    
    
    Dim idx As Long
    Dim retRange As Range
    
    'SORT ALL COLUMNS
    For idx = LBound(Columns) To UBound(Columns)
        If idx = LBound(Columns) Then
            AddSort lo, Columns(idx), xlAscending, True, True
        Else
            AddSort lo, Columns(idx), xlAscending, False, False
        End If
    Next idx
    
    Dim firstIDX As Long, lastIDX As Long
    For idx = LBound(Columns) To UBound(Columns)
        'If first, then use FindFirstRow/LastRow
        If idx = LBound(Columns) Then
            firstIDX = GetFirstRowIndex(lo, Columns(idx), criteria(idx), ExactMatch, False, False)
            lastIDX = GetLastRowIndex(lo, Columns(idx), criteria(idx), ExactMatch, False, False)
            If firstIDX = 0 Or lastIDX = 0 Then Exit Function
        Else
            'we're on at least  the 2nd pass
            Dim lookInRange As Range
            Set lookInRange = lo.ListColumns(Columns(idx)).DataBodyRange
            
            Set lookInRange = lookInRange.offset(rowOffset:=firstIDX - 1).Resize(rowSize:=lastIDX - firstIDX + 1)
            Dim subFirst As Long, subLast As Long
            subFirst = MatchFirst(criteria(idx), lookInRange, ExactMatch)
            subLast = MatchLast(criteria(idx), lookInRange, ExactMatch)
            If subFirst > 0 And subLast > 0 Then
                firstIDX = firstIDX + subFirst - 1
                lastIDX = firstIDX + (subLast - subFirst)
            Else
                firstIDX = 0
                lastIDX = 0
            End If
        End If
    Next idx
    
    If firstIDX > 0 And lastIDX > 0 Then
        Set retRange = lo.ListColumns(returnColumn).DataBodyRange.offset(rowOffset:=firstIDX - 1).Resize(lastIDX - firstIDX + 1)
        If Not IsMissing(setRangeValue) Then
            retRange.value = setRangeValue
        End If
    End If
    
    
    Set GetRangeMultipleCriteria = retRange
        
    


End Function


Public Function AddSortMultipleColumns(lstObj As ListObject, clearFilters As Boolean, sortOrder As XlSortOrder, ParamArray cols() As Variant) As Boolean
On Error GoTo E:

    If Not lstObj Is Nothing Then
        If lstObj.listRows.Count <= 1 Then
            AddSortMultipleColumns = True
            Exit Function
        End If
    End If

    If lstObj.Range.Worksheet.ProtectContents = True And lstObj.Range.Worksheet.Protection.AllowSorting = False Then
        If pbUnprotectSheet(lstObj.Range.Worksheet) = False Then
            Err.Raise 419, Description:="Protected sheet does not allow filtering"
        End If
    End If
    
    
    
    
    If clearFilters Then ClearFilter lstObj
    Dim colIDX As Long
    Dim needsSort As Boolean
    Dim tCols As Variant
    
    tCols = ArrParams(cols)
    
    For colIDX = LBound(tCols) To UBound(tCols)
        If CheckSort(lstObj, tCols(colIDX, 1), colIDX, sortOrder) = False Then
            needsSort = True
            Exit For
        End If
    Next colIDX
    
    If needsSort Then
    
        With lstObj.Sort
            .SortFields.Clear
            For colIDX = LBound(tCols) To UBound(tCols)
                .SortFields.add lstObj.ListColumns(tCols(colIDX, 1)).DataBodyRange, SortOn:=xlSortOnValues, Order:=sortOrder
            Next colIDX
            .Apply
        End With
    
    End If
   
   AddSortMultipleColumns = True
Finalize:
   
   Exit Function
E:
   AddSortMultipleColumns = False
   ErrorCheck
    Resume Finalize:
End Function

'Add Sort to ListObject, optionally clearing previous sorts
Public Function AddSort(listObj As ListObject, field As Variant, Optional Order As XlSortOrder = xlAscending, Optional clearPreviousSorts As Boolean = False, Optional clearFilters As Boolean = True) As Boolean
On Error GoTo E:

    
    If listObj.Range.Worksheet.ProtectContents = True And listObj.Range.Worksheet.Protection.AllowSorting = False Then
        If pbUnprotectSheet(listObj.Range.Worksheet) = False Then
            Err.Raise 419, Description:="Protected sheet does not allow filtering"
        End If
    End If
    


    Dim fieldIdx As Integer, RngInfo As String
    fieldIdx = GetFieldIndex(listObj, field)
    
    RngInfo = listObj.Name & "[" & listObj.ListColumns(fieldIdx).Name & "]"
    If listObj.listRows.Count > 0 Then
        If clearFilters Then
            ClearFilter listObj
        End If
        Dim sortAlreadyValid As Boolean
        sortAlreadyValid = True
        With listObj.Sort
            If clearPreviousSorts = True Then
                .SortFields.Clear
            Else
            End If
            .SortFields.add listObj.ListColumns(field).DataBodyRange, SortOn:=xlSortOnValues, Order:=Order
            .Apply
        End With
    End If
   AddSort = True
Finalize:
   
    
   Exit Function
E:
   AddSort = False
    ErrorCheck
    Resume Finalize:
    
        
End Function

Private Function ConvertConditionOperatorToString(cndOper As XlFormatConditionOperator) As String

    Dim retVal As String

    Select Case cndOper
        Case XlFormatConditionOperator.xlEqual
            retVal = "="
        Case XlFormatConditionOperator.xlGreater
            retVal = ">"
        Case XlFormatConditionOperator.xlGreaterEqual
            retVal = ">="
        Case XlFormatConditionOperator.xlLess
            retVal = "<"
        Case XlFormatConditionOperator.xlLessEqual
            retVal = "<="
        Case XlFormatConditionOperator.xlNotEqual
            retVal = "<>"
        Case Else
            retVal = ""
    End Select

    ConvertConditionOperatorToString = retVal

End Function

Public Function AddFilterPartialMatch(listObj As ListObject, field As Variant, crit1 As Variant, Optional clearExistFilters As Boolean = False) As Long
On Error GoTo E:
    

    Dim cnt As Long
    Dim fieldIdx As Long
    fieldIdx = GetFieldIndex(listObj, field)

    If clearExistFilters = False And ColumnFiltered(listObj, fieldIdx) = True Then
        
        Err.Raise 5, Description:="A filter is already applied on " & listObj.Name & "." & listObj.ListColumns(field).Name
    End If

    If Not listObj Is Nothing And listObj.listRows.Count > 0 Then
        If clearExistFilters Then
            ClearFilter listObj
        End If
        Dim handled As Boolean
        With listObj
            If Len(crit1) > 0 Then
                If Strings.left(CStr(crit1), 1) <> "*" Then crit1 = "*" & crit1
                If Strings.Right(CStr(crit1), 1) <> "*" Then crit1 = crit1 & "*"
            End If
            .Range.AutoFilter field:=fieldIdx, Criteria1:=crit1, operator:=xlFilterValues
            AddFilterPartialMatch = WorksheetFunction.Subtotal(3, .ListColumns(fieldIdx).DataBodyRange)
        End With
    End If

    
    Exit Function
E:
    ErrorCheck

End Function


Public Function AddFilter_MatchesAnyCrit(listObj As ListObject, field As Variant, crit1 As Variant, Optional clearExistFilters As Boolean = False) As Long
On Error GoTo E:
    

    Dim cnt As Long
    Dim fieldIdx As Long
    fieldIdx = GetFieldIndex(listObj, field)
    
    If clearExistFilters = False And ColumnFiltered(listObj, fieldIdx) = True Then
        
        Err.Raise 5, Description:="A filter is already applied on " & listObj.Name & "." & listObj.ListColumns(field).Name
    End If
    
    If Not listObj Is Nothing And listObj.listRows.Count > 0 Then
        If clearExistFilters Then
            ClearFilter listObj
        End If
        Dim handled As Boolean
        With listObj
            .Range.AutoFilter field:=fieldIdx, Criteria1:=crit1, operator:=xlFilterValues
            AddFilter_MatchesAnyCrit = WorksheetFunction.Subtotal(3, .ListColumns(fieldIdx).DataBodyRange)
        End With
    End If
    
    
    Exit Function
E:
    ErrorCheck

End Function




'Filters for exact match, Returns count of filtered rows
Public Function AddFilterSimple(listObj As ListObject, field As Variant, conditionOperator As XlFormatConditionOperator, crit1 As Variant, Optional clearExistFilters As Boolean = False) As Long
On Error GoTo E:
    

    Dim cnt As Long
    Dim fieldIdx As Long
    fieldIdx = GetFieldIndex(listObj, field)
    
    Dim cndOper As String: cndOper = ConvertConditionOperatorToString(conditionOperator)
    
    If listObj.Range.Worksheet.ProtectContents And listObj.Range.Worksheet.Protection.AllowFiltering = False Then
        If pbUnprotectSheet(listObj.Range.Worksheet) = False Then
            
            Err.Raise 419, Description:="Protected sheet does not allow filtering"
        End If
    End If
    
    If clearExistFilters = False And ColumnFiltered(listObj, fieldIdx) = True Then
        
        Err.Raise 5, Description:="A filter is already applied on " & listObj.Name & "." & listObj.ListColumns(field).Name
    End If
    
    If Not listObj Is Nothing And listObj.listRows.Count > 0 Then
        If clearExistFilters Then
            ClearFilter listObj
        End If
        Dim handled As Boolean
        With listObj
            If TypeName(crit1) = "Date" Then
                handled = True
                If conditionOperator = xlEqual Then
                    If Len(crit1) = 0 Then
                        .Range.AutoFilter field:=fieldIdx, Criteria1:="="
                    Else
                        .Range.AutoFilter field:=fieldIdx, Criteria1:=">=" & Int(crit1), operator:=xlAnd, Criteria2:="<" & Int(crit1) + 1
                    End If
                Else
                    .Range.AutoFilter field:=fieldIdx, Criteria1:=cndOper & Int(crit1)
                End If
            End If
            If handled = False And TypeName(crit1) = "String" Then
                handled = True
                If Len(crit1) = 0 Then
                    .Range.AutoFilter field:=fieldIdx, Criteria1:="="
                Else
                    .Range.AutoFilter field:=fieldIdx, Criteria1:=cndOper & crit1
                End If
            End If
            If handled = False And TypeName(crit1) = "Boolean" Then
                handled = True
                .Range.AutoFilter field:=fieldIdx, Criteria1:=cndOper & crit1
            End If
            If handled = False And IsNumeric(crit1) Then
                handled = True
                crit1 = Format(crit1, .ListColumns(fieldIdx).DataBodyRange.numberFormat)
                .Range.AutoFilter field:=fieldIdx, Criteria1:=cndOper & crit1
            End If
            'If We're filtering for blank, find the first visible ListColumn, then count visible range
            If Len(crit1) = 0 Then
                Dim findVisIdx As Long
                For findVisIdx = 1 To .ListColumns.Count
                    If .ListColumns(findVisIdx).Range.EntireColumn.Hidden = False Then
                        Dim vRng As Range
                        Set vRng = GetVisible(.ListColumns(findVisIdx).DataBodyRange)
                        If Not vRng Is Nothing Then
                            cnt = vRng.Count
                        End If
                        Exit For
                    End If
                Next findVisIdx
            Else
                'This will work for cells that have a value of any data type
                cnt = WorksheetFunction.Subtotal(3, .ListColumns(fieldIdx).DataBodyRange)
            End If
        End With
    End If

    AddFilterSimple = cnt
    
    Exit Function
E:
    Beeper
    Trace Err.Source & ", " & Err.Description
    Err.Clear
End Function

Public Function ColumnFiltered(listObj As ListObject, col As Variant) As Boolean
    Dim colIDX As Long
    colIDX = GetFieldIndex(listObj, col)
    If FilterCount(listObj) = 0 Then
        ColumnFiltered = False
    Else
        If Not listObj.AutoFilter.Filters Is Nothing Then
            ColumnFiltered = listObj.AutoFilter.Filters(colIDX).On
        End If
    End If
End Function

Public Function AddFilterBetween(listObj As ListObject, field As Variant, crit1 As Variant, crit2 As Variant, Optional clearExistFilters As Boolean = False) As Long
On Error GoTo E:
    
    Dim cnt As Long
    Dim fieldIdx As Long
    fieldIdx = GetFieldIndex(listObj, field)
    
    If listObj.Range.Worksheet.ProtectContents And listObj.Range.Worksheet.Protection.AllowFiltering = False Then
        If pbUnprotectSheet(listObj.Range.Worksheet) = False Then
            
            Err.Raise 419, listObj, "Protected sheet does not allow filtering"
        End If
    End If
    
    If clearExistFilters = False And ColumnFiltered(listObj, fieldIdx) = True Then
        
        Err.Raise 5, Description:="A filter is already applied on " & listObj.Name & "." & listObj.ListColumns(field).Name
    End If

    
    If Not listObj Is Nothing And listObj.listRows.Count > 0 Then
        If clearExistFilters Then
            ClearFilter listObj
        Else
            'If we're not clearing filters, and a filter is already applied to [field], then throw invalid error
            
        End If
        Dim handled As Boolean
        With listObj
            If TypeName(crit1) = "Date" And TypeName(crit2) = "Date" Then
                handled = True
                .Range.AutoFilter field:=fieldIdx, Criteria1:=">=" & Int(crit1), operator:=xlAnd, Criteria2:="<=" & Int(crit2)
            End If
            If handled = False And TypeName(crit1) = "String" Then
                handled = True
                .Range.AutoFilter field:=fieldIdx, Criteria1:=">=" & crit1, operator:=xlAnd, Criteria2:="<=" & crit2
            End If
            If handled = False And IsNumeric(crit1) Then
                handled = True
                .Range.AutoFilter field:=fieldIdx, Criteria1:=">=" & crit1, operator:=xlAnd, Criteria2:="<=" & crit2
            End If
            cnt = WorksheetFunction.Subtotal(3, .ListColumns(fieldIdx).DataBodyRange)
        End With
    End If

    AddFilterBetween = cnt
    
    Exit Function
E:
    Beeper
    Trace Err.Source & ", " & Err.Description
    Err.Clear
End Function


Public Function ListObjectIsFiltered(listObj As ListObject) As Boolean
On Error Resume Next
    Dim isFiltered As Boolean
    If Not listObj.AutoFilter Is Nothing Then
        If listObj.AutoFilter.FilterMode Then
            isFiltered = True
        End If
        If isFiltered = False Then
            If listObj.Range.Worksheet.FilterMode = True Then
                isFiltered = True
            End If
        End If
    End If

    ListObjectIsFiltered = isFiltered
    If Err.Number <> 0 Then Err.Clear
End Function




Public Function ClearFilter(listObj As ListObject) As Boolean
On Error GoTo E:
    
    
    Dim failed As Boolean
    
    If listObj Is Nothing Then
        GoTo Finalize:
    End If
    If listObj.listRows.Count = 0 Then
        GoTo Finalize:
    End If
    
    If listObj.Range.Worksheet.ProtectContents And listObj.Range.Worksheet.Protection.AllowFiltering = False Then
        If pbUnprotectSheet(listObj.Range.Worksheet) = False Then
            
            Err.Raise 419, listObj, "Protected sheet does not allow filtering"
        End If
    End If
    
    If Not listObj.AutoFilter Is Nothing Then
        listObj.AutoFilter.ShowAllData
    Else
        If listObj.Range.Worksheet.FilterMode = True Then
            listObj.Range.Worksheet.ShowAllData
        End If
    End If
    

    
    
Finalize:
    On Error Resume Next

    ClearFilter = Not failed
    
    If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
   failed = True
   Err.Raise Err.Number, Err.Source, Err.Description
   Resume Finalize:
    
End Function



'Get count of enabled filters
Private Function FilterCount(lstObj As ListObject) As Long
    Dim fltrCount As Long
    If lstObj.Range.Worksheet.FilterMode = False Then
        fltrCount = 0
    Else
        With lstObj
            If Not .AutoFilter.Filters Is Nothing Then
                Dim lcIDX As Long
                For lcIDX = 1 To .ListColumns.Count
                    If .AutoFilter.Filters(lcIDX).On Then
                        fltrCount = fltrCount + 1
                    End If
                Next lcIDX
            End If
        End With
    End If
    
    FilterCount = fltrCount
End Function

'Returns true if all items in [rng] are in the listObject in validcolIdx
Public Function RangeIsInsideListColumn(rng As Range, lstObj As ListObject, validColIdx As Long) As Boolean

    If rng Is Nothing Or lstObj Is Nothing Or validColIdx = 0 Then
        RangeIsInsideListColumn = False
        Exit Function
    End If
    If rng.Count = 1 Then
        If Intersect(rng, lstObj.ListColumns(validColIdx).DataBodyRange) Is Nothing Then
            RangeIsInsideListColumn = False
            Exit Function
        Else
            RangeIsInsideListColumn = True
            Exit Function
        End If
    End If
    
    If rng.Areas.Count = 1 Then
        If Intersect(rng, lstObj.ListColumns(validColIdx).DataBodyRange) Is Nothing Then
            RangeIsInsideListColumn = False
            Exit Function
        Else
            'check columns
            If rng.Columns.Count <> 1 Then
                RangeIsInsideListColumn = False
                Exit Function
            End If
            RangeIsInsideListColumn = True
            Exit Function
        End If
    Else
        Dim rngArea As Range
        For Each rngArea In rng.Areas
            If RangeIsInsideListColumn(rngArea, lstObj, validColIdx) = False Then
                RangeIsInsideListColumn = False
                Exit Function
            End If
        Next rngArea
    End If
    
    RangeIsInsideListColumn = True

End Function

Public Function FirstVisibleListRowIdx(lstObj As ListObject, Optional goBeyondListObjectIfNoVisible As Boolean = True) As Long

    If lstObj.listRows.Count > 0 Then
        Dim firstVisColIdx As Long: firstVisColIdx = FirstVisibleListColIndex(lstObj)
        Dim rowIDX As Long
        If firstVisColIdx > 0 Then
            Dim vRng As Range
            Set vRng = lstObj.ListColumns(firstVisColIdx).Range.SpecialCells(xlCellTypeVisible)
            Dim areaRng As Range
            For Each areaRng In vRng.Areas
                If areaRng.Row > lstObj.HeaderRowRange.Row Then
                    rowIDX = areaRng.Row
                    Exit For
                Else
                    If areaRng.Rows.Count > 1 Then
                        rowIDX = areaRng.Rows(2).Row
                        Exit For
                    End If
                End If
            Next areaRng
        End If
        If goBeyondListObjectIfNoVisible Then
            If rowIDX = 0 Then
                rowIDX = lstObj.HeaderRowRange.Row + lstObj.listRows.Count + 1
                If lstObj.ShowTotals Then
                    rowIDX = rowIDX + 1
                End If
            End If
        End If
        If rowIDX > 0 Then
            FirstVisibleListRowIdx = rowIDX - lstObj.HeaderRowRange.Row
        End If
    End If

End Function

Public Function FirstVisibleListColIndex(lstObj As ListObject) As Long

    If lstObj Is Nothing Then Exit Function
    
    Dim vRng As Range
    Set vRng = GetVisible(lstObj.HeaderRowRange)
    If Not vRng Is Nothing Then
        FirstVisibleListColIndex = (vRng.column - lstObj.HeaderRowRange.column) + 1
    End If
    
End Function

'SpecialCells(xlCellTypeVisible) throws an err if there is nothing visible in the range, so use this
Public Function GetVisible(srcRng As Range) As Range
On Error GoTo nothingVisible:
    Dim visRng As Range
    Set visRng = srcRng.SpecialCells(xlCellTypeVisible)
    
    If Not visRng Is Nothing Then
        Set GetVisible = visRng
    End If
    
    Exit Function
nothingVisible:
    Err.Clear
    Set visRng = Nothing
    Set GetVisible = Nothing
End Function

Public Function VisibleListObjRows(lstObj As ListObject) As Long
On Error GoTo E:

    If lstObj Is Nothing Then
        Exit Function
    End If

    Trace "Getting Visible List Obj Row Count for: " & lstObj.Name


    Dim rwCount As Long, visColIdx As Long
    visColIdx = FirstVisibleListColIndex(lstObj)
    If visColIdx > 0 Then
        If lstObj.listRows.Count > 0 Then
            Dim visRng As Range
            Set visRng = lstObj.listRows(1).Range(1, visColIdx)
            Set visRng = visRng.Resize(rowSize:=lstObj.listRows.Count)
            
            If visRng.Count = 1 And visRng.EntireRow.Hidden = False Then
                rwCount = 1
            Else
                Set visRng = GetVisible(visRng)
                If Not visRng Is Nothing Then
                    rwCount = visRng.Count
                End If
            End If
        End If
    End If
    
    VisibleListObjRows = rwCount
    
    Exit Function
E:
    Beeper
    Trace "Error getting VisibleListObjRows for " & lstObj.Name, True
    Trace "Error: - " & Err.Number & ", " & Err.Source & ", " & Err.Description
    Err.Clear
    
End Function

Public Function FindInColRange(ByRef rng As Range, searchVal As Variant, Optional isSortedAsc As Boolean = False) As ftFound
    If rng.Columns.Count > 1 Then
        Err.Raise ERR_INVALID_RANGE_SIZE, Description:="Invalid Range Size: Column Count <> 1"
    End If
    
    Dim retV As ftFound
    
    If rng Is Nothing Then
        FindInColRange = retV
        Exit Function
    End If
    
    If isSortedAsc Then
        retV.matchExactFirstIDX = MatchFirst(searchVal, rng, ExactMatch, srchMode:=searchBinaryAsc)
    Else
        retV.matchExactFirstIDX = MatchFirst(searchVal, rng, ExactMatch)
    End If
    
    If retV.matchExactFirstIDX > 0 Then
        retV.matchExactLastIDX = MatchLast(searchVal, rng, ExactMatch)
        retV.realRowFirst = rng(RowIndex:=retV.matchExactFirstIDX).Row
        retV.realRowLast = rng(RowIndex:=retV.matchExactLastIDX).Row
    Else
        If isSortedAsc Then
            retV.matchSmallerIDX = MatchFirst(searchVal, rng, ExactMatchOrNextSmaller, srchMode:=searchBinaryAsc)
            retV.matchLargerIDX = MatchFirst(searchVal, rng, ExactMatchOrNextLarger, srchMode:=searchBinaryAsc)
        Else
            retV.matchSmallerIDX = MatchFirst(searchVal, rng, ExactMatchOrNextSmaller)
            retV.matchLargerIDX = MatchFirst(searchVal, rng, ExactMatchOrNextLarger)
        End If
        If retV.matchSmallerIDX > 0 Then
            retV.realRowSmaller = rng(RowIndex:=retV.matchSmallerIDX).Row
        End If
        If retV.matchLargerIDX > 0 Then
            retV.realRowLarger = rng(RowIndex:=retV.matchLargerIDX).Row
        End If
    End If
    
    FindInColRange = retV

End Function


'All cols in colArray will be sorted in order ascending
'If all values are matched in range, the matching position will be returned
'If all values are not found in range, the matching position will be the
'position to insert which will maintain the sort orders
'-1 returned if Error
Public Function InsertPosition(lo As ListObject, colArray As Variant, searchValues As Variant, Optional assumeSorted As Boolean = True) As Long
On Error GoTo E:
    
    
    Dim failed As Boolean

    Dim retV As Long
    Dim curColIdx As Long, curRng As Range, tmpValue As Variant
    Dim arrCurRng As Variant
    Dim tmpFound As ftFound
    
    Dim loCount As Long
    loCount = lo.listRows.Count

    If lo.listRows.Count = 0 Then
        InsertPosition = 1
        GoTo Finalize:
    End If
'    If lo.listRows.Count = 1 Then
'        'freaking one off exception
'        retV = InsertPositionHasOneRow(lo, colArray, searchValues)
'        GoTo Finalize:
'    End If

    If assumeSorted = False Then
        AddSortMultipleColumns lo, True, xlAscending, colArray
    End If
    
    For curColIdx = LBound(colArray) To UBound(colArray)
        Dim loCol As Long: loCol = colArray(curColIdx)
        If curColIdx = LBound(colArray) Then
            Set curRng = Nothing
            tmpFound = NewFtFound
            Set curRng = lo.ListColumns(loCol).DataBodyRange
        Else
            Set curRng = lo.ListColumns(loCol).DataBodyRange(RowIndex:=tmpFound.matchExactFirstIDX)
            If (tmpFound.matchExactLastIDX - tmpFound.matchExactFirstIDX) + 1 > 1 Then
                Set curRng = curRng.Resize(rowSize:=(tmpFound.matchExactLastIDX - tmpFound.matchExactFirstIDX) + 1)
            End If
        End If
        
        tmpFound = FindInColRange(curRng, searchValues(curColIdx), isSortedAsc:=True)
        
        If tmpFound.matchExactFirstIDX <= 0 Then
            If tmpFound.matchLargerIDX > 0 Then
                retV = tmpFound.realRowLarger - lo.HeaderRowRange.Row
                GoTo Finalize:
            Else
                If tmpFound.matchSmallerIDX > 0 Then
                    retV = (tmpFound.realRowSmaller - lo.HeaderRowRange.Row) + 1
                    GoTo Finalize:
                End If
            End If
        Else
            'have exact match(es)
            If curColIdx = UBound(colArray) Then
                retV = tmpFound.realRowFirst - lo.HeaderRowRange.Row
                GoTo Finalize:
            End If
        End If
    Next curColIdx
    
    
Finalize:
    On Error Resume Next
    If failed Then
        retV = -1
    End If
    
    InsertPosition = retV
    
    Set curRng = Nothing
    Erase arrCurRng
    
    
    If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:
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


Public Function RangeRowCount(ByVal rng As Range) As Long

    Dim tmpCount As Long
    Dim rowDict As Dictionary
    Dim rCount As Long, areaIDX As Long, rwIDX As Long
    
    If rng Is Nothing Then
        GoTo Finalize:
    End If
    
    'Check first if all First/Count are the same, if they are, no need to loop through everything
    If AreasMatchRows(rng) Then
        tmpCount = rng.Areas(1).Rows.Count
    Else
        Set rowDict = New Dictionary
        For areaIDX = 1 To rng.Areas.Count
            For rwIDX = 1 To rng.Areas(areaIDX).Rows.Count
                Dim realRow As Long
                realRow = rng.Areas(areaIDX).Rows(rwIDX).Row
                rowDict(realRow) = realRow
            Next rwIDX
        Next areaIDX
        tmpCount = rowDict.Count
    End If

Finalize:
    RangeRowCount = tmpCount
    Set rowDict = Nothing

End Function

'returns 0 if any area has different numbers of columns than another
Public Function RangeColCount(ByVal rng As Range) As Long

    Dim tmpCount As Long
    Dim colDict As Dictionary
    Dim firstCol As Long, areaIDX As Long, colIDX As Long
    
    If rng Is Nothing Then
        GoTo Finalize:
    End If
    
    
    If AreasMatchCols(rng) Then
        tmpCount = rng.Areas(1).Columns.Count
    Else
        Set colDict = New Dictionary
        For areaIDX = 1 To rng.Areas.Count
            firstCol = rng.Areas(areaIDX).column
            colDict(firstCol) = firstCol
            For colIDX = 1 To rng.Areas(areaIDX).Columns.Count
                If colIDX > 1 Then colDict(firstCol + (colIDX - 1)) = firstCol + (colIDX - 1)
            Next colIDX
        Next areaIDX
        tmpCount = colDict.Count
    End If
    
Finalize:
    RangeColCount = tmpCount
    Set colDict = Nothing

End Function

Private Function AreasMatchRows(rng As Range) As Boolean

    If rng Is Nothing Then
        AreasMatchRows = True
        Exit Function
    End If

    Dim retV As Boolean
    If rng.Areas.Count = 1 Then
        retV = True
    Else
        Dim firstRow As Long, firstCount As Long, noMatch As Boolean, aIDX As Long
        firstRow = rng.Areas(1).Row
        firstCount = rng.Areas(1).Rows.Count
        For aIDX = 2 To rng.Areas.Count
            With rng.Areas(aIDX)
                If .Row <> firstRow Or .Rows.Count <> firstCount Then
                    noMatch = True
                    Exit For
                End If
            End With
        Next aIDX
        retV = Not noMatch
    End If

    AreasMatchRows = retV

End Function

Private Function AreasMatchCols(rng As Range) As Boolean

    Dim retV As Boolean
    If rng.Areas.Count = 1 Then
        retV = True
    Else
        Dim firstCol As Long, firstCount As Long, noMatch As Boolean, aIDX As Long
        firstCol = rng.Areas(1).column
        firstCount = rng.Areas(1).Columns.Count
        For aIDX = 2 To rng.Areas.Count
            With rng.Areas(aIDX)
                If .column <> firstCol Or .Columns.Count <> firstCount Then
                    noMatch = True
                    Exit For
                End If
            End With
        Next aIDX
        retV = Not noMatch
    End If

    AreasMatchCols = retV

End Function


Public Function ContiguousRows(rng As Range) As Boolean
'RETURNS TRUE IF HAS 1 AREA OR ALL AREAS SHARE SAME FIRST/LAST ROW
    Dim retV As Boolean

    If rng Is Nothing Then
        ContiguousRows = True
        Exit Function
    End If

    If rng.Areas.Count = 1 Then
        retV = True
    Else
        'If any Area is outside the min/max row of any other area then return false
        Dim loop1 As Long, loop2 As Long, isDiffRange As Boolean
        Dim l1Start As Long, l1End As Long, l2Start As Long, l2End As Long
        
        For loop1 = 1 To rng.Areas.Count
            l1Start = rng.Areas(loop1).Row
            l1End = l1Start + rng.Areas(loop1).Rows.Count - 1
            
            For loop2 = 1 To rng.Areas.Count
                l2Start = rng.Areas(loop2).Row
                l2End = l1Start + rng.Areas(loop2).Rows.Count - 1
                If l1Start < l2Start Or l1End > l2End Then
                    isDiffRange = True
                End If
                If isDiffRange Then Exit For
            Next loop2
            If isDiffRange Then Exit For
        Next loop1
    End If

    retV = Not isDiffRange

    ContiguousRows = retV

End Function

Public Function ContiguousColumns(rng As Range) As Boolean

'RETURNS TRUE IF HAS 1 AREA OR ALL AREAS SHARE SAME FIRST/LAST COLUMN
    Dim retV As Boolean

    If rng Is Nothing Then
        ContiguousColumns = True
        Exit Function
    End If


    If rng.Areas.Count = 1 Then
        retV = True
    Else
        'If any Area is outside the min/max row of any other area then return false
        Dim loop1 As Long, loop2 As Long, isDiffRange As Boolean
        Dim l1Start As Long, l1End As Long, l2Start As Long, l2End As Long
        
        For loop1 = 1 To rng.Areas.Count
            l1Start = rng.Areas(loop1).column
            l1End = l1Start + rng.Areas(loop1).Columns.Count - 1
            
            For loop2 = 1 To rng.Areas.Count
                l2Start = rng.Areas(loop2).column
                l2End = l1Start + rng.Areas(loop2).Columns.Count - 1
                If l1Start < l2Start Or l1End > l2End Then
                    isDiffRange = True
                End If
                If isDiffRange Then Exit For
            Next loop2
            If isDiffRange Then Exit For
        Next loop1
    End If

    retV = Not isDiffRange

    ContiguousColumns = retV

End Function


