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

Public Function IsSorted(rng As Range) As Boolean
    If rng.rows.Count > 1 Then
        Dim rng1 As Range, rng2 As Range
        Set rng1 = rng.Resize(RowSize:=rng.rows.Count - 1)
        Set rng2 = rng1.Offset(rowOffset:=1)
        Dim expr As String
        expr = "AND(" & "'[" & ThisWorkbook.Name & "]" & rng1.Worksheet.Name & "'!" & rng1.Address & "<='[" & ThisWorkbook.Name & "]" & rng2.Worksheet.Name & "'!" & rng2.Address & ")"
        Debug.Print expr
        '=AND('[RangeFindBenchmark.xlsm]TableA'!$C$2:$C$12<='[RangeFindBenchmark.xlsm]TableA'!$C$3:$C$13)
        IsSorted = Evaluate(expr)
    Else
        IsSorted = True
    End If
End Function


Public Function IsValueUnique(searchRng As Range, crit As Variant) As Boolean
    On Error Resume Next
    IsValueUnique = WorksheetFunction.CountIfs(searchRng, crit) = 1
End Function

Public Function IsListColSorted(lstObj As ListObject, lstCol As Variant) As Boolean
    If lstObj.listRows.Count <= 1 Then
        IsListColSorted = True
    Else
        IsListColSorted = IsSorted(lstObj.ListColumns(lstCol).DataBodyRange)
    End If
End Function

    ''   * GIVEN A KEY, AND VALUE, LOOK FOR ALL MATCHING
    ''   * KEYS IN [TARGETRANGE].[TARGETKEYCOL]
    ''   * Returns Count of ROWS WITH MATCHING KEY AND
    ''          MISMATCHED VALUE ([TargetRange].[targetRefCol] <> [srcRefVal] )
    ''   * Optionally, If 'updateInvalid' = True, then
    ''          mismatched values will be changed to equal [srcRefVal], and
    ''          (Return count then equal number of items changed)
    Public Function ReferenceMisMatch( _
        srcKey As Variant, _
        srcRefVal As Variant, _
        targetRange As Range, _
        targetKeyCol As Long, _
        targetRefCol As Long, _
        Optional updateInvalid As Boolean = False) As Long
        
    On Error GoTo E:
        
        Dim failed  As Boolean, evts As Boolean
        Dim mismatchCount  As Long
        Dim keyRng As Range, valRng As Range
        evts = Application.EnableEvents
        Application.EnableEvents = False
        
        Set keyRng = targetRange(1, targetKeyCol).Resize(RowSize:=targetRange.rows.Count)
        Set valRng = targetRange(1, targetRefCol).Resize(RowSize:=targetRange.rows.Count)
        Dim changedValues As Boolean
        Dim keyARR() As Variant, valARR() As Variant
        
        If targetRange.rows.Count = 1 Then
            ReDim keyARR(1 To 1, 1 To 1)
            ReDim valARR(1 To 1, 1 To 1)
            keyARR(1, 1) = keyRng(1, 1)
            valARR(1, 1) = valRng(1, 1)
        Else
            keyARR = keyRng
            valARR = valRng
        End If
        
        Dim rowIdx As Long, curInvalid As Boolean
        For rowIdx = LBound(keyARR) To UBound(keyARR)
            curInvalid = False
            If TypeName(srcKey) = "String" Then
                If StringsMatch(srcKey, keyARR(rowIdx, 1), smEqual) Then
                   If StringsMatch(srcRefVal, valARR(rowIdx, 1), smEqual) = False Then curInvalid = True
                End If
            ElseIf srcKey = keyARR(rowIdx, 1) Then
                If srcRefVal <> valARR(rowIdx, 1) Then curInvalid = True
            End If
            If curInvalid Then
                mismatchCount = mismatchCount + 1
                If updateInvalid Then valARR(rowIdx, 1) = srcRefVal
            End If
        Next rowIdx
    
        If mismatchCount > 0 And updateInvalid Then
            valRng = valARR
        End If
        
        ReferenceMisMatch = mismatchCount
    
Finalize:
        On Error Resume Next
            
            If failed Then
                'optional handling
            End If
            Application.EnableEvents = evts
            
        Exit Function
E:
        failed = True
        Debug.Print Err.number, Err.Description
        ErrorCheck
        Resume Finalize:
    
    End Function

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
    Dim areaIdx As Long
    If rng.Areas.Count >= 1 Then
        firstCol = rng.Areas(1).column
        totCols = rng.Areas(1).Columns.Count
        For areaIdx = 1 To rng.Areas.Count
            If Not rng.Areas(areaIdx).column = firstCol _
                Or Not rng.Areas(areaIdx).Columns.Count = totCols Then
                Err.Raise 17, Description:="FindDuplicateRows can not support mismatched columns for multiple Range Areas"
            End If
        Next areaIdx
    End If
    
    Dim retDict As New Dictionary, tmpDict As New Dictionary, compareColCount As Long, tmpIdx As Long
    Dim checkCols() As Long
    retDict.CompareMode = TextCompare
    tmpDict.CompareMode = TextCompare
    
    If rng.Areas.Count = 1 And rng.rows.Count = 1 Then
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
    
    For areaIdx = 1 To rng.Areas.Count
        Dim rowIdx As Long, checkCol As Long, compareArr As Variant, curKey As String
        For rowIdx = 1 To rng.Areas(areaIdx).rows.Count
            compareArr = GetCompareValues(rng.Areas(areaIdx), rowIdx, checkCols)
            curKey = Join(compareArr, ", ")
            If Not tmpDict.Exists(curKey) Then
                tmpDict(curKey) = rng.rows(rowIdx).Row
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
        Next rowIdx
    Next areaIdx
    
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
    MsgBox "FindDuplicateRows failed. (Error: " & Err.number & ", " & Err.Description & ")"
    Err.Clear
    Resume Finalize:

End Function

Private Function GetCompareValues(rngArea As Range, rngRow As Long, compCols() As Long) As Variant
    Dim valsArr As Variant
    Dim colCount As Long
    Dim IDX As Long, curCol As Long, valCount As Long
    colCount = UBound(compCols) - LBound(compCols) + 1
    ReDim valsArr(1 To colCount)
    For IDX = LBound(compCols) To UBound(compCols)
        valCount = valCount + 1
        curCol = compCols(IDX)
        valsArr(valCount) = CStr(rngArea(rngRow, curCol).Value2)
    Next IDX
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
        shtName = Left(sheetBangAddress, posBang - 1)
        rADd = Mid(sheetBangAddress, posBang + 1)
        Set RangeBySheetAdd = ThisWorkbook.Worksheets(shtName).Range(rADd)
    End If
    
Finalize:
    On Error Resume Next
     
    If failed Then
        Set RangeBySheetAdd = Nothing
    End If
    If Err.number <> 0 Then Err.Clear
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:
    
End Function

Public Function TestM()

    Dim rng As Range
    Set rng = activeSheet.Range("M5:N5")
    Debug.Assert IsRangeMerged(rng) = ebTRUE + ebPartial

End Function

'   RETURNS ebTRUE if any cells are merged, ebTrue+ebPartial is some cells are merged
Public Function IsRangeMerged(checkRng As Range) As ExtendedBool
    Dim extBool As ExtendedBool
    If IsNull(checkRng.MergeCells) Then
        extBool = ebTRUE + ebPartial
    Else
        If checkRng.MergeCells = True Then
            extBool = ebTRUE
        Else
            extBool = ebFALSE
        End If
    End If
    IsRangeMerged = extBool
End Function

Public Function MergeRange(mRng As Range, Optional options As MergeRangeEnum = MergeRangeEnum.mrDefault_MergeAll, _
    Optional vertAlign As XlVAlign, Optional horAlign As XlHAlign) As Boolean
    On Error Resume Next
    Dim failed As Boolean
    
    
    
    
    If mRng.Areas.Count <> 1 Then
        RaiseError ERR_INVALID_RANGE_SIZE, "pbRange.MergeRange 'mRng' argument allows for only 1 area! ('" & mRng.Worksheet.Name & "!" & mRng.Address & " contains * " & mRng.Areas.Count & " * areas!"
    End If
    
    If EnumCompare(options, MergeRangeEnum.mrUnprotect) Then
        If mRng.Worksheet.protectContents Then UnprotectSheet mRng.Worksheet
    End If
    If EnumCompare(options, MergeRangeEnum.mrClearContents + MergeRangeEnum.mrClearFormatting, ecAnd) Then
        mRng.Clear
    ElseIf EnumCompare(options, MergeRangeEnum.mrClearContents) Then
        mRng.ClearContents
    ElseIf EnumCompare(options, MergeRangeEnum.mrClearFormatting) Then
        mRng.ClearFormats
    End If
        
    If IsRangeMerged(mRng) = ebFALSE Then
        If EnumCompare(options, MergeRangeEnum.mrMergeAcrossOnly) Then
            mRng.Merge Across:=True
        Else
            mRng.Merge
        End If
    End If
    
    If Not IsMissing(vertAlign) Then mRng.VerticalAlignment = vertAlign
    If Not IsMissing(horAlign) Then mRng.HorizontalAlignment = horAlign

Finalize:
    On Error Resume Next
        
    MergeRange = Not failed
    
    If Err.number <> 0 Then Err.Clear
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
        HasAnyOverlappingValue = Not rng1.Find(rng2.value) Is Nothing
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
    If Err.number <> 0 Then
        If IsDev Then
            Beep
            MsgBox_FT "ERROR in pbRange.CoountUniqueInRange"
        End If
        Err.Clear
        cnt = -1
    End If
    If Err.number <> 0 Then Err.Clear
    On Error GoTo 0

Finalize:
    CountUniqueInRange = cnt
    
End Function

' This could return the wrong result
Public Function CheckSort(lstObj As ListObject, col As Variant, sortPosition As Long, sortOrder As XlSortOrder) As Boolean
    Dim retV As Boolean
    Dim colCount As Long
    Dim sidx As Long
    Dim tmpIdx As Long
    If lstObj.Sort.SortFields.Count >= sortPosition Then
        retV = True
        Dim sortFld As SortField
        Set sortFld = lstObj.Sort.SortFields(sortPosition)
        If sortFld.Key.Columns.Count <> 1 Then
            retV = False
            Exit Function
        End If
        If StrComp(sortFld.Key.Address, lstObj.ListColumns(col).DataBodyRange.Address, vbTextCompare) <> 0 Then
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
    Dim colCount As Long
    colCount = UBound(Columns) - LBound(Columns) + 1
    If lstObj.Sort.SortFields.Count < colCount Then sortcols = True
    Dim sidx As Long
    Dim tmpIdx As Long
    If sortcols = False Then
        For sidx = LBound(Columns) To UBound(Columns)
            tmpIdx = tmpIdx + 1
            If StrComp(lstObj.Sort.SortFields(tmpIdx).Key.Address, lstObj.ListColumns(Columns(sidx)).DataBodyRange.Address, vbTextCompare) <> 0 Then
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
            Set curRng = lstObj.ListColumns(currCol).Range.Offset(rowOffset:=firstOuter).Resize(RowSize:=lastOuter - firstOuter + 1)
        Else
            Set curRng = curRng.Offset(ColumnOffset:=lstObj.ListColumns(currCol).Range.EntireColumn.column - curRng.column)
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
            Set curRng = curRng.Offset(rowOffset:=rwOffset).Resize(RowSize:=lastOuter - firstOuter + 1)
        End If
    Next cIdx
         
    If Not curRng Is Nothing Then
        Dim lstRowIdxStart As Long
        lstRowIdxStart = curRng.Row - lstObj.HeaderRowRange.Row
        Set curRng = lstObj.listRows(lstRowIdxStart).Range.Resize(RowSize:=curRng.rows.Count)
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
            FirstRowInRange = MatchFirst(crit, rng, exactMatch)
        End If
    Else
        If rng(1, 1).value = crit Then
            FirstRowInRange = 1
        Else
            FirstRowInRange = MatchFirst(crit, rng, exactMatch)
        End If
    End If

End Function

Public Function LastRowInRange(rng As Range, crit As Variant) As Long
    If ShouldFormatCriteria(crit) Then
        crit = Format(crit, rng(1, 1).numberFormat)
    End If
    If TypeName(crit) = "String" Then
        If StrComp(rng(rng.rows.Count, 1).value, crit, vbTextCompare) = 0 Then
            LastRowInRange = rng.rows.Count
        Else
            LastRowInRange = MatchLast(crit, rng, exactMatch)
        End If
    Else
        If rng(rng.rows.Count, 1).value = crit Then
            LastRowInRange = rng.rows.Count
        Else
            LastRowInRange = MatchLast(crit, rng, exactMatch)
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
        If TypeName(critArray(critIdx)) <> "Boolean" And lstObj.ListColumns(colArray(critIdx)).DataBodyRange(1, 1).numberFormat <> "General" Then
            critArray(critIdx) = Format(critArray(critIdx), lstObj.ListColumns(colArray(critIdx)).DataBodyRange(1, 1).numberFormat)
        End If
    Next critIdx

Finalize:

    FormatSearchCriteria = critArray

End Function

Public Function FindFirstListObjectRow(lstObj As ListObject, Columns As Variant, crit As Variant, Optional RebuildDict As Boolean = False) As Long
    FindFirstListObjectRow = pbListObj.getFirstRow(lstObj, Columns, crit, forceRebuild:=RebuildDict)
End Function

'   ~~~ ~~~ Very Fast Function to find the first row of a ListObject where EXACT MATCH filters can be applied for up to ALL the columns in the list object
'   ~~~ ~~~ Recommend at a minimum the List Object be sorted Ascending by the First Column in the [Columns] Array
'   ~~~ ~~~
'   ~~~ ~~~ Arguments:
'   ~~~ ~~~ [lstObj] = Reference to list object being searched
'   ~~~ ~~~ [Columns] = An Array of ListColumn Names or ListColumn Indexes
'   ~~~ ~~~ [Crit] = An Array of Search Criteria - 1 criteria for each ListObject Column in the [Columns] Array
'   ~~~ ~~~ Example: lstObjRowIndex = FindFirstListObjectRow([myListObject], Array("LastName","DOB"), Array("Smith",CDate("12/28/80"))
'Public Function FindFirstListObjectRow(lstObj As ListObject, Columns As Variant, crit As Variant) As Long
'On Error GoTo E:
'
'
'    Dim failed As Boolean
'
'    '   If no rows, no play
'    If lstObj.listRows.Count = 0 Then Exit Function
'
'    '   Get reference to worksheet. Cutting out the 'middle man (list object)' for the Range.Find calls, to save even a few clock ticks
'    Dim ws As Worksheet
'    Set ws = lstObj.Range.Worksheet
'
'    Dim matchedListObjIdx As Long, matchedWSRow As Long
'    Dim firstRow As Long, lastRow As Long, rowOffset As Long, colOffset As Long
'
'    '   get worksheet first/last row for possible ListObject range that can be searched
'    firstRow = lstObj.listRows(1).Range.Row
'    lastRow = lstObj.HeaderRowRange.Row + lstObj.listRows.Count
'
'    '   Since we're searching columns based on ListObject Column Index, and since were returning the ListObject RowIndex if found,
'    '   get the offset of the ListObject to the Worksheet, so we can search the right worksheet columns, and return the right ListObject row
'     colOffset = lstObj.ListColumns(1).Range.column - 1
'     rowOffset = lstObj.listRows(1).Range.Row - 1
'
'    '   this reformats the search criteria so it can find results based on the Range.NumberFormat of the list columns.
'    '   you may want to tweak for your own purposes as this will allow you to find, for example "$100.50" even though the actual value
'    '   might be something like 100.5012
'    Dim critIdx As Long
'    For critIdx = LBound(crit) To UBound(crit)
'        If TypeName(crit(critIdx)) <> "Boolean" And TypeName(crit(critIdx)) <> "String" And lstObj.ListColumns(Columns(critIdx)).DataBodyRange(1, 1).numberFormat <> "General" Then
'            'crit(critIdx) = Format(crit(critIdx), lstObj.ListColumns(Columns(critIdx)).DataBodyRange(1, 1).numberFormat)
'        End If
'    Next critIdx
'
'    Dim startLooking As Range
'    Dim lastCheckedRow As Long, colsArrIDX As Long, matched As Boolean, finalMatch As Boolean
'    Dim evalCol As Long, evalRow As Long
'    Dim arrLB As Long, arrUB As Long
'    arrLB = LBound(Columns)
'    arrUB = UBound(Columns)
'
'    '   Search for the First matched filter for the first Column
'
''    Set startLooking = Nothing
''    DoEvents
'
'
'    Set startLooking = lstObj.ListColumns(Columns(arrLB)).Range.find(crit(arrLB), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
'    If startLooking Is Nothing Then
'        GoTo Finalize:
'    End If
'    '   If only 1 column is being search, we're done
'    If arrUB - arrLB = 0 Then
'        matchedWSRow = startLooking.Row
'        finalMatch = True
'        GoTo Finalize:
''        lastCheckedRow = startLooking.Row
'    Else
'        '   Matched first column, but need to match addtional columns
''        matchedWSRow = startLooking.Row
'        lastCheckedRow = startLooking.Row
'    End If
'
'    '   Go through the remaining columns (we already matched the first) to see if the row matches
'    Do While lastCheckedRow <= lastRow
'        finalMatch = False
'        For colsArrIDX = arrLB + 1 To arrUB
'            evalCol = lstObj.ListColumns(Columns(colsArrIDX)).Index + colOffset
'            If IsDEV Then Debug.Print "Looking For " & lstObj.ListColumns(Columns(colsArrIDX)).Name & " = " & crit(colsArrIDX)
'            If IsDEV Then Debug.Print " -- Value to Compare Against Is:  " & ws.Cells(lastCheckedRow, evalCol)
'            evalRow = lastCheckedRow
'            matched = False
'            If TypeName(crit(colsArrIDX)) = "String" Then
'                matched = StrComp(ws.Cells(evalRow, evalCol).Text, CStr(crit(colsArrIDX)), vbTextCompare) = 0
'            Else
'                If TypeName(crit(colsArrIDX)) = "Date" Then
'                    matched = CLng(ws.Cells(evalRow, evalCol).value) = CLng(CDate(crit(colsArrIDX)))
'                Else
'                    If IsNumeric(crit(colsArrIDX)) Then
'                        matched = CDbl(ws.Cells(evalRow, evalCol).value) = CDbl(crit(colsArrIDX))
'                    Else
'                        matched = ws.Cells(evalRow, evalCol) = crit(colsArrIDX)
'                    End If
'                End If
'            End If
'            If IsDEV Then Debug.Print "Matched: " & matched & " on col: " & colsArrIDX & " of " & UBound(crit)
'
'            If colsArrIDX = arrUB And matched Then
'                    matchedWSRow = lastCheckedRow
'                    finalMatch = True
'                '   positive row match
'                    GoTo Finalize:
'            End If
'            If matched = False Then
'                Exit For
'            End If
'        Next colsArrIDX
'        'lastCheckedRow was Not a match, look for the next row to check
'        If lastCheckedRow < lastRow Then
'            Dim adjRange As Range
'            If startLooking.Row >= lastRow Then GoTo Finalize:
'            Set adjRange = ws.Cells(startLooking.Row, lstObj.ListColumns(Columns(arrLB)).Range.column)
'            Set adjRange = adjRange.Resize(rowSize:=(lastRow - startLooking.Row + 1))
'            Set startLooking = adjRange.find(crit(arrLB), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
'            Set adjRange = Nothing
'            DoEvents
'            If startLooking Is Nothing Then GoTo Finalize:
'            If startLooking.Row <= lastCheckedRow Then GoTo Finalize:
'            lastCheckedRow = startLooking.Row
'        Else
'            Exit Do
'        End If
'    Loop
'
'
'Finalize:
'    On Error Resume Next
'    Set ws = Nothing
'    Set startLooking = Nothing
'    If failed Then
'        matchedListObjIdx = 0
'    ElseIf finalMatch And matchedWSRow > 0 Then
'        '   Adjust result to reflect the ListObjectRowIndex
'        matchedListObjIdx = matchedWSRow - rowOffset
'    End If
'    FindFirstListObjectRow = matchedListObjIdx
'
'    Exit Function
'E:
'    failed = True
'    'MsgBox "(Implement your own error handling) An error occured in FindFirstListRowMultCriteria: " & Err.Number & ", " & Err.Description
'    ErrorCheck
'    Resume Finalize:
'
'End Function





Public Function FindFirstListRowMatchingMultCrit(lstObj As ListObject, Columns As Variant, crit As Variant) As Long
    FindFirstListRowMatchingMultCrit = FindFirstListObjectRow(lstObj, Columns, crit)
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
        Set retRng = listObjCell.Offset(ColumnOffset:=(toFldIdx - fromLOCol))
    End If

    Set GetOffsetRange = retRng

End Function

Public Function CountBlankInListCol(lstObj As ListObject, field As Variant) As Long
    If lstObj.listRows.Count > 0 Then
        CountBlankInListCol = CountBlankInRange(lstObj.ListColumns(field).DataBodyRange)
    End If
End Function

Public Function ReplaceBlankInListCol(lstObj As ListObject, field As Variant, replaceWith As Variant)
    ReplaceBlankInRange lstObj.ListColumns(field).DataBodyRange, replaceWith
End Function

Public Function ReplaceBlankInRange(srcRng As Range, replaceWith As Variant)
On Error GoTo E:
    If srcRng Is Nothing Then Exit Function
    
    
    Dim evts As Boolean
    evts = Events
    EventsOff
    
    If srcRng.Worksheet.protectContents Then
        'reapply protect
        ProtectSheet srcRng.Worksheet
    End If

    If CountBlankInRange(srcRng) > 0 Then
        srcRng.Replace "", replaceWith, LookAt:=xlWhole
    End If

Finalize:
    On Error Resume Next
    Events = evts

    Exit Function
E:
    ErrorCheck
    
End Function

 Public Function CountBlankInRange(srcRange As Range) As Variant
 On Error Resume Next
        If Not srcRange Is Nothing Then
            CountBlankInRange = WorksheetFunction.CountBlank(srcRange)
        End If
    If Not Err.number = 0 Then Err.Clear
 End Function
 
Public Function CountInRange(srcRng As Range, criteria As Variant) As Long
On Error GoTo E:

    Dim cnt As Long
    If TypeName(criteria) = "Date" Then criteria = CDbl(criteria)
    cnt = Application.WorksheetFunction.CountIfs(srcRng, criteria)
    
    CountInRange = cnt
    Exit Function
E:
    ftBeep btError
    LogERROR "pbRange.CountInRange - Error in RangeMonger.CountInRange: Range: " & srcRng.Worksheet.Name & "." & srcRng.Address & ", for Criteria: " & CStr(criteria & "")

End Function

Public Function MinMaxNumber(lstObj As ListObject, field As Variant, operator As MinMax) As Variant

    If lstObj.listRows.Count = 0 Then Exit Function

    If Not IsNumeric(field) Then
        field = GetFieldIndex(lstObj, field)
    End If
    
    Dim retV As Variant
    Select Case operator
        Case MinMax.minValue
            retV = WorksheetFunction.Min(lstObj.ListColumns(field).DataBodyRange)
        Case MinMax.maxValue
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
    
    If Err.number <> 0 Then Err.Clear
End Function



'The reason behind this stupid magic is that values that are numbers and are ** PASTED ** into a range formatted as text, cannot be searched as text.
'Thanks Mr. Gates
Public Function MatchFirst(crit As Variant, ByRef rng As Range, matchMode As XMatchMode, Optional secondPass As Boolean = False, Optional srchMode As XSearchMode = searchFirstToLast) As Long
On Error GoTo E:
    Dim fnd As Variant
    If Len(crit & vbNullString) = 0 Then
        crit = vbEmpty
        'crit now equals EMPTY, which will pass numeric validation.  we don't want to look for a '0' value, so set secondPass = true
        secondPass = True
    End If
     
    If IsDate(rng(1, 1)) Then crit = CDbl(crit)
    
    fnd = WorksheetFunction.XMatch(crit, rng, matchMode, srchMode)
      
Finalize:
    On Error Resume Next

    If fnd = 0 And secondPass = False And IsNumeric(crit) Then
        fnd = MatchFirst(CDbl(crit), rng, exactMatch, True, srchMode)
        Exit Function
    End If
      
    MatchFirst = fnd
   
     If Err.number <> 0 Then Err.Clear
    Exit Function
E:
    Err.Clear
    fnd = 0
    Resume Finalize:
    
End Function


'The reason behind this stupid magic is that values that are numbers and are ** PASTED ** into a range formatted as text, cannot be searched as text.
'Thanks Mr. Gates
Public Function MatchLast(crit As Variant, ByRef rng As Range, matchMode As XMatchMode, Optional secondPass As Boolean = False, Optional srchMode As XSearchMode = searchLastToFirst) As Long
On Error GoTo E:
    Dim fnd As Variant
    If Len(crit & vbNullString) = 0 Then
        'We're looking for Blanks, use RangeSys("BlankCell")
        crit = vbEmpty
        'crit now equals EMPTY, which will pass numeric validation.  we don't want to look for a '0' value, so set secondPass = true
        secondPass = True
    End If
    
    If IsDate(rng(1, 1)) Then crit = CDbl(crit)

    
    fnd = WorksheetFunction.XMatch(crit, rng, matchMode, srchMode)
      
Finalize:
    On Error Resume Next

    If fnd = 0 And secondPass = False And IsNumeric(crit) Then
        fnd = MatchLast(CDbl(crit), rng, exactMatch, True, srchMode)
    End If
     
     MatchLast = fnd
     If Err.number <> 0 Then Err.Clear
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
    
    If lstObj.Range.Worksheet.protectContents And lstObj.Range.Worksheet.protection.allowDeletingRows = False Then
        UnprotectSheet lstObj.Range.Worksheet
    End If
    
    If (startListRowIndex + delRowCount) - 1 > lstObj.listRows.Count Then
        Err.Raise 419, Description:="delete listRowStart + Row Count is > than total listrows"
    End If
    

    Dim delRng As Range
    Set delRng = lstObj.listRows(startListRowIndex).Range.Resize(RowSize:=delRowCount)
    
    delRng.Delete xlShiftUp
Finalize:
    On Error Resume Next
    
    DeleteListRowsRange_ShiftUp = Not failed
    
    If Err.number <> 0 Then Err.Clear
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:
End Function

Public Function DeleteListRows_UsesSort(listObj As ListObject, field As Variant, criteria As Variant, matchMode As XMatchMode) As Long
On Error GoTo E:
    Dim failed As Boolean

    Dim evts As Boolean, scrn As Boolean

   If listObj.listRows.Count = 0 Then Exit Function
    
    If listObj.Range.Worksheet.protectContents And listObj.Range.Worksheet.protection.allowDeletingRows = False Then
        UnprotectSheet listObj.Range.Worksheet
    End If
    
    
    Dim firstIdx As Long, lastIdx As Long, toDeleteCount As Long, fldIDX As Long
    fldIDX = GetFieldIndex(listObj, field)
    If GetFirstRowIndex(listObj, fldIDX, criteria, exactMatch, False, False) > 0 Then
        If AddSort(listObj, fldIDX, xlAscending, True, True) Then
            firstIdx = GetFirstRowIndex(listObj, fldIDX, criteria, exactMatch, False, False)
            lastIdx = GetLastRowIndex(listObj, fldIDX, criteria, exactMatch, False, False)
            toDeleteCount = (lastIdx - firstIdx) + 1
            
            Dim delRng As Range
            Set delRng = listObj.listRows(firstIdx).Range
            Set delRng = delRng.Resize(RowSize:=toDeleteCount)
            delRng.Delete (xlShiftUp)
            Set delRng = Nothing
            
            LogTRACE "Deleted " & toDeleteCount & " rows from " & listObj.Name & " where [" & field & "] = " & criteria
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
    
    If Err.number <> 0 Then Err.Clear
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
Public Function DeleteFoundListRows(listObj As ListObject, field As Variant, criteria As Variant, matchMode As XMatchMode) As Long
   If listObj.listRows.Count = 0 Then Exit Function
   
    
    If listObj.Range.Worksheet.protectContents And listObj.Range.Worksheet.protection.allowDeletingRows = False Then
        UnprotectSheet listObj.Range.Worksheet
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
            fndRow = MatchFirst(criteria, .ListColumns(fieldIdx).DataBodyRange, matchMode)
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
    Dim CountBlank As Long: CountBlank = WorksheetFunction.CountBlank(listObj.ListColumns(fieldIdx).DataBodyRange)
    If CountBlank > 0 Then
        Set rng = listObj.ListColumns(fieldIdx).DataBodyRange.Resize(RowSize:=listObj.ListColumns(fieldIdx).DataBodyRange.rows.Count - CountBlank)
    End If
    
    On Error Resume Next
    firstRowIndex = WorksheetFunction.XMatch(greaterThanOrEqual, rng, 1, 1)
    lastRowIndex = WorksheetFunction.XMatch(lessThanOrEqual, rng, -1, -1)
    
    
    If firstRowIndex > 0 And lastRowIndex > 0 And lastRowIndex >= firstRowIndex Then
        Set foundRange = listObj.ListColumns(fieldIdx).DataBodyRange(RowIndex:=firstRowIndex).Resize(RowSize:=(lastRowIndex - firstRowIndex) + 1)
    End If

    Set GetFoundRangeBetweenSortedRange = foundRange
If Err.number <> 0 Then Err.Clear
End Function



Public Function GetFoundSheetRowsArray(ByVal srchRng As Range, criteria As Variant) As Long()
On Error GoTo E:
    Dim failed As Boolean
    
    Dim dicItems As New Dictionary
    Dim startLooking As Range
    Dim retV() As Long
    
    Dim rInfo As RngInfo, ai As ArrInformation
    Dim fieldIdx As Long
    Dim rowIdx As Variant, realRow As Long
    
    rInfo = RangeInfo(srchRng)
    If rInfo.Columns > 1 Then
        RaiseError ERR_INVALID_RANGE_SIZE, "pbRange.GetFoundRangeRowsArray 'srchRng' must only contain 1 column"
    End If
    
    fieldIdx = 1
    Set startLooking = srchRng
    rowIdx = MatchFirst(criteria, startLooking, exactMatch)
    Do While rowIdx > 0
        realRow = startLooking(RowIndex:=rowIdx).Row
        dicItems(realRow) = 1
        If (startLooking.rows.Count - rowIdx) = 0 Then
            Exit Do
        End If
        Set startLooking = startLooking.Offset(rowIdx).Resize(startLooking.rows.Count - rowIdx)
        rowIdx = MatchFirst(criteria, startLooking, exactMatch)
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
    If Err.number <> 0 Then Err.Clear
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
    Dim rowIdx As Variant
    Dim lstObjRowIdx As Long
    
    fieldIdx = GetFieldIndex(lstObj, field)
    Set startLooking = lstObj.ListColumns(fieldIdx).DataBodyRange
    rowIdx = MatchFirst(criteria, startLooking, exactMatch)
    Do While rowIdx > 0
        lstObjRowIdx = startLooking.rows(rowIdx).Row - lstObj.HeaderRowRange.Row
        dicItems(lstObjRowIdx) = 1
        
        If (startLooking.rows.Count - rowIdx) = 0 Then
            Exit Do
        End If
        
        Set startLooking = startLooking.Offset(rowIdx).Resize(startLooking.rows.Count - rowIdx)
        rowIdx = MatchFirst(criteria, startLooking, exactMatch)
    Loop

Finalize:
    On Error Resume Next

    If dicItems.Count > 0 Then
        Dim K As Variant, retV() As Variant
        ReDim retV(1 To dicItems.Count)
        Dim cnt As Long
        cnt = 1
        For Each K In dicItems.Keys
            retV(cnt) = K
            cnt = cnt + 1
        Next K
        GetFoundListIndexArray = retV
    Else
        GetFoundListIndexArray = Array()
    End If

    Set dicItems = Nothing
    If Err.number <> 0 Then Err.Clear
    Exit Function
E:
    ErrorCheck

End Function

'Return the range in a ListObject of all the matching values in a single ListColumn
Public Function GetFoundRange(listObj As ListObject, field As Variant, forceSortAS As XlSortOrder, criteria As Variant, matchMode As XMatchMode, Optional returnColumn As Variant) As Range

    If listObj.listRows.Count = 0 Then Exit Function
    AddSort listObj, field, forceSortAS, True

    Dim foundRange As Range
    If TypeName(criteria) = "Date" Then criteria = CDbl(criteria)

    Dim fieldIdx As Integer, foundRow As Long
    fieldIdx = GetFieldIndex(listObj, field)

    Dim firstRow As Long, lastRow As Long
    firstRow = GetFirstRowIndex(listObj, fieldIdx, criteria, matchMode, True, True)
    If firstRow > 0 Then
        lastRow = GetLastRowIndex(listObj, fieldIdx, criteria, matchMode, True, True)
        Set foundRange = listObj.ListColumns(fieldIdx).DataBodyRange(RowIndex:=firstRow)
        Set foundRange = foundRange.Resize(RowSize:=(lastRow - firstRow) + 1)
    End If

    If Not foundRange Is Nothing Then
        If IsMissing(returnColumn) = False Then
            Dim retColIdx As Long
            retColIdx = GetFieldIndex(listObj, returnColumn)
            Dim offsetBy As Long
            offsetBy = retColIdx - fieldIdx
            Set foundRange = foundRange.Offset(ColumnOffset:=offsetBy)
        End If
    End If

    Set GetFoundRange = foundRange

End Function

Public Function GetFirstRowIndex(ByRef listObj As ListObject, field As Variant, criteria As Variant, matchMode As XMatchMode, sortIfNeeded As Boolean, ClearFilter As Boolean) As Long
    GetFirstRowIndex = GetRowIndex(listObj, field, criteria, True, matchMode, sortIfNeeded, ClearFilter)
End Function
Public Function GetLastRowIndex(ByRef listObj As ListObject, field As Variant, criteria As Variant, matchMode As XMatchMode, sortIfNeeded As Boolean, ClearFilter As Boolean) As Long
    GetLastRowIndex = GetRowIndex(listObj, field, criteria, False, matchMode, sortIfNeeded, ClearFilter)
End Function



Public Function FindFirstInRange(ByVal rng As Range, criteria As Variant, Optional matchMode As XMatchMode = XMatchMode.exactMatch) As Double
On Error GoTo E:
    Dim foundIdx As Double
    
    If TypeName(criteria) = "Date" Then criteria = CDbl(criteria)
 
    foundIdx = WorksheetFunction.XMatch(criteria, rng, matchMode, 1)
    
    FindFirstInRange = foundIdx
    
    Exit Function
E:
    'eat the error, return 0
    Err.Clear
    FindFirstInRange = 0
End Function

Public Function FindLastInRange(ByVal rng As Range, criteria As Variant, Optional matchMode As XMatchMode = XMatchMode.exactMatch) As Double
On Error GoTo E:
    Dim foundIdx As Double
    
    If TypeName(criteria) = "Date" Then criteria = CDbl(criteria)
    
    foundIdx = WorksheetFunction.XMatch(criteria, rng, matchMode, -1)
    
    FindLastInRange = foundIdx
    
    Exit Function
E:
    'eat the error, return 0
    Err.Clear
    FindLastInRange = 0
End Function
Private Function GetRowIndex(ByRef listObj As ListObject, field As Variant, criteria As Variant, firstRowIndex As Boolean, matchMode As XMatchMode, sortIfNeeded As Boolean, clearFilters As Boolean) As Long
    
    
    
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
        If MatchFirst(criteria, .ListColumns(fieldIdx).DataBodyRange, matchMode) = 0 Then
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
                If srtField.Key.column <> .ListColumns(fieldIdx).Range.column Then
                    AddSort listObj, field, clearPreviousSorts:=True
                End If
            End If
        End If
        
        If firstRowIndex = True Then
            GetRowIndex = MatchFirst(criteria, .ListColumns(fieldIdx).DataBodyRange, matchMode)
        Else
            GetRowIndex = MatchLast(criteria, .ListColumns(fieldIdx).DataBodyRange, matchMode)
        End If
        
        
    End With
    
    
End Function

Public Function GetRangeMultipleCriteria(lo As ListObject, Columns As Variant, criteria As Variant, returnColumn As Variant, Optional setRangeValue As Variant) As Range

    If UBound(Columns) <> UBound(criteria) Then Exit Function
    
    
    Dim IDX As Long
    Dim retRange As Range
    
    'SORT ALL COLUMNS
    For IDX = LBound(Columns) To UBound(Columns)
        If IDX = LBound(Columns) Then
            AddSort lo, Columns(IDX), xlAscending, True, True
        Else
            AddSort lo, Columns(IDX), xlAscending, False, False
        End If
    Next IDX
    
    Dim firstIdx As Long, lastIdx As Long
    For IDX = LBound(Columns) To UBound(Columns)
        'If first, then use FindFirstRow/LastRow
        If IDX = LBound(Columns) Then
            firstIdx = GetFirstRowIndex(lo, Columns(IDX), criteria(IDX), exactMatch, False, False)
            lastIdx = GetLastRowIndex(lo, Columns(IDX), criteria(IDX), exactMatch, False, False)
            If firstIdx = 0 Or lastIdx = 0 Then Exit Function
        Else
            'we're on at least  the 2nd pass
            Dim lookInRange As Range
            Set lookInRange = lo.ListColumns(Columns(IDX)).DataBodyRange
            
            Set lookInRange = lookInRange.Offset(rowOffset:=firstIdx - 1).Resize(RowSize:=lastIdx - firstIdx + 1)
            Dim subFirst As Long, subLast As Long
            subFirst = MatchFirst(criteria(IDX), lookInRange, exactMatch)
            subLast = MatchLast(criteria(IDX), lookInRange, exactMatch)
            If subFirst > 0 And subLast > 0 Then
                firstIdx = firstIdx + subFirst - 1
                lastIdx = firstIdx + (subLast - subFirst)
            Else
                firstIdx = 0
                lastIdx = 0
            End If
        End If
    Next IDX
    
    If firstIdx > 0 And lastIdx > 0 Then
        Set retRange = lo.ListColumns(returnColumn).DataBodyRange.Offset(rowOffset:=firstIdx - 1).Resize(lastIdx - firstIdx + 1)
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
    If lstObj.Range.Worksheet.protectContents = True And lstObj.Range.Worksheet.protection.allowSorting = False Then
        UnprotectSheet lstObj.Range.Worksheet
    End If
    
    
    If clearFilters Then ClearFilter lstObj
    Dim colIdx As Long
    Dim needsSort As Boolean
    Dim tCols As Variant
    
   If LBound(cols) = UBound(cols) Then
        If isArray(cols(LBound(cols))) Then
            tCols = ArrArray(cols(LBound(cols)), aoNone)
        Else
            tCols = ArrParams(cols)
        End If
    Else
        tCols = ArrParams(cols)
    End If
    
    needsSort = True
    
    Dim evts As Boolean
    evts = Application.EnableEvents
    Application.EnableEvents = False
    If needsSort Then
        
     
        With lstObj.Sort
            .SortFields.Clear
            For colIdx = LBound(tCols) To UBound(tCols)
                .SortFields.Add lstObj.ListColumns(tCols(colIdx, 1)).DataBodyRange, SortOn:=xlSortOnValues, Order:=sortOrder
            Next colIdx
            .Apply
            DoEvents
        End With
    
    End If
   
   AddSortMultipleColumns = True
Finalize:
    Application.EnableEvents = evts
   
   Exit Function
E:
   AddSortMultipleColumns = False
   ErrorCheck
    Resume Finalize:
End Function

'Add Sort to ListObject, optionally clearing previous sorts
Public Function AddSort(listObj As ListObject, field As Variant, Optional Order As XlSortOrder = xlAscending, Optional clearPreviousSorts As Boolean = False, Optional clearFilters As Boolean = True) As Boolean
On Error GoTo E:

    
    If listObj.Range.Worksheet.protectContents = True And listObj.Range.Worksheet.protection.allowSorting = False Then
        UnprotectSheet listObj.Range.Worksheet
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
            .SortFields.Add listObj.ListColumns(field).DataBodyRange, SortOn:=xlSortOnValues, Order:=Order
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
                If Strings.Left(CStr(crit1), 1) <> "*" Then crit1 = "*" & crit1
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
    
    If listObj.Range.Worksheet.protectContents And listObj.Range.Worksheet.protection.allowFiltering = False Then
        UnprotectSheet listObj.Range.Worksheet
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
    Beep
    LogTRACE Err.Source & ", " & Err.Description
    Err.Clear
End Function

Public Function ColumnFiltered(listObj As ListObject, col As Variant) As Boolean
    Dim colIdx As Long
    colIdx = GetFieldIndex(listObj, col)
    If FilterCount(listObj) = 0 Then
        ColumnFiltered = False
    Else
        If Not listObj.AutoFilter.Filters Is Nothing Then
            ColumnFiltered = listObj.AutoFilter.Filters(colIdx).On
        End If
    End If
End Function

Public Function AddFilterBetween(listObj As ListObject, field As Variant, crit1 As Variant, crit2 As Variant, Optional clearExistFilters As Boolean = False) As Long
On Error GoTo E:
    
    Dim cnt As Long
    Dim fieldIdx As Long
    fieldIdx = GetFieldIndex(listObj, field)
    
    If listObj.Range.Worksheet.protectContents And listObj.Range.Worksheet.protection.allowFiltering = False Then
        UnprotectSheet listObj.Range.Worksheet
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
    Beep
    LogTRACE Err.Source & ", " & Err.Description
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
    If Err.number <> 0 Then Err.Clear
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
    
    If listObj.Range.Worksheet.protectContents And listObj.Range.Worksheet.protection.allowFiltering = False Then
        UnprotectSheet listObj.Range.Worksheet
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
    
    If Err.number <> 0 Then Err.Clear
    Exit Function
E:
   failed = True
   Err.Raise Err.number, Err.Source, Err.Description
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
        Dim rowIdx As Long
        If firstVisColIdx > 0 Then
            Dim vRng As Range
            Set vRng = lstObj.ListColumns(firstVisColIdx).Range.SpecialCells(xlCellTypeVisible)
            Dim areaRng As Range
            For Each areaRng In vRng.Areas
                If areaRng.Row > lstObj.HeaderRowRange.Row Then
                    rowIdx = areaRng.Row
                    Exit For
                Else
                    If areaRng.rows.Count > 1 Then
                        rowIdx = areaRng.rows(2).Row
                        Exit For
                    End If
                End If
            Next areaRng
        End If
        If goBeyondListObjectIfNoVisible Then
            If rowIdx = 0 Then
                rowIdx = lstObj.HeaderRowRange.Row + lstObj.listRows.Count + 1
                If lstObj.ShowTotals Then
                    rowIdx = rowIdx + 1
                End If
            End If
        End If
        If rowIdx > 0 Then
            FirstVisibleListRowIdx = rowIdx - lstObj.HeaderRowRange.Row
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

    LogTRACE "Getting Visible List Obj Row Count for: " & lstObj.Name


    Dim rwCount As Long, visColIdx As Long
    visColIdx = FirstVisibleListColIndex(lstObj)
    If visColIdx > 0 Then
        If lstObj.listRows.Count > 0 Then
            Dim visRng As Range
            Set visRng = lstObj.listRows(1).Range(1, visColIdx)
            Set visRng = visRng.Resize(RowSize:=lstObj.listRows.Count)
            
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
    Beep
    LogTRACE "Error getting VisibleListObjRows for " & lstObj.Name, True
    LogTRACE "Error: - " & Err.number & ", " & Err.Source & ", " & Err.Description
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
        retV.matchExactFirstIDX = MatchFirst(searchVal, rng, exactMatch, srchMode:=searchBinaryAsc)
    Else
        retV.matchExactFirstIDX = MatchFirst(searchVal, rng, exactMatch)
    End If
    
    If retV.matchExactFirstIDX > 0 Then
        retV.matchExactLastIDX = MatchLast(searchVal, rng, exactMatch)
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
                Set curRng = curRng.Resize(RowSize:=(tmpFound.matchExactLastIDX - tmpFound.matchExactFirstIDX) + 1)
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
    
    
    If Err.number <> 0 Then Err.Clear
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
    If Err.number <> 0 Then
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
    If Err.number <> 0 Then Err.Clear
Finalize:
    HasValidation = retV
    If Err.number <> 0 Then Err.Clear
    On Error GoTo 0
End Function
Public Function UniqueRowNumberInRange(ByVal Target As Range) As Long()

    Dim tmpD As New Dictionary
    tmpD.CompareMode = BinaryCompare
    Dim areaIdx As Long, rwIdx As Long
    Dim realRow As Long
    For areaIdx = 1 To Target.Areas.Count
        For rwIdx = 1 To Target.Areas(areaIdx).rows.Count
            realRow = Target.Areas(areaIdx).rows(rwIdx).Row
            If Not tmpD.Exists(realRow) Then
                tmpD(realRow) = realRow
            End If
        Next rwIdx
    Next areaIdx
    
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
    If Not wks.UsedRange Is Nothing Then
        LastColumnWithData = wks.UsedRange.Columns.Count + (wks.UsedRange.column - 1)
    End If
End Property

Public Function LastPopulatedRow(wks As Worksheet, Optional column As Variant) As Long
    Dim lpr As Long
    Dim rowOffset As Long, colOffset As Long, urColEnd As Long, urRowEnd As Long
    rowOffset = wks.UsedRange.Row - 1
    colOffset = wks.UsedRange.column - 1
    urColEnd = wks.UsedRange.Columns.Count + colOffset
    urRowEnd = wks.UsedRange.rows.Count + rowOffset
    Dim exactCol As Long: If Not IsMissing(column) Then exactCol = column
    If exactCol > 0 And exactCol < urColEnd Then urColEnd = exactCol
    '   If asking for Column outside last col with data, then return 0
    If exactCol > urColEnd Then
        lpr = 0
    '   HANDLE EMPTY SHEET ( OR JUST $A$1 HAS DATA)
    ElseIf urRowEnd = 1 And urColEnd = 1 Then
        lpr = IIf(Len(wks.Range("A1").Text) > 0, 1, 0)
        If Not exactCol > 0 And lpr = 1 Then
            If exactCol > 1 Then lpr = 0
        End If
    '   HANDLE SINGLE CELL POPULATED OTHER THAN $A$1
    ElseIf (VarType(wks.UsedRange.Text) And VbVarType.vbArray) = 0 Then
        lpr = urRowEnd
        If exactCol > 0 Then
            ' ONLY ONE CELL POPULATED IIF urColEnd doesn't match Column, then return 0
            If exactCol <> urColEnd Then lpr = 0
        End If
    Else
        lpr = urRowEnd
    End If
   'SHOULD BE GOOD, UNLESS THE ROW [LPR] ISN'T VISIBLE
   '(HIDDEN ROW THAT HAS NOT DATA IS STILL COUNTED IN USED RANGE, SO
   ' NOW WE NEED TO CHECK THAT)
   If lpr > 0 Then
        Dim deepCheck As Boolean
        If wks.rows(lpr).Hidden Then deepCheck = True
        If Not deepCheck And exactCol > 0 Then
            If Len(wks.Cells(lpr, exactCol).Text) = 0 Then deepCheck = True
        End If
        If deepCheck Then
            Dim rowIdx As Long, colIdx As Long
            Dim hasRowData As Boolean
            For rowIdx = lpr To 1 Step -1
                If exactCol > 0 Then
                    hasRowData = Len(wks.Cells(rowIdx, exactCol + colOffset)) > 0
                    If hasRowData Then
                        lpr = rowIdx
                    ElseIf rowIdx = 1 Then
                        'NO ROWS IN [COLUMN] HAVE ANY DATA
                        lpr = 0
                        Exit For
                    End If
                Else
                    For colIdx = 1 To urColEnd
                        hasRowData = Len(wks.Cells(rowIdx, colIdx + colOffset)) > 0
                        If hasRowData Then
                            lpr = rowIdx
                            Exit For
                        End If
                    Next colIdx
                End If
                DoEvents
                If hasRowData Then Exit For
            Next rowIdx
        End If
    End If
    LastPopulatedRow = lpr
End Function

'Public Property Get LastPopulatedRow(wks As Worksheet, Optional column As Variant) As Long
'    Dim ret As Long
'    ret = -1
'    If Not IsMissing(column) Then
'        If IsNumeric(column) Then
'            ret = wks.Cells(wks.Rows.count, CLng(column)).End(xlUp).Row
'        Else
'            ret = wks.Cells(wks.Rows.count, CStr(column)).End(xlUp).Row
'        End If
'    Else
'        ret = wks.usedRange.Rows.count + (wks.usedRange.Row - 1)
'    End If
'    LastPopulatedRow = ret
'End Property


Public Function GetA1CellRef(fromRng As Range, Optional colOffset As Long = 0, Optional rowCount As Long = 1, Optional colCount As Long = 1, Optional rowOffset As Long = 0, Optional fixedRef As Boolean = False, Optional visibleCellsOnly As Boolean = False) As String
'   return A1 style reference (e.g. "A10:A116") from selection
'   Optional offsets, resized ranges supported
    Dim tmpRng As Range
    Set tmpRng = fromRng.Offset(rowOffset, colOffset)
    If colCount > 1 Or rowCount > 1 Then
        Set tmpRng = tmpRng.Resize(rowCount, colCount)
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
    Dim rCount As Long, areaIdx As Long, rwIdx As Long
    
    If rng Is Nothing Then
        GoTo Finalize:
    End If
    
    'Check first if all First/Count are the same, if they are, no need to loop through everything
    If AreasMatchRows(rng) Then
        tmpCount = rng.Areas(1).rows.Count
    Else
        Set rowDict = New Dictionary
        For areaIdx = 1 To rng.Areas.Count
            For rwIdx = 1 To rng.Areas(areaIdx).rows.Count
                Dim realRow As Long
                realRow = rng.Areas(areaIdx).rows(rwIdx).Row
                rowDict(realRow) = realRow
            Next rwIdx
        Next areaIdx
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
    Dim firstCol As Long, areaIdx As Long, colIdx As Long
    
    If rng Is Nothing Then
        GoTo Finalize:
    End If
    
    
    If AreasMatchCols(rng) Then
        tmpCount = rng.Areas(1).Columns.Count
    Else
        Set colDict = New Dictionary
        For areaIdx = 1 To rng.Areas.Count
            firstCol = rng.Areas(areaIdx).column
            colDict(firstCol) = firstCol
            For colIdx = 1 To rng.Areas(areaIdx).Columns.Count
                If colIdx > 1 Then colDict(firstCol + (colIdx - 1)) = firstCol + (colIdx - 1)
            Next colIdx
        Next areaIdx
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
        Dim firstRow As Long, firstCount As Long, noMatch As Boolean, aIdx As Long
        firstRow = rng.Areas(1).Row
        firstCount = rng.Areas(1).rows.Count
        For aIdx = 2 To rng.Areas.Count
            With rng.Areas(aIdx)
                If .Row <> firstRow Or .rows.Count <> firstCount Then
                    noMatch = True
                    Exit For
                End If
            End With
        Next aIdx
        retV = Not noMatch
    End If

    AreasMatchRows = retV

End Function

Private Function AreasMatchCols(rng As Range) As Boolean

    Dim retV As Boolean
    If rng.Areas.Count = 1 Then
        retV = True
    Else
        Dim firstCol As Long, firstCount As Long, noMatch As Boolean, aIdx As Long
        firstCol = rng.Areas(1).column
        firstCount = rng.Areas(1).Columns.Count
        For aIdx = 2 To rng.Areas.Count
            With rng.Areas(aIdx)
                If .column <> firstCol Or .Columns.Count <> firstCount Then
                    noMatch = True
                    Exit For
                End If
            End With
        Next aIdx
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
            l1End = l1Start + rng.Areas(loop1).rows.Count - 1
            
            For loop2 = 1 To rng.Areas.Count
                l2Start = rng.Areas(loop2).Row
                l2End = l1Start + rng.Areas(loop2).rows.Count - 1
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





