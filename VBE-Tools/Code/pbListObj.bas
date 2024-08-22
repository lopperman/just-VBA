Attribute VB_Name = "pbListObj"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' pbListObj v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' General  Helper Utilities for Working with ListObnects
'
' @module pbListObj
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
'   KEY FUNCTIONS IN THIS MODULE
'
'   ** FORMULA RELATED **
'   ReplaceFormulasWithStatic: (Replaces all ListColumns in ListObject that have formulas, with the currnet values of those formulas)
'   ReplaceListColFormulaWithStatic: (replaces specified ListObject column with current values from ListColumn formula)
'   PopulateStaticFromFormula: (Set the Formula for a ListColumn, and then replace the ListColumn contents with the values)
'
'   ** TRANSFORMATION, STRUCTURE **
'   ResizeListObject_Totalsange (Resize the Totals Range of the ListObject)
'
'   ListColumnAsArray: (Return .DataBodyRange of ListColumn as 2D Array (1 to n, 1 to 1))
'   ListColumnExists: (Check if a value is a valid ListColumn name in specified ListObject)
'   ListColumnIndex: (return the ListColumn index for a listcolumn name.  Returns 0 (zero) if non-existant)
'   AddColumnIfMissing: (If missing, adds ListColumn. Optionally set ListColumn Index and NumberFormat

Option Explicit
Option Compare Text
Option Base 1

'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'       PAUL - MOVE THE THINGS BELOW TO THE MASTER COPY
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'Public Enum CellErrorCheck
'    ceFirstInRange = 1
'    ceAnyInRange = 2
'    ceAllInRange = 3
'End Enum

'Public Function IsCellError(ByVal srcRng As Range, Optional cellCheck As CellErrorCheck = CellErrorCheck.ceFirstInRange) As Boolean
'    Dim failed As Boolean
'    Dim c
'    If cellCheck = ceFirstInRange Then
'        IsCellError = isError(srcRng(1, 1))
'    ElseIf cellCheck = ceAnyInRange Then
'        For Each c In srcRng
'            If isError(c) Then
'                IsCellError = True
'                Exit For
'            End If
'        Next c
'    ElseIf cellCheck = ceAllInRange Then
'        Dim errCount As Long
'        For Each c In srcRng
'            If Not isError(c) Then
'                IsCellError = False
'                Exit For
'            Else
'                errCount = errCount + 1
'            End If
'        Next c
'        If errCount = srcRng.Count Then
'            IsCellError = True
'        End If
'    End If
'End Function



'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

'Public Function MoveColumn(lstObj As ListObject, colIdx As Long, newColIdx As Long) As Boolean
'    On Error Resume Next
'    If lstObj.ListColumns.Count >= colIdx And lstObj.ListColumns.Count >= newColIdx Then
'        With lstObj
'
'        End With
'    End If
'
'
'
'End Function

Public Function FindAndReplaceMatchingCols(lstColumnName As String, OldVal, NewVal, valType As VbVarType, Optional wkbk As Workbook, Optional strMatch As strMatchEnum = strMatchEnum.smEqual) As Boolean
    If wkbk Is Nothing Then
        Set wkbk = ThisWorkbook
    End If
    Dim tws As Worksheet, tLO As ListObject
    For Each tws In wkbk.Worksheets
        For Each tLO In tws.ListObjects
            If HasData(tLO) And ListColumnIndex(tLO, lstColumnName) > 0 Then
                FindAndReplaceListCol tLO, lstColumnName, OldVal, NewVal, valType, strMatch:=strMatch
            End If
        Next tLO
    Next tws
End Function

Public Function FindAndReplaceListCol(lstObj As ListObject, lstColIdxOrName, OldVal, NewVal, valType As VbVarType, Optional strMatch As strMatchEnum = strMatchEnum.smEqual) As Long
    ' ** NOTE ABOUT WORKSHEET PROTECTION **
    ' Although you do need to unprotect a sheet in some cases, you do NOT need to unprotect anything for this method to work,
    '   as long as your code to protect your worksheets sets the 'UserInterfaceOnly' parameter to TRUE
    ' VBA code can make certain changes to worksheets,** AS LONG AS ** the protection code has been called since the
    '  workbook was opened. (Worksheet may still be locked and prevent users for editing, but VBA will not be able to
    '  make changes unless the protection call has been call
    ' (I'm adding this note in because a lot of people follow the UNPROTECT --> MAKE CHANGES --> REPROTECT.  This IS needed
    '  in certain situations, like adding rows to a list object.  You just don't need that here)
    ' HOWEVER, you likely **DO** need to 'reprotect'.
    
    'Returns Count of items changed
    On Error GoTo e:
    Dim failed As Boolean
    
    Dim ERR_TYPE_MISMATCH As Long: ERR_TYPE_MISMATCH = 13
    Dim ERR_REPROTECT_SHEET As Long: ERR_REPROTECT_SHEET = 0
    
    
    
    Dim colIdx As Long, evts As Boolean, itemValid As Boolean
    evts = Events
    If StringsMatch(TypeName(lstColIdxOrName), "String") Then
        colIdx = ListColumnIndex(lstObj, CStr(lstColIdxOrName))
    End If
    If colIdx > 0 And colIdx <= lstObj.ListColumns.Count And HasData(lstObj) Then
        Select Case valType
            Case VbVarType.vbArray, VbVarType.vbDataObject, VbVarType.vbEmpty, VbVarType.vbError, VbVarType.vbObject, VbVarType.vbUserDefinedType, VbVarType.vbVariant
                Err.Raise ERR_TYPE_MISMATCH, Source:="pbListObj.FindAndReplaceListCol", Description:="VbVarType: " & valType & " is not supported"
        End Select
    
        Dim tmpArr, changedCount As Long, tmpIdx As Long, tmpValue As Variant
        tmpArr = ArrRange(lstObj.ListColumns(colIdx).DataBodyRange, aoNone)
        
        For tmpIdx = LBound(tmpArr) To UBound(tmpArr)
            itemValid = False
            tmpValue = tmpArr(tmpIdx, 1)
            If valType = vbBoolean Then
                If CBool(tmpValue) = CBool(OldVal) Then itemValid = True
            ElseIf valType = vbString Then
                If StringsMatch(tmpValue, OldVal, strMatch) Then itemValid = True
            ElseIf valType = vbDate Then
                If CDate(tmpValue) = CDate(OldVal) Then itemValid = True
            ElseIf IsNumeric(OldVal) Then
                If valType = vbDouble Then
                    If CDbl(tmpValue) = CDbl(OldVal) Then itemValid = True
                ElseIf valType = vbByte Then
                    If CByte(tmpValue) = CByte(OldVal) Then itemValid = True
                ElseIf valType = vbCurrency Then
                    If CCur(tmpValue) = CCur(OldVal) Then itemValid = True
                ElseIf valType = vbInteger Then
                    If CInt(tmpValue) = CInt(OldVal) Then itemValid = True
                ElseIf valType = vbLong Then
                    If CLng(tmpValue) = CLng(OldVal) Then itemValid = True
                ElseIf valType = vbSingle Then
                    If CSng(tmpValue) = CSng(OldVal) Then itemValid = True
                ElseIf valType = vbDecimal Then
                    If VarType(tmpValue) = vbDecimal And VarType(OldVal) = vbDecimal Then
                        If tmpValue = OldVal Then itemValid = True
                    End If
                End If
                If Not itemValid Then
                    If IsNumeric(tmpValue) Then
                        If tmpValue = OldVal Then itemValid = True
                    End If
                End If
            End If
            
            If itemValid Then
                tmpArr(tmpIdx, 1) = NewVal
                changedCount = changedCount + 1
            End If
        Next tmpIdx
    
        If changedCount > 0 Then
            EventsOff
            If lstObj.Range.Worksheet.protectContents Then
                ' You must call your code to PROTECT worksheet: lst.Range.Worksheet
                ' make sure UserInterfaceOnly:=True is included
                ProtectSheet lstObj.Range.Worksheet
            End If
            
            
            LogDEBUG ConcatWithDelim(" ", "replacing", changedCount, "items in", lstObj.Name, "table", "," & lstObj.ListColumns(colIdx).Name, "column", "OldValue:=", OldVal, "NewValue:=", NewVal)
            lstObj.ListColumns(colIdx).DataBodyRange.value = tmpArr
        End If
    End If
        
Finalize:
    On Error Resume Next
    Events = evts
    If Not failed Then
        FindAndReplaceListCol = changedCount
    End If
    
    If Err.number <> 0 Then Err.Clear
    Exit Function
e:
    failed = True
    ErrorCheck
    Resume Finalize:
End Function



'   ~~~ ~~~ Sort's a List Object by [lstColIdx] Ascending.  If already
'                   sorted Ascending, sorts Descending
Public Function UserSort(lstObj As ListObject, lstColIdx As Long)
        If lstObj.listRows.Count = 0 Then Exit Function
        If lstObj.Range.Worksheet.protection.allowSorting = False Then
            MsgBox_FT "This table does not allow custom sorting", vbOKOnly + vbInformation, "SORRY"
            Exit Function
        End If
        Dim clearPreviousSort As Boolean, hdrCol As Long
        Dim orderBy As XlSortOrder
        orderBy = xlAscending
        hdrCol = lstObj.HeaderRowRange.column
        If lstObj.Sort.SortFields.Count > 1 Then
            clearPreviousSort = True
        Else
            If lstObj.Sort.SortFields.Count = 1 Then
                If lstObj.Sort.SortFields(1).Key.column - hdrCol + 1 <> lstColIdx Then
                    clearPreviousSort = True
                End If
            End If
        End If
        If clearPreviousSort Then
            lstObj.Sort.SortFields.Clear
        End If
        With lstObj.Sort
            If .SortFields.Count = 1 Then
                If Not .SortFields(1).SortOn = xlSortOnValues Then .SortFields(1).SortOn = xlSortOnValues
                If .SortFields(1).Order = xlAscending Then
                    .SortFields(1).Order = xlDescending
                Else
                    .SortFields(1).Order = xlAscending
                End If
            Else
                .SortFields.Add lstObj.ListColumns(lstColIdx).DataBodyRange, SortOn:=XlSortOn.xlSortOnValues, Order:=XlSortOrder.xlAscending
            End If
            .Apply
        End With
Finalize:
    On Error Resume Next
    If Err.number <> 0 Then Err.Clear
    Exit Function
e:
    ErrorCheck
    Resume Finalize:
End Function

'   ~~~ ~~~ Resize ListObject (Rows) ~~~ ~~~
Public Function ResizeListObjectRows(lstObj As ListObject, Optional totalRowCount As Long = 0, Optional addRowCount As Long = 0, Optional canDecreaseRowCount As Boolean = False) As Range
'   Resize the DataBodyRange area of a ListObject By resizing over existing sheet area (must faster than adding inserting/pushing down)
    
On Error GoTo e:
    Dim failed                              As Boolean
    Dim newRowCount             As Long
    Dim nonBodyRowCount     As Long
    Dim lastRealDataRow         As Long
    Dim newRowRange             As Range
    
    If totalRowCount <= 0 And addRowCount <= 0 Then Exit Function
    
    '   Confirm both rowcount parameters were not used (totalRowCount, addRowCount)
    If totalRowCount > 0 And totalRowCount < lstObj.listRows.Count And canDecreaseRowCount = False Then
        Err.Raise 17
        'RaiseError ERR_LIST_OBJECT_RESIZE_CANNOT_DELETE, "Resizing Failed because new 'totalRowCount' is less than existing row count"
    End If
    If totalRowCount > 0 And addRowCount > 0 Then
        Err.Raise 17
        'RaiseError ERR_LIST_OBJECT_RESIZE_INVALID_ARGUMENTS, "Resizing Failed, cannot set totalRowCount AND addRowCount"
    End If
        
    If totalRowCount > 0 And lstObj.listRows.Count > totalRowCount And canDecreaseRowCount = True Then
        Dim deleteRowCount As Long
        deleteRowCount = lstObj.listRows.Count - totalRowCount
        Dim delRowRange As Range
        Set delRowRange = lstObj.listRows(totalRowCount + 1).Range
        Set delRowRange = delRowRange.Resize(RowSize:=deleteRowCount)
        delRowRange.Delete xlShiftUp
        Set ResizeListObjectRows = lstObj.DataBodyRange
        Exit Function
    End If
        
    If addRowCount > 0 Then
        newRowCount = lstObj.listRows.Count + addRowCount
    Else
        addRowCount = totalRowCount - lstObj.listRows.Count
        newRowCount = totalRowCount
    End If
    '   Include Header range and TotalsRange (if applicable) in overall ListObject Range Size
    nonBodyRowCount = HeaderRangeRows(lstObj) + TotalsRangeRows(lstObj)
        
    lastRealDataRow = lstObj.HeaderRowRange.Row + lstObj.HeaderRowRange.rows.Count - 1
    If lstObj.listRows.Count > 0 Then lastRealDataRow = lastRealDataRow + lstObj.listRows.Count
    
    '   Resize ListObject Range with new range
    lstObj.Resize lstObj.Range.Resize(RowSize:=newRowCount + nonBodyRowCount)
    
    Set newRowRange = lstObj.Range.Worksheet.Cells(lastRealDataRow + 1, lstObj.Range.column)
    Set newRowRange = newRowRange.Resize(RowSize:=addRowCount, ColumnSize:=lstObj.ListColumns.Count)

Finalize:
    On Error Resume Next
    
    If Not failed Then
        Set ResizeListObjectRows = newRowRange
    End If
    
    Set newRowRange = Nothing
    
    Exit Function
e:
    failed = True
    '   add your own error handline rules here
    MsgBox "Error: " & Err.number & ", " & Err.Description
    'ErrorCheck
    If Err.number <> 0 Then Err.Clear
    Resume Finalize:

End Function


    Public Function HeaderRangeRows(lstObj As ListObject) As Long
    ' *** Returns -1 for Error, other Num Rows in HeaderRowRange
    On Error Resume Next
        HeaderRangeRows = lstObj.HeaderRowRange.rows.Count
        If Err.number <> 0 Then
            HeaderRangeRows = -1
        End If
        If Err.number <> 0 Then Err.Clear
    End Function
    
    Public Function TotalsRangeRows(lstObj As ListObject) As Long
    ' *** Returns -1 for Error, other Num Rows in HeaderRowRange
    On Error Resume Next
        If Not lstObj.TotalsRowRange Is Nothing Then
            If lstObj.ShowTotals Then
                TotalsRangeRows = lstObj.TotalsRowRange.rows.Count
            End If
        End If
        If Err.number <> 0 Then
            TotalsRangeRows = -1
        End If
        If Err.number <> 0 Then Err.Clear
    End Function




Public Function ReplaceFormulasWithStatic(lstObj As ListObject) As Boolean
'   REPLACES ALL FORMULAS IN LIST COLUMS, WITH THE VALUES
'   Helpful for situations like creating static copies of tables/listobjects
    If lstObj.listRows.Count = 0 Then Exit Function
    Dim lc As listColumn
    For Each lc In lstObj.ListColumns
        If lc.DataBodyRange(1, 1).HasFormula Then
            ReplaceListColFormulaWithStatic lstObj, lc.Index
        End If
    Next lc
    Set lc = Nothing
End Function

Public Function ReplaceListColFormulaWithStatic(lstObj As ListObject, column As Variant)
'   Replaces listColumn formula with static values
'   Column = Name of ListColumn or Index of ListColumn
    
    Dim tARR As Variant
    If Not lstObj.ListColumns(column).DataBodyRange Is Nothing Then
        If lstObj.ListColumns(column).DataBodyRange(1, 1).HasFormula Then
            tARR = ListColumnAsArray(lstObj, lstObj.ListColumns(column).Name)
            If ArrDimensions(tARR) = 2 Then
                lstObj.ListColumns(column).DataBodyRange.ClearContents
                lstObj.ListColumns(column).DataBodyRange.value = tARR
            End If
        End If
    End If
    
    If ArrDimensions(tARR) > 0 Then Erase tARR
    
    
End Function

Public Function CreateListColumnFormula(lstObj As ListObject, lstColName As String, _
    r1c1Formula As String, _
    Optional createColumnIfMissing As Boolean = True, _
    Optional convertToValues As Boolean = False, _
    Optional numberFormat As String = vbNullString) As Boolean

'   Create a formula in ListColumn, and then optionally replace the ListColumn contents with the values from the Formula
'   Note: Sets the 'Formula2R1C1' Formula Property.
On Error GoTo e:
    Dim failed As Boolean
    If lstObj.listRows.Count = 0 Then
        Exit Function
    End If
    
    Dim tmpColArr As Variant
    If ListColumnExists(lstObj, lstColName) = False And createColumnIfMissing = True Then
        AddColumnIfMissing lstObj, lstColName
    End If
    If Not ListColumnExists(lstObj, lstColName) Then
        failed = True
        GoTo Finalize:
    End If
    
    lstObj.ListColumns(lstColName).DataBodyRange.ClearContents
    lstObj.ListColumns(lstColName).DataBodyRange.numberFormat = "General"
    lstObj.ListColumns(lstColName).DataBodyRange.Formula2R1C1 = r1c1Formula
     
    If Len(numberFormat) > 0 Then
        lstObj.ListColumns(lstColName).DataBodyRange.numberFormat = numberFormat
    End If
     
    If convertToValues Then
        ReplaceListColFormulaWithStatic lstObj, lstColName
'        tmpColArr = ListColumnAsArray(lstObj, lstColName)
'        lstObj.ListColumns(lstColName).DataBodyRange.ClearContents
'        If Len(numberFormat) > 0 Then
'            lstObj.ListColumns(lstColName).DataBodyRange.numberFormat = numberFormat
'        End If
'        If ArrDimensions(tmpColArr) = 2 Then
'            lstObj.ListColumns(lstColName).DataBodyRange.value = tmpColArr
'        Else
'            failed = True
'        End If
    End If

Finalize:
    On Error Resume Next
    
    If ArrDimensions(tmpColArr) > 0 Then Erase tmpColArr
    CreateListColumnFormula = Not failed
    
    If Err.number <> 0 Then Err.Clear
    Exit Function
e:
    failed = True
    ErrorCheck
    Resume Finalize:
End Function

'Public Function PopulateStaticFromFormula(lstObj As ListObject, lstColName As String, r1c1Formula As String, Optional createIfMissing As Boolean = True, Optional numberFormat As String = vbNullString) As Boolean
''   Create a formula in ListColumn, and then replace the ListColumn contents with the values from the Formula
''   Note: Sets the 'Formula2R1C1' Formula Property.
'On Error GoTo E:
'    Dim failed As Boolean
'    If lstObj.listRows.Count = 0 Then
'        Exit Function
'    End If
'
'
'    Dim tmpColArr As Variant
'    If ListColumnExists(lstObj, lstColName) = False And createIfMissing = True Then
'        AddColumnIfMissing lstObj, lstColName
'    End If
'
'    If Not ListColumnExists(lstObj, lstColName) Then
'        failed = True
'        GoTo Finalize:
'    End If
'
'    lstObj.ListColumns(lstColName).DataBodyRange.ClearContents
'    lstObj.ListColumns(lstColName).DataBodyRange.numberFormat = "General"
'    lstObj.ListColumns(lstColName).DataBodyRange.Formula2R1C1 = r1c1Formula
'
'    ReplaceFormulasWithStatic lstObj
'
'    tmpColArr = ListColumnAsArray(lstObj, lstColName)
'    lstObj.ListColumns(lstColName).DataBodyRange.ClearContents
'    If Len(numberFormat) > 0 Then
'        lstObj.ListColumns(lstColName).DataBodyRange.numberFormat = numberFormat
'    End If
'    If ArrDimensions(tmpColArr) = 2 Then
'        lstObj.ListColumns(lstColName).DataBodyRange.value = tmpColArr
'    Else
'        failed = True
'    End If
'
'Finalize:
'    On Error Resume Next
'
'    If ArrDimensions(tmpColArr) > 0 Then Erase tmpColArr
'    PopulateStaticFromFormula = Not failed
'
'
'    If Err.Number <> 0 Then Err.Clear
'    Exit Function
'E:
'    failed = True
'    ErrorCheck
'    Resume Finalize:
'End Function
Public Function ListColumnAsArray(lstObj As ListObject, colName As String) As Variant
'   Get's the **DATABODYRANGE** Of ListObject column 'colName' into a 2D array
'   (states already dealt with)
    Dim tmpArr As Variant
    Dim tmpAI As ArrInformation
    
    If lstObj.listRows.Count = 0 Then
        Exit Function
    End If
    tmpArr = ArrListCols(lstObj, aoNone, colName)
    tmpAI = ArrayInfo(tmpArr)
    If tmpAI.Dimensions = 2 Then
        ListColumnAsArray = tmpArr
    End If
    If tmpAI.Dimensions > 0 Then
        Erase tmpArr
    End If
End Function


Public Function FindMatchingListColumns(colName As String, Optional matchType As strMatchEnum = strMatchEnum.smEqual) As Collection
'   for any matching column, return collection of array {[listObjectName], [listColumnName]}

    Dim retCol As New Collection

    Dim tws As Worksheet, tLO As ListObject, tCol As listColumn
    For Each tws In ThisWorkbook.Worksheets
        For Each tLO In tws.ListObjects
            For Each tCol In tLO.ListColumns
                If StringsMatch(tCol.Name, colName, matchType) Then
                    retCol.Add Array(tLO.Name, tCol.Name)
                End If
            Next
        Next
    Next

    Set FindMatchingListColumns = retCol
    
    Set retCol = Nothing

End Function

Public Function ListColumnExists(ByRef lstObj As ListObject, lstColName As String) As Boolean
'   Return true If 'lstColName' is a valid ListColumn in 'lstObj'
    ListColumnExists = ListColumnIndex(lstObj, lstColName) > 0
End Function

Public Function ListColumnIndex(ByRef lstObj As ListObject, lstColName As String) As Long
    Dim lstCol As listColumn
    For Each lstCol In lstObj.ListColumns
        If StringsMatch(lstCol.Name, lstColName) Then
            ListColumnIndex = lstCol.Index
            Exit For
        End If
    Next lstCol
End Function



Public Function AddColumnIfMissing(lstObj As ListObject, colName As String, Optional position As Long = -1, Optional numberFormat As String = vbNullString) As Boolean
'   Add column 'colName' to 'lstObj', if missing. Optionally provide ListColumn position, and NumberFormat for data display
On Error Resume Next
    
    
    Dim retV As Boolean
    If Not ListColumnExists(lstObj, colName) Then
        Dim nc As listColumn
        If position > 0 Then
            Set nc = lstObj.ListColumns.Add(position:=position)
        Else
            Set nc = lstObj.ListColumns.Add
        End If
        nc.Name = colName
        If lstObj.listRows.Count > 0 And numberFormat <> vbNullString Then
            nc.DataBodyRange.numberFormat = numberFormat
        End If
    End If
    AddColumnIfMissing = (Err.number = 0)
    If Err.number <> 0 Then Err.Clear
    Set nc = Nothing
    
    If Err.number <> 0 Then Err.Clear
End Function

Public Function CountBlanks(lstObj As ListObject, listColumn As Variant) As Long
On Error Resume Next
    Dim rng As Range
    If HasData(lstObj) = False Then
        CountBlanks = 0
        Exit Function
    End If
    Set rng = lstObj.ListColumns(listColumn).DataBodyRange.SpecialCells(xlCellTypeBlanks)
    If Not rng Is Nothing Then
        CountBlanks = rng.rows.Count
    End If
    Set rng = Nothing
    If Err.number <> 0 Then Err.Clear

End Function

Public Function ClearFormulas(lstObj As ListObject, ListCol As Variant, Optional ByVal showMsg As Boolean = False) As Long
On Error Resume Next
    
    Dim fmlaRng As Range, retV As Long
    If HasData(lstObj) Then
        Set fmlaRng = lstObj.ListColumns(ListCol).DataBodyRange
        Set fmlaRng = fmlaRng.SpecialCells(xlCellTypeFormulas)
        If Err.number <> 0 Then
            Err.Clear
        Else
            If Not fmlaRng Is Nothing Then
                retV = fmlaRng.Count
                fmlaRng.value = Empty
            End If
        End If
    End If
        
    ClearFormulas = retV
    If showMsg And retV > 0 Then
        MsgBox_FT Concat(retV, " values were removed from ", lstObj.Range.Worksheet.Name, "[", lstObj.Name, "].[", lstObj.ListColumns(ListCol).Name, "] becase formulas are not allowed to be manually entered anywere in the Financial Tool"), vbCritical + vbOKOnly, "OOPS"
        DoEvents
    End If

    

End Function

Public Function HasData(lstObj As Variant) As Boolean
On Error Resume Next
    If TypeName(lstObj) = "ListObject" Then
        HasData = lstObj.listRows.Count > 0
    ElseIf TypeName(lstObj) = "String" Then
        HasData = wt(CStr(lstObj)).listRows.Count > 0
    End If
    If Err.number <> 0 Then Err.Clear
End Function


Public Function ListColumnRange(srcListobj As ListObject, lstColumn As Variant, Optional includeHeaderRow As Boolean = False, Optional includeTotalsRow As Boolean = False) As Range

    Dim tmpRange As Range, tmpRowCount As Long
    
    With srcListobj
        tmpRowCount = .listRows.Count
        If includeHeaderRow Then tmpRowCount = tmpRowCount + HeaderRangeRows(srcListobj)
        If includeTotalsRow And .ShowTotals Then tmpRowCount = tmpRowCount + TotalsRangeRows(srcListobj)
        
        If tmpRowCount = 0 Then Exit Function
        
        Set tmpRange = srcListobj.ListColumns(lstColumn).Range
    
        If includeHeaderRow = False Then
            Set tmpRange = tmpRange.Offset(rowOffset:=HeaderRangeRows(srcListobj)).Resize(RowSize:=tmpRange.rows.Count - HeaderRangeRows(srcListobj))
        End If
        
        If includeTotalsRow = False And .ShowTotals Then
            Set tmpRange = tmpRange.Resize(RowSize:=tmpRange.rows.Count - TotalsRangeRows(srcListobj))
        End If
    End With

    Set ListColumnRange = tmpRange
    
    Set tmpRange = Nothing

End Function

' *** GET NEW RANGE AREA OF RESIZED LIST OBJECT
Public Function NewRowsRange(ByRef lstObj As ListObject, addRowCount As Long) As Range
On Error GoTo e:
    
    Dim evts As Boolean, prot As Boolean
    evts = Application.EnableEvents
    prot = lstObj.Range.Worksheet.protectContents
    
    
    If addRowCount <= 0 Then Exit Function
    
    Application.EnableEvents = False
    If prot Then
        UnprotectSheet lstObj.Range.Worksheet
    End If
    
    pbRange.ClearFilter lstObj
    
    Dim hdrRngRows As Long, totRngRows As Long, listRows As Long
    hdrRngRows = lstObj.HeaderRowRange.rows.Count
    If lstObj.TotalsRowRange Is Nothing Then
        totRngRows = 0
    Else
        totRngRows = lstObj.TotalsRowRange.rows.Count
    End If
    listRows = lstObj.listRows.Count
    
    Dim endCount As Long
    endCount = lstObj.listRows.Count + addRowCount
    
    lstObj.Resize lstObj.Range.Resize(RowSize:=(hdrRngRows + totRngRows + listRows) + addRowCount)
    Do While lstObj.listRows.Count < endCount
        lstObj.listRows.Add
    Loop
    
    Dim firstNewRow As Long, lastNewRow As Long
    firstNewRow = lstObj.listRows.Count - addRowCount + 1
    lastNewRow = firstNewRow + (addRowCount - 1)
        
    Set NewRowsRange = lstObj.listRows(firstNewRow).Range.Resize(RowSize:=addRowCount)
    
    
Finalize:
    On Error Resume Next
        If prot Then
            ProtectSheet lstObj.Range.Worksheet
        End If
        Application.EnableEvents = evts
    If Err.number <> 0 Then Err.Clear
    Exit Function
e:
    ErrorCheck
    Resume Finalize:
    

End Function

Public Function ClearData(lstObj As ListObject) As Boolean
    On Error Resume Next
    If HasData(lstObj) Then
        lstObj.DataBodyRange.Delete
    End If
    If Err.number <> 0 Then
        Beep
        LogERROR "Could not delete rows from " & lstObj.Range.Worksheet.Name & ":" & lstObj.Name
        Err.Clear
    End If
    ClearData = lstObj.listRows.Count = 0

End Function

Public Function FindListObject(extWB As Workbook, lstName As String) As ListObject

    Dim retLO As ListObject

    Dim eWS As Worksheet, eLO As ListObject
    For Each eWS In extWB.Worksheets
        For Each eLO In eWS.ListObjects
            If StringsMatch(lstName, eLO.Name) Then
                Set retLO = eLO
                Exit For
            End If
        Next eLO
        If Not retLO Is Nothing Then
            Exit For
        End If
    Next eWS
    
    If Not retLO Is Nothing Then
        Set FindListObject = retLO
    End If
    Set retLO = Nothing

End Function

'Obtain the first row given a set of headers and values
'@param {ListObject} The listobject to search within.
'@param {Array<String>} Headers to look in
'@param {Array<Variant>} Key to lookup
'@returns {Long} Row index containing the key data
Function getFirstRow(ByVal lo As ListObject, Headers, values, Optional cachedDict As Dictionary, Optional forceRebuild As Boolean = False) As Long
   
'    forceRebuild = True
   
   Dim chi As Long
   For chi = LBound(Headers) To UBound(Headers)
    If IsNumeric(Headers(chi)) Then Headers(chi) = lo.ListColumns(Headers(chi)).Name
   Next chi
  
    If Not cachedDict Is Nothing Then
        Dim vID As String: vID = Join(values, "|") & "|"
        If cachedDict.Exists(vID) Then
            getFirstRow = cachedDict(vID).item(1)
        Else
            getFirstRow = 0
        End If
        Exit Function
    Else
       Static sID As String: If sID = "" Then sID = lo.Name & "|" & Join(Headers, "|")
       Static oIndex As Dictionary
        If forceRebuild Or oIndex Is Nothing Or sID <> (lo.Name & "|" & Join(Headers, "|")) Then
            Set oIndex = getIndex(lo, Headers)
        End If
        If oIndex.Exists(Join(values, "|") & "|") Then
            getFirstRow = oIndex(Join(values, "|") & "|").item(1)
        End If
    End If
  
End Function

'Obtain a lookup dictionary which helps you find rows matching a set of headers.
'@param {ListObject} The listobject to search within.
'@param {Array<String>} Headers to create a lookup for
'@returns {Dictionary<string, Collection<Long>>} Lookup to find rows matching criteria
'Private Function getIndex(ByVal lo As ListObject, headers As Variant) As Object
Public Function getIndex(ByVal lo As ListObject, Headers As Variant) As Dictionary
On Error GoTo e:

   Dim chi As Long
   For chi = LBound(Headers) To UBound(Headers)
    If IsNumeric(Headers(chi)) Then Headers(chi) = lo.ListColumns(Headers(chi)).Name
   Next chi
  
  Dim v: v = lo.Range.value
  Dim iLB As Long: iLB = LBound(Headers)
  Dim iUB As Long: iUB = UBound(Headers)
  
  'Create field indices
  Dim iHeaders() As Long: ReDim iHeaders(iLB To iUB)
  Dim iUBH As Long: iUBH = UBound(v, 2)
  Dim i As Long
  For i = iLB To iUB
    Dim j As Long
    For j = 1 To iUBH
      If Headers(i) = v(1, j) Then
        iHeaders(i) = j
        Exit For
      End If
    Next
  Next
  
  'Create indexer
    Dim oDict As Dictionary
'    #If Mac Then
        ' if needed, get Dictionary Class from: ' (c) Tim Hall - https://github.com/timhall/VBA-Dictionary
        ' (Can be used on Mac and PC)
        Set oDict = New Dictionary
'    #Else
'        Set oDict = CreateObject("Scripting.Dictionary")
'    #End If
  For i = 1 To UBound(v, 1)
    Dim sID As String: sID = ""
    For j = iLB To iUB
      sID = sID & v(i, iHeaders(j)) & "|"
    Next
    If Not oDict.Exists(sID) Then Set oDict(sID) = New Collection
'    Call oDict(sID).Add(i - 1)
    oDict(sID).Add (i - 1)
  Next
  Set getIndex = oDict
  Exit Function
e:
    ftBeep btError
    Stop
  Resume
End Function



Public Function ListRowIdxFromWksht(lstObj As ListObject, worksheetRow As Long) As Long
    Dim hdrRow As Long
    hdrRow = lstObj.HeaderRowRange.Row + (1 - lstObj.HeaderRowRange.rows.Count)
    If worksheetRow - hdrRow > 0 And worksheetRow - hdrRow <= lstObj.listRows.Count Then
        ListRowIdxFromWksht = worksheetRow - hdrRow
    End If
End Function
Public Function ListColIdxFromWksht(lstObj As ListObject, worksheetCol As Long) As Long
    Dim firstCol As Long
    firstCol = lstObj.Range.column
    If worksheetCol - firstCol + 1 <= lstObj.ListColumns.Count Then
        ListColIdxFromWksht = worksheetCol - firstCol + 1
    End If
End Function

