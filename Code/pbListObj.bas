Attribute VB_Name = "pbListObj"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  Utility Methods for ListObjects
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbListObj.bas
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
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
Option Private Module

Public Function FindAndReplaceMatchingCols(lstColumnName As String, oldVal, newVal, valType As VbVarType, Optional wkbk As Workbook, Optional strMatch As strMatchEnum = strMatchEnum.smEqual) As Boolean
    If wkbk Is Nothing Then
        Set wkbk = ThisWorkbook
    End If
    Dim tWS As Worksheet, tLO As ListObject
    For Each tWS In wkbk.Worksheets
        For Each tLO In tWS.ListObjects
            If hasData(tLO) And ListColumnIndex(tLO, lstColumnName) > 0 Then
                FindAndReplaceListCol tLO, lstColumnName, oldVal, newVal, valType, strMatch:=strMatch
            End If
        Next tLO
    Next tWS
End Function

Public Function FindAndReplaceListCol(lstObj As ListObject, lstColIdxOrName, oldVal, newVal, valType As VbVarType, Optional strMatch As strMatchEnum = strMatchEnum.smEqual) As Long
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
    On Error GoTo E:
    Dim failed As Boolean
    
    Dim ERR_TYPE_MISMATCH As Long: ERR_TYPE_MISMATCH = 13
    Dim ERR_REPROTECT_SHEET As Long: ERR_REPROTECT_SHEET = 0
    
    
    
    Dim colidx As Long, evts As Boolean, itemValid As Boolean
    evts = Events
    If StringsMatch(TypeName(lstColIdxOrName), "String") Then
        colidx = ListColumnIndex(lstObj, CStr(lstColIdxOrName))
    End If
    If colidx > 0 And colidx <= lstObj.ListColumns.Count And hasData(lstObj) Then
        Select Case valType
            Case VbVarType.vbArray, VbVarType.vbDataObject, VbVarType.vbEmpty, VbVarType.vbError, VbVarType.vbObject, VbVarType.vbUserDefinedType, VbVarType.vbVariant
                Err.Raise ERR_TYPE_MISMATCH, Source:="pbListObj.FindAndReplaceListCol", Description:="VbVarType: " & valType & " is not supported"
        End Select
    
        Dim tmpARR, changedCount As Long, tmpIdx As Long, tmpValue As Variant
        tmpARR = ArrRange(lstObj.ListColumns(colidx).DataBodyRange, aoNone)
        
        For tmpIdx = LBound(tmpARR) To UBound(tmpARR)
            itemValid = False
            tmpValue = tmpARR(tmpIdx, 1)
            If valType = vbBoolean Then
                If CBool(tmpValue) = CBool(oldVal) Then itemValid = True
            ElseIf valType = vbString Then
                If StringsMatch(tmpValue, oldVal, strMatch) Then itemValid = True
            ElseIf valType = vbDate Then
                If CDate(tmpValue) = CDate(oldVal) Then itemValid = True
            ElseIf IsNumeric(oldVal) Then
                If valType = vbDouble Then
                    If CDbl(tmpValue) = CDbl(oldVal) Then itemValid = True
                ElseIf valType = vbByte Then
                    If CByte(tmpValue) = CByte(oldVal) Then itemValid = True
                ElseIf valType = vbCurrency Then
                    If CCur(tmpValue) = CCur(oldVal) Then itemValid = True
                ElseIf valType = vbInteger Then
                    If CInt(tmpValue) = CInt(oldVal) Then itemValid = True
                ElseIf valType = vbLong Then
                    If CLng(tmpValue) = CLng(oldVal) Then itemValid = True
                ElseIf valType = vbSingle Then
                    If CSng(tmpValue) = CSng(oldVal) Then itemValid = True
                ElseIf valType = vbDecimal Then
                    If VarType(tmpValue) = vbDecimal And VarType(oldVal) = vbDecimal Then
                        If tmpValue = oldVal Then itemValid = True
                    End If
                End If
                If Not itemValid Then
                    If IsNumeric(tmpValue) Then
                        If tmpValue = oldVal Then itemValid = True
                    End If
                End If
            End If
            
            If itemValid Then
                tmpARR(tmpIdx, 1) = newVal
                changedCount = changedCount + 1
            End If
        Next tmpIdx
    
        If changedCount > 0 Then
            EventsOff
            If lstObj.Range.Worksheet.ProtectContents Then
                ' You must call your code to PROTECT worksheet: lst.Range.Worksheet
                ' make sure UserInterfaceOnly:=True is included
                ProtectSht lstObj.Range.Worksheet
            End If
            LogDEV ConcatWithDelim(" ", "replacing", changedCount, "items in", lstObj.Name, "table", "," & lstObj.ListColumns(colidx).Name, "column", "OldValue:=", oldVal, "NewValue:=", newVal)
            lstObj.ListColumns(colidx).DataBodyRange.Value = tmpARR
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
E:
    failed = True
    ErrorCheck
    Resume Finalize:
End Function



'   ~~~ ~~~ Sort's a List Object by [lstColIdx] Ascending.  If already
'                   sorted Ascending, sorts Descending
Public Function UserSort(lstObj As ListObject, lstColIdx As Long)
        If lstObj.listRows.Count = 0 Then Exit Function
        If lstObj.Range.Worksheet.Protection.AllowSorting = False Then
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
                If lstObj.Sort.SortFields(1).key.column - hdrCol + 1 <> lstColIdx Then
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
E:
    ErrorCheck
    Resume Finalize:
End Function

'   ~~~ ~~~ Resize ListObject (Rows) ~~~ ~~~
Public Function ResizeListObjectRows(lstObj As ListObject, Optional totalRowCount As Long = 0, Optional addRowCount As Long = 0) As Range
'   Resize the DataBodyRange area of a ListObject By resizing over existing sheet area (must faster than adding inserting/pushing down)
    
On Error GoTo E:
    Dim failed                              As Boolean
    Dim newRowCount             As Long
    Dim nonBodyRowCount     As Long
    Dim lastRealDataRow         As Long
    Dim newRowRange             As Range
    
    If totalRowCount <= 0 And addRowCount <= 0 Then Exit Function
    
    '   Confirm both rowcount parameters were not used (totalRowCount, addRowCount)
    If totalRowCount > 0 And totalRowCount < lstObj.listRows.Count Then
        Err.Raise 17
        'RaiseError ERR_LIST_OBJECT_RESIZE_CANNOT_DELETE, "Resizing Failed because new 'totalRowCount' is less than existing row count"
    End If
    If totalRowCount > 0 And addRowCount > 0 Then
        Err.Raise 17
        'RaiseError ERR_LIST_OBJECT_RESIZE_INVALID_ARGUMENTS, "Resizing Failed, cannot set totalRowCount AND addRowCount"
    End If
        
    If addRowCount > 0 Then
        newRowCount = lstObj.listRows.Count + addRowCount
    Else
        addRowCount = totalRowCount - lstObj.listRows.Count
        newRowCount = totalRowCount
    End If
    '   Include Header range and TotalsRange (if applicable) in overall ListObject Range Size
    nonBodyRowCount = HeaderRangeRows(lstObj) + TotalsRangeRows(lstObj)
        
    lastRealDataRow = lstObj.HeaderRowRange.Row + lstObj.HeaderRowRange.Rows.Count - 1
    If lstObj.listRows.Count > 0 Then lastRealDataRow = lastRealDataRow + lstObj.listRows.Count
    
    '   Resize ListObject Range with new range
    lstObj.Resize lstObj.Range.Resize(rowSize:=newRowCount + nonBodyRowCount)
    
    Set newRowRange = lstObj.Range.Worksheet.Cells(lastRealDataRow + 1, lstObj.Range.column)
    Set newRowRange = newRowRange.Resize(rowSize:=addRowCount, ColumnSize:=lstObj.ListColumns.Count)

Finalize:
    On Error Resume Next
    
    If Not failed Then
        Set ResizeListObjectRows = newRowRange
    End If
    
    Set newRowRange = Nothing
    
    Exit Function
E:
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
        HeaderRangeRows = lstObj.HeaderRowRange.Rows.Count
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
                TotalsRangeRows = lstObj.TotalsRowRange.Rows.Count
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
            ReplaceListColFormulaWithStatic lstObj, lc.index
        End If
    Next lc
    Set lc = Nothing
End Function

Public Function ReplaceListColFormulaWithStatic(lstObj As ListObject, column As Variant)
'   Replaces listColumn formula with static values
'   Column = Name of ListColumn or Index of ListColumn
    
    Dim tArr As Variant
    If Not lstObj.ListColumns(column).DataBodyRange Is Nothing Then
        If lstObj.ListColumns(column).DataBodyRange(1, 1).HasFormula Then
            tArr = ListColumnAsArray(lstObj, lstObj.ListColumns(column).Name)
            If ArrDimensions(tArr) = 2 Then
                lstObj.ListColumns(column).DataBodyRange.ClearContents
                lstObj.ListColumns(column).DataBodyRange.Value = tArr
            End If
        End If
    End If
    
    If ArrDimensions(tArr) > 0 Then Erase tArr
    
    
End Function

Public Function CreateListColumnFormula(lstObj As ListObject, lstColName As String, _
    r1c1Formula As String, _
    Optional createColumnIfMissing As Boolean = True, _
    Optional convertToValues As Boolean = False, _
    Optional numberFormat As String = vbNullString) As Boolean

'   Create a formula in ListColumn, and then optionally replace the ListColumn contents with the values from the Formula
'   Note: Sets the 'Formula2R1C1' Formula Property.
On Error GoTo E:
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
E:
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
    Dim tmpARR As Variant
    Dim tmpAI As ArrInformation
    
    If lstObj.listRows.Count = 0 Then
        Exit Function
    End If
    tmpARR = ArrListCols(lstObj, aoNone, colName)
    tmpAI = ArrayInfo(tmpARR)
    If tmpAI.Dimensions = 2 Then
        ListColumnAsArray = tmpARR
    End If
    If tmpAI.Dimensions > 0 Then
        Erase tmpARR
    End If
End Function


Public Function FindMatchingListColumns(colName As String, Optional matchType As strMatchEnum = strMatchEnum.smEqual) As Collection
'   for any matching column, return collection of array {[listObjectName], [listColumnName]}

    Dim retCol As New Collection

    Dim tWS As Worksheet, tLO As ListObject, tCol As listColumn
    For Each tWS In ThisWorkbook.Worksheets
        For Each tLO In tWS.ListObjects
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
            ListColumnIndex = lstCol.index
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
    If hasData(lstObj) = False Then
        CountBlanks = 0
        Exit Function
    End If
    Set rng = lstObj.ListColumns(listColumn).DataBodyRange.SpecialCells(xlCellTypeBlanks)
    If Not rng Is Nothing Then
        CountBlanks = rng.Rows.Count
    End If
    Set rng = Nothing
    If Err.number <> 0 Then Err.Clear

End Function

Public Function ClearFormulas(lstObj As ListObject, listCol As Variant, Optional ByVal showMsg As Boolean = False) As Long
On Error Resume Next
    
    Dim fmlaRng As Range, retV As Long
    If hasData(lstObj) Then
        Set fmlaRng = lstObj.ListColumns(listCol).DataBodyRange
        Set fmlaRng = fmlaRng.SpecialCells(xlCellTypeFormulas)
        If Err.number <> 0 Then
            Err.Clear
        Else
            If Not fmlaRng Is Nothing Then
                retV = fmlaRng.Count
                fmlaRng.Value = Empty
            End If
        End If
    End If
        
    ClearFormulas = retV
    If showMsg And retV > 0 Then
        MsgBox_FT Concat(retV, " values were removed from ", lstObj.Range.Worksheet.Name, "[", lstObj.Name, "].[", lstObj.ListColumns(listCol).Name, "] becase formulas are not allowed to be manually entered anywere in the Financial Tool"), vbCritical + vbOKOnly, "OOPS"
        DoEvents
    End If

    

End Function

Public Function hasData(lstObj As Variant) As Boolean
On Error Resume Next
    If TypeName(lstObj) = "ListObject" Then
        hasData = lstObj.listRows.Count > 0
    ElseIf TypeName(lstObj) = "String" Then
        hasData = wt(CStr(lstObj)).listRows.Count > 0
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
            Set tmpRange = tmpRange.offSet(rowOffset:=HeaderRangeRows(srcListobj)).Resize(rowSize:=tmpRange.Rows.Count - HeaderRangeRows(srcListobj))
        End If
        
        If includeTotalsRow = False And .ShowTotals Then
            Set tmpRange = tmpRange.Resize(rowSize:=tmpRange.Rows.Count - TotalsRangeRows(srcListobj))
        End If
    End With

    Set ListColumnRange = tmpRange
    
    Set tmpRange = Nothing

End Function

' *** GET NEW RANGE AREA OF RESIZED LIST OBJECT
Public Function NewRowsRange(ByRef lstObj As ListObject, addRowCount As Long) As Range
On Error GoTo E:
    
    Dim evts As Boolean, prot As Boolean
    evts = Application.EnableEvents
    prot = lstObj.Range.Worksheet.ProtectContents
    
    
    If addRowCount <= 0 Then Exit Function
    
    Application.EnableEvents = False
    If prot Then
        UnprotectSht lstObj.Range.Worksheet
    End If
    
    pbRange.ClearFilter lstObj
    
    Dim hdrRngRows As Long, totRngRows As Long, listRows As Long
    hdrRngRows = lstObj.HeaderRowRange.Rows.Count
    If lstObj.TotalsRowRange Is Nothing Then
        totRngRows = 0
    Else
        totRngRows = lstObj.TotalsRowRange.Rows.Count
    End If
    listRows = lstObj.listRows.Count
    
    Dim endCount As Long
    endCount = lstObj.listRows.Count + addRowCount
    
    lstObj.Resize lstObj.Range.Resize(rowSize:=(hdrRngRows + totRngRows + listRows) + addRowCount)
    Do While lstObj.listRows.Count < endCount
        lstObj.listRows.Add
    Loop
    
    Dim firstNewRow As Long, lastNewRow As Long
    firstNewRow = lstObj.listRows.Count - addRowCount + 1
    lastNewRow = firstNewRow + (addRowCount - 1)
        
    Set NewRowsRange = lstObj.listRows(firstNewRow).Range.Resize(rowSize:=addRowCount)
    
    
Finalize:
    On Error Resume Next
        If prot Then
            ProtectSht lstObj.Range.Worksheet
        End If
        Application.EnableEvents = evts
    If Err.number <> 0 Then Err.Clear
    Exit Function
E:
    ErrorCheck
    Resume Finalize:
    

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
Function getFirstRow(ByVal lo As ListObject, headers, Values, Optional cachedDict As Dictionary, Optional forceRebuild As Boolean = False) As Long
   
'    forceRebuild = True
   
   Dim chi As Long
   For chi = LBound(headers) To UBound(headers)
    If IsNumeric(headers(chi)) Then headers(chi) = lo.ListColumns(headers(chi)).Name
   Next chi
  
    If Not cachedDict Is Nothing Then
        Dim vID As String: vID = Join(Values, "|") & "|"
        If cachedDict.Exists(vID) Then
            getFirstRow = cachedDict(vID).Item(1)
        Else
            getFirstRow = 0
        End If
        Exit Function
    Else
       Static sID As String: If sID = "" Then sID = lo.Name & "|" & Join(headers, "|")
       Static oIndex As Dictionary
        If forceRebuild Or oIndex Is Nothing Or sID <> (lo.Name & "|" & Join(headers, "|")) Then
            Set oIndex = getIndex(lo, headers)
        End If
        If oIndex.Exists(Join(Values, "|") & "|") Then
            getFirstRow = oIndex(Join(Values, "|") & "|").Item(1)
        End If
    End If
  
End Function

'Obtain a lookup dictionary which helps you find rows matching a set of headers.
'@param {ListObject} The listobject to search within.
'@param {Array<String>} Headers to create a lookup for
'@returns {Dictionary<string, Collection<Long>>} Lookup to find rows matching criteria
'Private Function getIndex(ByVal lo As ListObject, headers As Variant) As Object
Public Function getIndex(ByVal lo As ListObject, headers As Variant) As Dictionary
On Error GoTo E:

   Dim chi As Long
   For chi = LBound(headers) To UBound(headers)
    If IsNumeric(headers(chi)) Then headers(chi) = lo.ListColumns(headers(chi)).Name
   Next chi
  
  Dim v: v = lo.Range.Value
  Dim iLB As Long: iLB = LBound(headers)
  Dim iUB As Long: iUB = UBound(headers)
  
  'Create field indices
  Dim iHeaders() As Long: ReDim iHeaders(iLB To iUB)
  Dim iUBH As Long: iUBH = UBound(v, 2)
  Dim i As Long
  For i = iLB To iUB
    Dim j As Long
    For j = 1 To iUBH
      If headers(i) = v(1, j) Then
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
E:
    ftBeep btError
    Stop
  Resume
End Function

Public Function ListRowIdxFromWksht(lstObj As ListObject, worksheetRow As Long) As Long
    Dim hdrRow As Long
    hdrRow = lstObj.HeaderRowRange.Row + (1 - lstObj.HeaderRowRange.Rows.Count)
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

Public Function CopyTableToNewWorkbook(lstObj As ListObject, Optional newListObjName) As ListObject
On Error GoTo E:
    Dim failed As Boolean
    Dim evts As Boolean: evts = Events
    EventsOff
    Dim wk As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Set wk = Workbooks.Add
    Set ws = wk.Worksheets(1)
    Dim arrData() As Variant, tarRng As Range
    arrData = ArrListObj(lstObj, aoIncludeListObjHeaderRow)
    Set tarRng = ws.Range("A1")
    Set tarRng = tarRng.Resize(rowSize:=lstObj.Range.Rows.Count, ColumnSize:=lstObj.Range.Columns.Count)
    tarRng.Value = arrData
    Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tarRng, XlListObjectHasHeaders:=xlYes)
    If Not IsMissing(newListObjName) Then
        lo.Name = CStr(newListObjName)
    End If
    Dim sc As listColumn, tc As listColumn
    For Each sc In lstObj.ListColumns
        For Each tc In lo.ListColumns
            If StringsMatch(tc.Name, sc.Name) Then
                If Not IsNull(sc.DataBodyRange.HorizontalAlignment) Then
                    tc.DataBodyRange.HorizontalAlignment = sc.DataBodyRange.HorizontalAlignment
                End If
                If Not IsNull(sc.DataBodyRange.VerticalAlignment) Then
                    tc.DataBodyRange.VerticalAlignment = sc.DataBodyRange.VerticalAlignment
                End If
                If Not IsNull(sc.DataBodyRange.numberFormat) Then
                    tc.DataBodyRange.numberFormat = sc.DataBodyRange.numberFormat
                End If
            End If
        Next tc
    Next sc
    lo.Range.EntireColumn.AutoFit
Finalize:
    On Error Resume Next
    Events = evts
    If Not failed Then
        Set CopyTableToNewWorkbook = lo
    End If
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:
End Function

Public Function RemoveTableColumns(lstObj As ListObject, ParamArray colNames() As Variant)
    Dim msg As String
    msg = ConcatWithDelim(" ", "You are about to permanently remove ListColumns from the *", lstObj.Name, "* table in the *", lstObj.Range.Worksheet.Parent.Name, "* workbook.", vbNewLine, "Continue?")
    If MsgBox_FT(msg, vbYesNo + vbQuestion + vbDefaultButton2, "DELETE LIST COLUMNS") = vbYes Then
    
        Dim i As Long, isInvalid As Boolean
        For i = LBound(colNames) To UBound(colNames)
            If StringsMatch(TypeName(colNames(i)), "String") = False Then
                isInvalid = True
                Exit For
            End If
        Next i
        If isInvalid Then
            Err.Raise 1004, Source:="pbListObj.RemoveTableColumns", Description:="colNames must be actual column names in list object"
            Exit Function
        End If
        Dim cName
        For Each cName In colNames
            If ListColumnExists(lstObj, CStr(cName)) Then
                lstObj.ListColumns(cName).Delete
            End If
        Next cName
        lstObj.Range.EntireColumn.AutoFit
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Creates a new 'Activity' Column for [lstObj], and finds all
'   ListColumns that are formatted as Date/DateTime, and
'   creates a list of ColumnNames with datetime value from
'   found datetime listcolumns
'   Results in new ListColumn are ordered by Date Asc
'   Example:  If ListObject contained 'CreateDt','StatusDt'
'   ListColumns (formatted as datetime), the new Activity
'   Column might look like the following for a given row:
'       CreateDt - 04/20/23 08:15:31 AM
'       StatusDt - 04/21/23 04:13:05 PM
'   Any ListColumns you wish to be EXCLUDED from summary
'   can be included in the [colNamesExclude] Argument
'   If [activityColumnName] already exists, all values for that
'   column will be cleared before summarization
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'Public Function AddDateSummaryListColumn( _
'    lstObj As ListObject, _
'    copyToNewWkbk As Boolean, _
'    activityColumnName As String, _
'    ParamArray colNamesExclude() As Variant)
'
'    If lstObj.listRows.Count = 0 Then
'        Err.Raise 1004, Source:="pbListObj.AddDateSummaryListColumn", Description:="ListObject does not contain any rows"
'    End If
'
'    Dim lo As ListObject
'    If copyToNewWkbk Then
'        Set lo = CopyTableToNewWorkbook(lstObj)
'    Else
'        Set lo = lstObj
'    End If
'    If ListColumnExists(lo, activityColumnName) Then
'        lo.ListColumns(activityColumnName).DataBodyRange.ClearContents
'    Else
'        AddColumnIfMissing lo, activityColumnName, 1
'    End If
'
'
'
'End Function


