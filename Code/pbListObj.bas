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

'   ~~~ ~~~ Sort's a List Object by [lstColIdx] Ascending.  If already
'                   sorted Ascending, sorts Descending
Public Function UserSort(lstObj As ListObject, lstColIdx As Long)
        If lstObj.listRows.count = 0 Then Exit Function
        If lstObj.Range.Worksheet.Protection.AllowSorting = False Then
            MsgBox_FT "This table does not allow custom sorting", vbOKOnly + vbInformation, "SORRY"
            Exit Function
        End If
        Dim clearPreviousSort As Boolean, hdrCol As Long
        Dim orderBy As XlSortOrder
        orderBy = xlAscending
        hdrCol = lstObj.HeaderRowRange.column
        If lstObj.Sort.SortFields.count > 1 Then
            clearPreviousSort = True
        Else
            If lstObj.Sort.SortFields.count = 1 Then
                If lstObj.Sort.SortFields(1).Key.column - hdrCol + 1 <> lstColIdx Then
                    clearPreviousSort = True
                End If
            End If
        End If
        If clearPreviousSort Then
            lstObj.Sort.SortFields.Clear
        End If
        With lstObj.Sort
            If .SortFields.count = 1 Then
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
    If Err.Number <> 0 Then Err.Clear
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
    If totalRowCount > 0 And totalRowCount < lstObj.listRows.count Then
        Err.Raise 17
        'RaiseError ERR_LIST_OBJECT_RESIZE_CANNOT_DELETE, "Resizing Failed because new 'totalRowCount' is less than existing row count"
    End If
    If totalRowCount > 0 And addRowCount > 0 Then
        Err.Raise 17
        'RaiseError ERR_LIST_OBJECT_RESIZE_INVALID_ARGUMENTS, "Resizing Failed, cannot set totalRowCount AND addRowCount"
    End If
        
    If addRowCount > 0 Then
        newRowCount = lstObj.listRows.count + addRowCount
    Else
        addRowCount = totalRowCount - lstObj.listRows.count
        newRowCount = totalRowCount
    End If
    '   Include Header range and TotalsRange (if applicable) in overall ListObject Range Size
    nonBodyRowCount = HeaderRangeRows(lstObj) + TotalsRangeRows(lstObj)
        
    lastRealDataRow = lstObj.HeaderRowRange.Row + lstObj.HeaderRowRange.Rows.count - 1
    If lstObj.listRows.count > 0 Then lastRealDataRow = lastRealDataRow + lstObj.listRows.count
    
    '   Resize ListObject Range with new range
    lstObj.Resize lstObj.Range.Resize(rowSize:=newRowCount + nonBodyRowCount)
    
    Set newRowRange = lstObj.Range.Worksheet.Cells(lastRealDataRow + 1, lstObj.Range.column)
    Set newRowRange = newRowRange.Resize(rowSize:=addRowCount, ColumnSize:=lstObj.ListColumns.count)

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
    MsgBox "Error: " & Err.Number & ", " & Err.Description
    'ErrorCheck
    If Err.Number <> 0 Then Err.Clear
    Resume Finalize:

End Function


    Public Function HeaderRangeRows(lstObj As ListObject) As Long
    ' *** Returns -1 for Error, other Num Rows in HeaderRowRange
    On Error Resume Next
        HeaderRangeRows = lstObj.HeaderRowRange.Rows.count
        If Err.Number <> 0 Then
            HeaderRangeRows = -1
        End If
        If Err.Number <> 0 Then Err.Clear
    End Function
    
    Public Function TotalsRangeRows(lstObj As ListObject) As Long
    ' *** Returns -1 for Error, other Num Rows in HeaderRowRange
    On Error Resume Next
        If Not lstObj.TotalsRowRange Is Nothing Then
            If lstObj.ShowTotals Then
                TotalsRangeRows = lstObj.TotalsRowRange.Rows.count
            End If
        End If
        If Err.Number <> 0 Then
            TotalsRangeRows = -1
        End If
        If Err.Number <> 0 Then Err.Clear
    End Function




Public Function ReplaceFormulasWithStatic(lstObj As ListObject) As Boolean
'   REPLACES ALL FORMULAS IN LIST COLUMS, WITH THE VALUES
'   Helpful for situations like creating static copies of tables/listobjects
    If lstObj.listRows.count = 0 Then Exit Function
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
On Error GoTo E:
    Dim failed As Boolean
    If lstObj.listRows.count = 0 Then
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
    
    If Err.Number <> 0 Then Err.Clear
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
    
    If lstObj.listRows.count = 0 Then
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

Public Function ListColumnExists(ByRef lstObj As ListObject, lstColName As String) As Boolean
'   Return true If 'lstColName' is a valid ListColumn in 'lstObj'
    ListColumnExists = ListColumnIndex(lstObj, lstColName) > 0
End Function

Public Function ListColumnIndex(ByRef lstObj As ListObject, lstColName As String) As Long
    Dim retV As Long
    Dim colIDX As Long
    For colIDX = 1 To lstObj.ListColumns.count
        If StrComp(LCase(lstColName), LCase(lstObj.ListColumns(colIDX).Name), vbTextCompare) = 0 Then
            retV = colIDX
            Exit For
        End If
    Next colIDX
    ListColumnIndex = retV
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
        If lstObj.listRows.count > 0 And numberFormat <> vbNullString Then
            nc.DataBodyRange.numberFormat = numberFormat
        End If
    End If
    AddColumnIfMissing = (Err.Number = 0)
    If Err.Number <> 0 Then Err.Clear
    Set nc = Nothing
    
    If Err.Number <> 0 Then Err.Clear
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
        CountBlanks = rng.Rows.count
    End If
    Set rng = Nothing
    If Err.Number <> 0 Then Err.Clear

End Function

Public Function hasData(lstObj As Variant) As Boolean
On Error Resume Next
    If TypeName(lstObj) = "ListObject" Then
        hasData = lstObj.listRows.count > 0
    ElseIf TypeName(lstObj) = "String" Then
        hasData = wt(CStr(lstObj)).listRows.count > 0
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function ListColumnRange(srcListobj As ListObject, lstColumn As Variant, Optional includeHeaderRow As Boolean = False, Optional includeTotalsRow As Boolean = False) As Range

    Dim tmpRange As Range, tmpRowCount As Long
    
    With srcListobj
        tmpRowCount = .listRows.count
        If includeHeaderRow Then tmpRowCount = tmpRowCount + HeaderRangeRows(srcListobj)
        If includeTotalsRow And .ShowTotals Then tmpRowCount = tmpRowCount + TotalsRangeRows(srcListobj)
        
        If tmpRowCount = 0 Then Exit Function
        
        Set tmpRange = srcListobj.ListColumns(lstColumn).Range
    
        If includeHeaderRow = False Then
            Set tmpRange = tmpRange.Offset(rowOffset:=HeaderRangeRows(srcListobj)).Resize(rowSize:=tmpRange.Rows.count - HeaderRangeRows(srcListobj))
        End If
        
        If includeTotalsRow = False And .ShowTotals Then
            Set tmpRange = tmpRange.Resize(rowSize:=tmpRange.Rows.count - TotalsRangeRows(srcListobj))
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
        pbUnprotectSheet lstObj.Range.Worksheet
    End If
    
    pbRange.ClearFilter lstObj
    
    Dim hdrRngRows As Long, totRngRows As Long, listRows As Long
    hdrRngRows = lstObj.HeaderRowRange.Rows.count
    If lstObj.TotalsRowRange Is Nothing Then
        totRngRows = 0
    Else
        totRngRows = lstObj.TotalsRowRange.Rows.count
    End If
    listRows = lstObj.listRows.count
    
    Dim endCount As Long
    endCount = lstObj.listRows.count + addRowCount
    
    lstObj.Resize lstObj.Range.Resize(rowSize:=(hdrRngRows + totRngRows + listRows) + addRowCount)
    Do While lstObj.listRows.count < endCount
        lstObj.listRows.Add
    Loop
    
    Dim firstNewRow As Long, lastNewRow As Long
    firstNewRow = lstObj.listRows.count - addRowCount + 1
    lastNewRow = firstNewRow + (addRowCount - 1)
        
    Set NewRowsRange = lstObj.listRows(firstNewRow).Range.Resize(rowSize:=addRowCount)
    
    
Finalize:
    On Error Resume Next
        If prot Then
            pbProtectSheet lstObj.Range.Worksheet
        End If
        Application.EnableEvents = evts
    If Err.Number <> 0 Then Err.Clear
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
Function getFirstRow(ByVal lo As ListObject, headers, values, Optional cachedDict As Dictionary, Optional forceRebuild As Boolean = False) As Long
   
   Dim chi As Long
   For chi = LBound(headers) To UBound(headers)
    If IsNumeric(headers(chi)) Then headers(chi) = lo.ListColumns(headers(chi)).Name
   Next chi
  
    If Not cachedDict Is Nothing Then
        Dim vID As String: vID = Join(values, "|") & "|"
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
        If oIndex.Exists(Join(values, "|") & "|") Then
            getFirstRow = oIndex(Join(values, "|") & "|").Item(1)
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
  
  Dim v: v = lo.Range.value
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
    Beep
    Stop
  Resume
End Function
