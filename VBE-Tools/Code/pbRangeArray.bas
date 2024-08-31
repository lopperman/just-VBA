Attribute VB_Name = "pbRangeArray"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' pbRangeArray v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' General  Helper Utilities for Working with Arrays and Ranges
'
' @module pbRangeArray
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
'   KEY FUNCTIONS IN THIS MODULE
'
'   * NOTE:  Some Functions Require an Empty Worksheet to manipulate data. If Worksheet name "pbTmpUtilSheet"
'       does not exist, it will be created with 'VeryHidden' .Visible Property
'
'   NOTE:  Any Arr[X] Can use any of the following options, passed in with the ArrayOptionFlags Parameter
'       aoNone = (no additional changes)
'       aoUnique = (returns unique combinations -- row based)
'       aoUniqueNoSort = (Ignores sorting in final unique list)
'       aoAreaSizesMustMatch = (requires Rows OR Columns to match for all Areas in Range)
'       aoVisibleRangeOnly = (if source of array is range, limit array to visible range(s) only)
'
'   ~~~ GET ARRAY (Arr[X] Functions) ~~~
'   ArrRange: (get 1-based, minimum 2D array from Range (supports multiple Range Areas), with options available)
'       Note: multiple range areas where Rows do not match, ** AND ** Columns do not match may cause error
'   ArrArray: (converts any array to 1-based, minimum 2D array, with options available)
'   ArrParams: (converts 'ParamArray' argument to 1-based, 2D array)
'   ArrListObj: (get 1-based, 2D array from entire ListObject DataBodyRange, with options available)
'   ArrListCols: (get 1-based, 2D array from specified ListObject list columns, with options available)
'
'   ~~~ GET ARRAY (Other) ~~~
'   GetUniqueSortedListCol: (Returns unique 1-based, 2D array from specific ListObject ListColumn)
'       ReturnType = Array (default), Dictionary, or Collection
'   RangeTo1DArray: (return all cells in Range as 1-based,** 1 Dimensional ** array)
'
'   ~~~ INFORMATIONAL ~~~
'   ArrayInfo: (returns information about array)
'   ArrDimensions: (return number of dimensions for any array)
'   IsArrInit: (returns true if array is initialized with data)
'   RangeArea: (return info about Range area; allows 1 Area ONLY)
'   RangeInfo: (return information about Range; summarizes ** ALL ** range areas)

Option Explicit
Option Compare Text
Option Base 1

Private Const TMP_RANGE_UTIL_WORKSHEET As String = "pbTmpUtilSheet"

Private Function TmpUtilSheet() As Worksheet
    
    If WorksheetExists(TMP_RANGE_UTIL_WORKSHEET) = False Then
        Dim retWS As Worksheet
        
        Set retWS = Excel.Application.Worksheets.add(After:=Worksheets(Worksheets.Count))
        retWS.Name = TMP_RANGE_UTIL_WORKSHEET
        retWS.visible = xlSheetVeryHidden
        Set TmpUtilSheet = retWS
        Set retWS = Nothing
        
    Else
        Set TmpUtilSheet = ThisWorkbook.Worksheets(TMP_RANGE_UTIL_WORKSHEET)
    End If
End Function

Public Function RangeToUniqueArray(rng As Range, Optional Sorted As Boolean = True, Optional excludeEmpty As Boolean = True) As Variant()
    Dim d As New Dictionary
    d.CompareMode = TextCompare
    Dim i As Long, rngArea, rngCell, tmpVal
    For Each rngArea In rng.Areas
        For Each rngCell In rngArea
            tmpVal = rngCell.value
            If StringsMatch(TypeName(tmpVal), "String") Then tmpVal = Trim(tmpVal)
            If Not d.Exists(tmpVal) Then
                If excludeEmpty = True Then
                    If Not StringsMatch(TypeName(tmpVal), "Empty") Then
                        d.add tmpVal, tmpVal
                    End If
                Else
                        d.add tmpVal, tmpVal
                End If
            End If
        Next rngCell
    Next rngArea
    If d.Count > 0 Then
        If Sorted And d.Count > 1 Then
            If TmpUtilSheet.EnableAutoFilter = False Then TmpUtilSheet.EnableAutoFilter = True
            Dim sRng As Range
            TmpUtilSheet.Cells.Clear
            DoEvents
            Set sRng = TmpUtilSheet.Range("A1")
            Set sRng = sRng.Resize(RowSize:=d.Count)
            sRng.value = ArrArray(d.Keys, aoNone)
            
            TmpUtilSheet.Sort.SortFields.Clear
            TmpUtilSheet.Sort.SortFields.add2 Key:=Range( _
                sRng.Address), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                xlSortNormal
            With TmpUtilSheet.Sort
                .SetRange Range(sRng.Address)
                .Header = xlNo
                .MatchCase = False
                .orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With

            RangeToUniqueArray = RangeToUniqueArray(sRng, Sorted:=False)
            TmpUtilSheet.Cells.Clear
        Else
            RangeToUniqueArray = d.Keys
        End If
    End If
End Function

Public Function ArrRange(rng As Range, flags As ArrayOptionFlags) As Variant
    Dim retArray As Variant
    Dim unique As Boolean
    Dim aInfo As ArrInformation
    Dim tmpValue As Variant
    Dim RngInfo As RngInfo
    RngInfo = RangeInfo(rng)
    If EnumCompare(flags, aoAreaSizesMustMatch) Then
        If RngInfo.AreasSameColumns = False And RngInfo.AreasSameRows = False Then
            RaiseError ERR_INVALID_RANGE_SIZE, "Range Areas must all be the same size (ftRangeArray.ArrRange)"
        End If
    End If
    unique = EnumCompare(flags, ArrayOptionFlags.aoUnique + ArrayOptionFlags.aoUniqueNoSort, ecOR)
   If EnumCompare(flags, ArrayOptionFlags.aoVisibleRangeOnly) Then
        retArray = BuildRC1(rng.SpecialCells(xlCellTypeVisible))
   Else
        retArray = BuildRC1(rng)
   End If
    If unique Then
        retArray = UniqueRC1Arr(retArray, flags)
    End If
    
    aInfo = ArrayInfo(retArray)
    
    If aInfo.Dimensions = 1 Then
        retArray = ConvertArrToRCArr(retArray)
    Else
        If aInfo.Dimensions = 0 Then
            If Not IsEmpty(retArray) Then
                tmpValue = retArray
                'a single value was returned, convert to RC1 array
                ReDim retArray(1 To 1, 1 To 1)
                retArray(1, 1) = tmpValue
            Else
                'Will Return Empty
            End If
        End If
    End If
    
    
    ArrRange = retArray
    
    If ArrDimensions(retArray) > 0 Then
        Erase retArray
    End If
    
    

End Function
Public Function ArrListObj(lstObj As ListObject, flags As ArrayOptionFlags) As Variant
'   Returns 1-based, 2D array from entire ListObject Data Body Range, with options
' (don't need to deal with state here)
    If EnumCompare(flags, ArrayOptionFlags.aoIncludeListObjHeaderRow) Then
        If lstObj.ShowTotals = True Then
            Dim rng As Range
            Set rng = lstObj.Range
            Set rng = rng.Resize(RowSize:=rng.rows.Count - 1)
            ArrListObj = ArrRange(rng, flags)
        Else
            ArrListObj = ArrRange(lstObj.Range, flags)
        End If
    Else
        ArrListObj = ArrRange(lstObj.DataBodyRange, flags)
    End If
End Function

Public Function RangeInfo(rg As Range) As RngInfo
    Dim retV As RngInfo
    If rg Is Nothing Then
        retV.rows = 0
        retV.Columns = 0
        retV.AreasSameRows = False
        retV.AreasSameColumns = False
        retV.Areas = 0
    Else
        retV.rows = RangeRowCount(rg)
        retV.Columns = RangeColCount(rg)
        retV.AreasSameRows = ContiguousRows(rg)
        retV.AreasSameColumns = ContiguousColumns(rg)
        retV.Areas = rg.Areas.Count
    End If
    RangeInfo = retV
End Function

Public Function RangeArea(rg As Range) As AreaStruct
'   Return info about Range area
'   Range with 1 Area allowed, otherwise error
    If rg.Areas.Count <> 1 Then
        RaiseError ERR_RANGE_AREA_COUNT, "Range Area Count <> 1 (ftRangeArray.RangeArea)"
    End If
    
    Dim retV As AreaStruct
    retV.rowStart = rg.Row
    retV.RowEnd = rg.Row + rg.rows.Count - 1
    retV.ColStart = rg.column
    retV.ColEnd = rg.column + rg.Columns.Count - 1
    retV.rowCount = rg.rows.Count
    retV.columnCount = rg.Columns.Count
    
    RangeArea = retV

End Function

Public Function RangeTo1DArray(ByVal rng As Range) As Variant()
'TODO:  Optimizae to build 1D array from Arrays from each Area in rng
    

'   Return all cells in Range as 1D Array
    Dim retV() As Variant
    ''BASE 1
    ReDim retV(1 To rng.Count)
    Dim cl As Range, clIDX As Long
    clIDX = 1
    For Each cl In rng.Cells
        retV(clIDX) = cl.value
        clIDX = clIDX + 1
    Next cl
    RangeTo1DArray = retV
    
    
End Function

Public Function GetUniqueSortedListCol(lstObj As ListObject, lstCol As Variant, Optional returnType As ListReturnType = ListReturnType.lrtArray) As Variant
'   Returns unique 1-based, 2D array from specific ListObject ListColumn
'   Return Type = Array (default), Dictionary, or Collection
    If lstObj.listRows.Count = 0 Then Exit Function
    
    
    Dim tDic As Dictionary
    Dim tCol As Collection
    
    Dim aIdx As Long, arr As Variant
    arr = ArrListCols(lstObj, aoUnique, lstCol)
    
    Select Case returnType
        Case ListReturnType.lrtArray
            GetUniqueSortedListCol = arr
        
        Case ListReturnType.lrtDictionary
            Set tDic = New Dictionary
            For aIdx = LBound(arr) To UBound(arr)
                tDic(arr(aIdx, 1)) = arr(aIdx, 1)
            Next aIdx
            Set GetUniqueSortedListCol = tDic
        
        Case ListReturnType.lrtCollection
            Set tDic = New Collection
            For aIdx = LBound(arr) To UBound(arr)
                tCol.add arr(aIdx, 1)
            Next aIdx
            Set GetUniqueSortedListCol = tCol
    End Select
    
    Set tDic = Nothing
    Set tCol = Nothing
    
    If ArrDimensions(arr) > 0 Then
        Erase arr
    End If
    
    
    
End Function
Public Function ArrListCols(lstObj As ListObject, flags As ArrayOptionFlags, ParamArray listCols() As Variant) As Variant
'   Get Array from specific ListObject listColum(s)
    

    Dim IDX As Long, rng As Range, inclHeader As Boolean
    inclHeader = EnumCompare(flags, aoIncludeListObjHeaderRow)
    If lstObj.listRows.Count > 0 Then
        For IDX = LBound(listCols) To UBound(listCols)
            If rng Is Nothing Then
                If inclHeader Then
                    Set rng = lstObj.HeaderRowRange(1, lstObj.ListColumns(listCols(IDX)).Index)
                    Set rng = rng.Resize(RowSize:=lstObj.listRows.Count + lstObj.HeaderRowRange.rows.Count)
                Else
                    Set rng = lstObj.ListColumns(listCols(IDX)).DataBodyRange
                End If
            Else
                If inclHeader Then
                    Dim tRng As Range
                    Set tRng = lstObj.HeaderRowRange(1, lstObj.ListColumns(listCols(IDX)).Index)
                    Set tRng = tRng.Resize(RowSize:=lstObj.listRows.Count + lstObj.HeaderRowRange.rows.Count)
                    Set rng = Union(rng, tRng)
                    Set tRng = Nothing
                Else
                    Set rng = Union(rng, lstObj.ListColumns(listCols(IDX)).DataBodyRange)
                End If
            End If
        Next IDX
        ArrListCols = ArrRange(rng, flags)
    End If
    Set rng = Nothing
    
    
End Function

Public Function ArrParams(ParamArray vals() As Variant) As Variant
'   Build standard array from ParamsArray so it can be passed as Variant() to other functions
    ArrParams = ArrArray(vals, aoNone)
    
'    If IsMissing(vals) Or UBound(vals) = -1 Then
'        'return empty array
'        ArrParams = Array()
'        Exit Function
'    ElseIf LBound(vals) = 0 And UBound(vals) = 0 Then
'        If VarType(vals(0)) = vbArray + vbVariant Then
'            If UBound(vals(0)) = -1 Then
'                ArrParams = Array()
'                Exit Function
'            End If
'        End If
'    End If
'    'NEED TO CHECK FOR ARR(0) = 0 TO -1
'    Dim tmp As Variant, vIDX As Long, offset As Long
'    If LBound(vals) = UBound(vals) Then
'        If ArrDimensions(vals(LBound(vals))) > 0 Then
'            tmp = ArrArray(vals(LBound(vals)), aoNone)
'        Else
'            ReDim tmp(1 To 1, 1 To 1)
'            tmp(1, 1) = vals(LBound(vals))
'        End If
'    Else
'        If LBound(vals) = 0 Then offset = 1
'        ReDim tmp(1 To (UBound(vals) - LBound(vals) + 1), 1 To 1)
'        For vIDX = LBound(vals) To UBound(vals)
'            tmp(vIDX + offset, 1) = vals(vIDX)
'        Next vIDX
'    End If
'
'    ArrParams = tmp
'    If ArrDimensions(tmp) > 0 Then Erase tmp

End Function
Public Function ArrArray(ByVal arr As Variant, flags As ArrayOptionFlags, Optional zeroBasedAsColumns As Boolean = False) As Variant
'   By default, a zero-based array will become multiple rows.  Set 'zeroBasedAsColumns' to create 1 row with multiple columns
    Dim retArray As Variant
    Dim unique As Boolean
    
    '   CHECK TO DETERMINE IF 'ARR' IS FROM A PARAMARRAY -- WHICH MEANS WE SHOULD TAKE 'ARR(0)' AS THE INPUT ARRAY
    If ValidArray(arr) And LBound(arr) = UBound(arr) Then
        If EnumCompare(VarType(arr(LBound(arr))), VbVarType.vbArray) Then
            If UBound(arr) >= LBound(arr) Then
                arr = arr(LBound(arr))
            End If
        End If
    End If
    
    unique = EnumCompare(flags, ArrayOptionFlags.aoUnique)

    If ArrDimensions(arr) = 1 Then
        retArray = ConvertArrToRCArr(arr, zeroBasedAsColumns)
    Else
        retArray = arr
    End If
    Dim ai As ArrInformation
    ai = ArrayInfo(retArray)
    If unique Then
        If ai.Dimensions = 0 Then
            Err.Raise 427, Description:="Array not initialized"
        End If
        If ai.rows = 1 And ai.Columns = 1 Then
            'We're Good
        Else
            retArray = UniqueRC1Arr(retArray, flags)
        End If
    End If
    If ArrDimensions(retArray) = 1 Then
        retArray = ConvertArrToRCArr(retArray)
    Else
        If ArrDimensions(retArray) = 0 Then
            If UBound(retArray) >= 0 Then
                Dim tmpValue As Variant
                tmpValue = retArray
                'a single value was returned, convert to RC1 array
                ReDim retArray(1 To 1, 1 To 1)
                retArray(1, 1) = tmpValue
            Else
                ReDim retArray(1 To 1, 1 To 1)
                retArray(1, 1) = vbEmpty
            End If
        End If
    End If
    
    ArrArray = retArray
    
    If ArrDimensions(retArray) > 0 Then
        Erase retArray
    End If
    
    
End Function
Public Function IsArrInit(inpt As Variant) As Boolean
'   Returns True if Array is initialized and has data
    IsArrInit = ArrDimensions(inpt) > 0
End Function


'   ~~~ Test if anything is and ARRAY ~~~
Public Function ValidArray(tstArr As Variant) As Boolean
    Dim vt As Long: vt = VarType(tstArr)
    Dim compare As Long
    compare = vt And VbVarType.vbArray
    ValidArray = compare <> 0
End Function
    
'   ~~~ Check if array has been initialized  (can read or set values) ~~~
'   optionally raise an error if item passed in isn't an array
Public Function ArrayInitialized(tstArr As Variant, Optional errorIfNotArray As Boolean = False) As Boolean
On Error Resume Next
    If Not ValidArray(tstArr) Then
        If errorIfNotArray Then
            Err.Raise 427, Description:="ArrayInitialized - 'tstArr' Parameter was not of Type Array"
        Else
            ArrayInitialized = False
        End If
    Else
        Dim dimLen As Long
        dimLen = UBound(tstArr, 1) - LBound(tstArr, 1) + 1
        If Err.number <> 0 Then
            ArrayInitialized = False
        ElseIf UBound(tstArr, 1) < LBound(tstArr, 1) Then
            ArrayInitialized = False
        Else
            ArrayInitialized = True
        End If
    End If
    If Not Err.number = 0 Then Err.Clear
End Function
    

Public Function ArrayInfo(arr As Variant) As ArrInformation
'   Returns Information about array dimensions
'   Note: Use Arr[X] Functions in pbRangeArray (e.g. 'ArrRange', 'ArrArray', 'ArrListObject') to ensure all arrays
'       are 1-based, 2-dimensional - required for populating worksheet ranges in a 'table style rows/columns' convention
On Error Resume Next
    Dim tmp As ArrInformation
    tmp.isArray = ValidArray(arr)
    If tmp.isArray = False Then GoTo Finalize:
        
    If UBound(arr) = -1 Or LBound(arr) > UBound(arr) Then
        tmp.Dimensions = 0
    Else
        tmp.Dimensions = ArrDimensions(arr)
        If tmp.Dimensions > 0 Then
            tmp.LBound_first = LBound(arr, 1)
            tmp.Ubound_first = UBound(arr, 1)
            tmp.rows = (tmp.Ubound_first - tmp.LBound_first) + 1
        End If
        If tmp.Dimensions = 1 Then
            tmp.Columns = 1
        Else
            If tmp.Dimensions = 2 Then
                tmp.Columns = (UBound(arr, 2) - LBound(arr, 2)) + 1
            End If
        End If
        If tmp.Dimensions >= 2 Then
            tmp.LBound_second = LBound(arr, 2)
            tmp.UBound_second = UBound(arr, 2)
        End If
    End If
    
Finalize:
    ArrayInfo = tmp
    If Err.number <> 0 Then Err.Clear
End Function

Public Function ArrDimensions(ByRef checkArr As Variant) As Long
'   RETURNS Array Dimensions Count
'   RETURNS 0 'checkArr' argument is not an Array
'   Example Use:
'       If ArrDimensions(myArray) > 0 Then ... 'checkArr' is a valid array
On Error Resume Next
    Dim dimCount As Long
    If Not ValidArray(checkArr) Then
        GoTo Finalize:
    End If
    Do While Err.number = 0
        Dim tmp As Variant
        tmp = UBound(checkArr, dimCount + 1)
        If Err.number = 0 Then
            dimCount = dimCount + 1
        Else
            Err.Clear
            Exit Do
        End If
    Loop
    If dimCount > 0 Then
        If UBound(checkArr) < LBound(checkArr) Then
            dimCount = 0
        End If
    End If
    
Finalize:
    ArrDimensions = dimCount
    If Err.number <> 0 Then Err.Clear
End Function

' ***** PRIVATE FUNCTIONS ***** ' ***** PRIVATE FUNCTIONS ***** ' ***** PRIVATE FUNCTIONS ***** ' ***** PRIVATE FUNCTIONS *****
' ***** PRIVATE FUNCTIONS ***** ' ***** PRIVATE FUNCTIONS ***** ' ***** PRIVATE FUNCTIONS ***** ' ***** PRIVATE FUNCTIONS *****

Private Function UniqueRC1Arr(ByVal arr As Variant, flags As ArrayOptionFlags) As Variant
    
    Dim retArray As Variant
    Dim fixARR As Variant
    Dim aInfo As ArrInformation
    Dim retAI As ArrInformation
    Dim tmpRng As Range
    
    aInfo = ArrayInfo(arr)
    
    If aInfo.Dimensions = 0 Then
        ReDim fixARR(1 To 1, 1 To 1)
        fixARR(1, 1) = arr
        arr = fixARR
        aInfo = ArrayInfo(arr)
    End If
    
    If aInfo.Dimensions = 1 Or LBound(arr, 1) <= 0 Then
        arr = ConvertArrToRCArr(arr)
        aInfo = ArrayInfo(arr)
    End If
    
    ClearTempRangeUtil
    
    With TmpUtilSheet
        Set tmpRng = .Range("A1")
        Set tmpRng = tmpRng.Resize(RowSize:=aInfo.rows, ColumnSize:=aInfo.Columns)
        tmpRng.value = arr
        If Not EnumCompare(flags, ArrayOptionFlags.aoUniqueNoSort) Then
            Dim sidx As Long, sRng As Range
            .Sort.SortFields.Clear
            For sidx = 1 To tmpRng.Columns.Count
                Set sRng = tmpRng.Resize(ColumnSize:=1).Offset(ColumnOffset:=sidx - 1)
                .Sort.SortFields.add2 Key:=.Range(sRng.Address), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            Next sidx
            Set sRng = Nothing
            With .Sort
                .SetRange tmpRng
                .Header = xlNo
                .MatchCase = False
                .orientation = xlTopToBottom
                .SortMethod = xlPinYin
               .Apply
            End With
        End If
        retArray = WorksheetFunction.unique(tmpRng, False, False)
        retAI = ArrayInfo(retArray)
        If retAI.Columns < aInfo.Columns And retAI.Dimensions = 1 Then
            'got down to one column with > 1 value, so now it needs to be flipped
            Dim fixIDX As Long
            ReDim fixARR(1 To 1, 1 To retAI.rows)
            For fixIDX = 1 To retAI.rows
                fixARR(1, fixIDX) = retArray(fixIDX)
            Next fixIDX
            retArray = fixARR
        End If
        Set tmpRng = Nothing
    End With
    ClearTempRangeUtil
    
    If ArrDimensions(retArray) = 1 Then
        retArray = ConvertArrToRCArr(retArray)
    Else
        If ArrDimensions(retArray) = 0 Then
            Dim tmpValue As Variant
            tmpValue = retArray
            'a single value was returned, convert to RC1 array
            ReDim retArray(1 To 1, 1 To 1)
            retArray(1, 1) = tmpValue
        End If
    End If

    UniqueRC1Arr = retArray
    
    If ArrDimensions(fixARR) > 0 Then Erase fixARR
    If ArrDimensions(retArray) > 0 Then Erase retArray
    
    

End Function

Public Function DEVARR(pArr As Variant) As Variant

    DEVARR = ConvertArrToRCArr(pArr)

End Function

Private Function BuildRC1(rng As Range) As Variant
On Error GoTo e:
    Dim failed As Boolean
    Dim retArray As Variant
    Dim rgInfo As RngInfo
    
    rgInfo = RangeInfo(rng)
    
    If rgInfo.Areas = 1 Then
        If rgInfo.Columns = 1 And rgInfo.rows = 1 Then
            ReDim retArray(1 To 1, 1 To 1)
            retArray(1, 1) = rng.value
            GoTo Finalize:
        Else
            retArray = rng.value
            If ArrDimensions(retArray) = 1 Then
                retArray = ConvertArrToRCArr(retArray)
            End If
            GoTo Finalize:
        End If
    End If
    
    'Areas > 1
    If rgInfo.AreasSameRows = False And rgInfo.AreasSameColumns = False Then
        RaiseError ERR_INVALID_RANGE_SIZE, "All areas in Range must have matching RowCount or ColumnCount (ftRangeArray.BuildRC1)"
    End If
    
    ReDim retArray(1 To rgInfo.rows, 1 To rgInfo.Columns)
    
    ' ***** ***** ***** ***** ***** ***** ***** ***** *****
    Dim areaInfo As AreaStruct
    Dim idxAREA As Long, rngArea As Range, idxAreaRow As Long, idxAreaCol As Long
    Dim idxArrayRow As Long, idxArrayCol As Long
    Dim arrayRowOffset As Long, arrayColOffset As Long
    ' ***** ***** ***** ***** ***** ***** ***** ***** *****
    
    arrayRowOffset = 0
    arrayColOffset = 0
    
    If rgInfo.AreasSameRows Then
        ' *** *** *** *** *** ***
        ' *** SAME ROWS *** *
        ' *** *** *** *** *** ***
        For idxAREA = 1 To rgInfo.Areas
            areaInfo = RangeArea(rng.Areas(idxAREA))
            For idxAreaRow = 1 To areaInfo.rowCount
                For idxAreaCol = 1 To areaInfo.columnCount
                    retArray(idxAreaRow, idxAreaCol + arrayColOffset) = rng.Areas(idxAREA)(idxAreaRow, idxAreaCol)
                Next idxAreaCol
            Next idxAreaRow
            arrayColOffset = arrayColOffset + areaInfo.columnCount
        Next idxAREA
    
    Else
        ' *** *** *** *** *** ***
        ' *** SAME COLS *** *
        ' *** *** *** *** *** ***
        For idxAREA = 1 To rgInfo.Areas
            areaInfo = RangeArea(rng.Areas(idxAREA))
            For idxAreaRow = 1 To areaInfo.rowCount
                For idxAreaCol = 1 To areaInfo.columnCount
                    retArray(idxAreaRow + arrayRowOffset, idxAreaCol) = rng.Areas(idxAREA)(idxAreaRow, idxAreaCol)
                Next idxAreaCol
            Next idxAreaRow
            arrayRowOffset = arrayRowOffset + areaInfo.rowCount
        Next idxAREA
    End If

Finalize:
    On Error Resume Next
    
    If Not failed Then
        BuildRC1 = retArray
    End If
    
    If ArrDimensions(retArray) > 0 Then
        Erase retArray
    End If
    
    Set rngArea = Nothing
    If Err.number <> 0 Then Err.Clear
    Exit Function
e:
    failed = True
    ErrorCheck
    Resume Finalize:
    
End Function

Private Function ConvertArrToRCArr(ByVal arr As Variant, Optional zeroBasedAsColumns As Boolean = False) As Variant
    Dim retV() As Variant, rwCount As Long, isBase0 As Boolean, arrIdx As Long, colCount As Long
    If IsArrInit(arr) = False Then
        ReDim retV(1 To 1, 1 To 1)
        retV(1, 1) = arr
        ConvertArrToRCArr = retV
        Exit Function
    End If
    
    If ArrDimensions(arr) = 1 Then
        If zeroBasedAsColumns = False Then
            isBase0 = LBound(arr) = 0
            rwCount = UBound(arr) - LBound(arr) + 1
            If isBase0 Then
                ReDim retV(1 To UBound(arr) + 1, 1 To 1)
            Else
                ReDim retV(1 To UBound(arr), 1 To 1)
            End If
            For arrIdx = LBound(arr) To UBound(arr)
                If isBase0 Then
                    If IsObject(arr(arrIdx)) Then
                        Set retV(arrIdx + 1, 1) = arr(arrIdx)
                    Else
                        retV(arrIdx + 1, 1) = arr(arrIdx)
                    End If
                Else
                    If IsObject(arr(arrIdx)) Then
                        Set retV(arrIdx, 1) = arr(arrIdx)
                    Else
                        retV(arrIdx, 1) = arr(arrIdx)
                    End If
                End If
            Next arrIdx
            ConvertArrToRCArr = retV
        Else
            isBase0 = LBound(arr) = 0
            colCount = UBound(arr) - LBound(arr) + 1
            If isBase0 Then
                ReDim retV(1 To 1, 1 To UBound(arr) + 1)
            Else
                ReDim retV(1 To 1, 1 To UBound(arr))
            End If
            For arrIdx = LBound(arr) To UBound(arr)
                If isBase0 Then
                    If IsObject(arr(arrIdx)) Then
                        Set retV(1, arrIdx + 1) = arr(arrIdx)
                    Else
                        retV(1, arrIdx + 1) = arr(arrIdx)
                    End If
                Else
                    If IsObject(arr(arrIdx)) Then
                        Set retV(1, arrIdx) = arr(arrIdx)
                    Else
                        retV(1, arrIdx) = arr(arrIdx)
                    End If
                End If
            Next arrIdx
            ConvertArrToRCArr = retV
        End If
    Else
        ConvertArrToRCArr = arr
    End If
End Function



Private Function ClearTempRangeUtil()
    
    
    
    With TmpUtilSheet
        .Cells.EntireColumn.ColumnWidth = .StandardWidth
        .Cells.EntireRow.RowHeight = .StandardHeight
        .Cells.Clear
    End With
    
    
    
End Function
