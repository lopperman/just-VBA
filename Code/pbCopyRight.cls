VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbCopyRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  'COPY' DATA THE RIGHT WAY
'   (DOES NOT USE [RANGE].COPY/PASTE) -- EVER
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbCopyRight
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Option Explicit
Option Compare Text
Option Base 1

Private Enum ArrayOptionFlags
    aoNone = 0
    aoUnique = 2 ^ 0
    aoUniqueNoSort = 2 ^ 1
    aoAreaSizesMustMatch = 2 ^ 2
    aoVisibleRangeOnly = 2 ^ 3
    aoIncludeListObjHeaderRow = 2 ^ 4
End Enum
Private Type ArrInformation
    Rows As Long
    Columns As Long
    Dimensions As Long
    Ubound_first As Long
    LBound_first As Long
    UBound_second As Long
    LBound_second As Long
    IsArray As Boolean
End Type
Private Type AreaStruct
    RowStart As Long
    RowEnd As Long
    ColStart As Long
    ColEnd As Long
    rowCount As Long
    columnCount As Long
End Type
Private Type RngInfo
    Rows As Long
    Columns As Long
    AreasSameRows As Boolean
    AreasSameColumns As Boolean
    Areas As Long
End Type
Private Enum strMatchEnum
    smEqual = 0
    smNotEqualTo = 1
    smContains = 2
    smStartsWithStr = 3
    smEndWithStr = 4
End Enum


Private Const TEMP_UTIL_SHEET As String = "justVBATempUtil"
Private lTempSheet As Worksheet

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   GET HIDDEN TEMP WORKSHEET
'   USED TO SORT 2D ARRAYS
'   * CREATES NEW HIDDEN WORKSHEET IF NEEDED
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Property Get TempSheet() As Worksheet
    Dim shtExists As Boolean, evtsOn As Boolean
    If lTempSheet Is Nothing Then
        Dim tWS As Worksheet
        For Each tWS In ThisWorkbook.Worksheets
            If StringsMatch(tWS.Name, TEMP_UTIL_SHEET) Then
                Set lTempSheet = tWS
                shtExists = True
                Exit For
            End If
        Next tWS
        If Not shtExists Then
            'Add To End
            evtsOn = Application.EnableEvents
            Application.EnableEvents = False
            Set lTempSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            Application.EnableEvents = evtsOn
        End If
    End If
    If Not lTempSheet Is Nothing Then
        If Not lTempSheet.visible = xlSheetVeryHidden Then
            lTempSheet.visible = xlSheetVeryHidden
        End If
        Set TempSheet = lTempSheet
    End If
End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  UTILITIES (Renamed with prefix added to not conflict
'   with copies just-VBA common modules)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function ArrRange(rng As Range, flags As ArrayOptionFlags) As Variant
    Dim retArray As Variant
    Dim Unique As Boolean
    Dim aInfo As ArrInformation
    Dim tmpValue As Variant
    Dim RInfo As RngInfo
    RInfo = RangeInfo(rng)
    If EnumCompare(flags, aoAreaSizesMustMatch) Then
        If RInfo.AreasSameColumns = False And RInfo.AreasSameRows = False Then
            RaiseError ERR_INVALID_RANGE_SIZE, "Range Areas must all be the same size (ftRangeArray.ArrRange)"
        End If
    End If
    Unique = EnumCompare(flags, ArrayOptionFlags.aoUnique + ArrayOptionFlags.aoUniqueNoSort, ecOR)
   If EnumCompare(flags, ArrayOptionFlags.aoVisibleRangeOnly) Then
        retArray = BuildRC1(rng.SpecialCells(xlCellTypeVisible))
   Else
        retArray = BuildRC1(rng)
   End If
    If Unique Then
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
Private Function ArrListObj(lstObj As ListObject, flags As ArrayOptionFlags) As Variant
'   Returns 1-based, 2D array from entire ListObject Data Body Range, with options
' (don't need to deal with state here)
    ArrListObj = ArrRange(lstObj.DataBodyRange, flags)
End Function

Private Function RangeInfo(rg As Range) As RngInfo
    Dim retV As RngInfo
    If rg Is Nothing Then
        retV.Rows = 0
        retV.Columns = 0
        retV.AreasSameRows = False
        retV.AreasSameColumns = False
        retV.Areas = 0
    Else
        retV.Rows = RangeRowCount(rg)
        retV.Columns = RangeColCount(rg)
        retV.AreasSameRows = ContiguousRows(rg)
        retV.AreasSameColumns = ContiguousColumns(rg)
        retV.Areas = rg.Areas.Count
    End If
    RangeInfo = retV
End Function

Private Function RangeArea(rg As Range) As AreaStruct
'   Return info about Range area
'   Range with 1 Area allowed, otherwise error
    If rg.Areas.Count <> 1 Then
        RaiseError ERR_RANGE_AREA_COUNT, "Range Area Count <> 1 (ftRangeArray.RangeArea)"
    End If
    
    Dim retV As AreaStruct
    retV.RowStart = rg.Row
    retV.RowEnd = rg.Row + rg.Rows.Count - 1
    retV.ColStart = rg.column
    retV.ColEnd = rg.column + rg.Columns.Count - 1
    retV.rowCount = rg.Rows.Count
    retV.columnCount = rg.Columns.Count
    
    RangeArea = retV

End Function


Private Function RangeToUniqueArray(rng As Range, Optional sorted As Boolean = True, Optional excludeEmpty As Boolean = True) As Variant()
    Dim d As New Dictionary
    d.CompareMode = TextCompare
    Dim i As Long, rngArea, rngCell, tmpVal
    For Each rngArea In rng.Areas
        For Each rngCell In rngArea
            tmpVal = rngCell.Value
            If StringsMatch(TypeName(tmpVal), "String") Then tmpVal = Trim(tmpVal)
            If Not d.Exists(tmpVal) Then
                If excludeEmpty = True Then
                    If Not StringsMatch(TypeName(tmpVal), "Empty") Then
                        d.Add tmpVal, tmpVal
                    End If
                Else
                        d.Add tmpVal, tmpVal
                End If
            End If
        Next rngCell
    Next rngArea
    If d.Count > 0 Then
        If sorted And d.Count > 1 Then
            If tmpRangeUtil.EnableAutoFilter = False Then tmpRangeUtil.EnableAutoFilter = True
            Dim sRng As Range
            tmpRangeUtil.Cells.Clear
            DoEvents
            Set sRng = tmpRangeUtil.Range("A1")
            Set sRng = sRng.Resize(rowSize:=d.Count)
            sRng.Value = ArrArray(d.Keys, aoNone)
            
            'Range("A1:A4").Select
            tmpRangeUtil.Sort.SortFields.Clear
            tmpRangeUtil.Sort.SortFields.add2 key:=Range( _
                sRng.Address), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                xlSortNormal
            With tmpRangeUtil.Sort
                .SetRange Range(sRng.Address)
                .Header = xlNo
                .MatchCase = False
                .orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With

            RangeToUniqueArray = RangeToUniqueArray(sRng, sorted:=False)
            tmpRangeUtil.Cells.Clear
        Else
            RangeToUniqueArray = d.Keys
        End If
    End If
End Function

Private Function RangeTo1DArray(ByVal rng As Range) As Variant()
'TODO:  Optimizae to build 1D array from Arrays from each Area in rng
    

'   Return all cells in Range as 1D Array
    Dim retV() As Variant
    ''BASE 1
    ReDim retV(1 To rng.Count)
    Dim cl As Range, clIDX As Long
    clIDX = 1
    For Each cl In rng.Cells
        retV(clIDX) = cl.Value
        clIDX = clIDX + 1
    Next cl
    RangeTo1DArray = retV
    
    
End Function

Private Function GetUniqueSortedListCol(lstObj As ListObject, lstCol As Variant, Optional returnType As ListReturnType = ListReturnType.lrtArray) As Variant
'   Returns unique 1-based, 2D array from specific ListObject ListColumn
'   Return Type = Array (default), Dictionary, or Collection
    If lstObj.listRows.Count = 0 Then Exit Function
    
    
    Dim tDic As Dictionary
    Dim tCol As Collection
    
    Dim aIDX As Long, arr As Variant
    arr = ArrListCols(lstObj, aoUnique, lstCol)
    
    Select Case returnType
        Case ListReturnType.lrtArray
            GetUniqueSortedListCol = arr
        
        Case ListReturnType.lrtDictionary
            Set tDic = New Dictionary
            For aIDX = LBound(arr) To UBound(arr)
                tDic(arr(aIDX, 1)) = arr(aIDX, 1)
            Next aIDX
            Set GetUniqueSortedListCol = tDic
        
        Case ListReturnType.lrtCollection
            Set tDic = New Collection
            For aIDX = LBound(arr) To UBound(arr)
                tCol.Add arr(aIDX, 1)
            Next aIDX
            Set GetUniqueSortedListCol = tCol
    End Select
    
    Set tDic = Nothing
    Set tCol = Nothing
    
    If ArrDimensions(arr) > 0 Then
        Erase arr
    End If
    
    
    
End Function
Private Function ArrListCols(lstObj As ListObject, flags As ArrayOptionFlags, ParamArray listCols() As Variant) As Variant
'   Get Array from specific ListObject listColum(s)
    Dim idx As Long, rng As Range, inclHeader As Boolean
    inclHeader = EnumCompare(flags, aoIncludeListObjHeaderRow)
    If lstObj.listRows.Count > 0 Then
        For idx = LBound(listCols) To UBound(listCols)
            If rng Is Nothing Then
                If inclHeader Then
                    Set rng = lstObj.HeaderRowRange(1, lstObj.ListColumns(listCols(idx)).index)
                    Set rng = rng.Resize(rowSize:=lstObj.listRows.Count + lstObj.HeaderRowRange.Rows.Count)
                Else
                    Set rng = lstObj.ListColumns(listCols(idx)).DataBodyRange
                End If
            Else
                If inclHeader Then
                    Dim tRng As Range
                    Set tRng = lstObj.HeaderRowRange(1, lstObj.ListColumns(listCols(idx)).index)
                    Set tRng = tRng.Resize(rowSize:=lstObj.listRows.Count + lstObj.HeaderRowRange.Rows.Count)
                    Set rng = Union(rng, tRng)
                    Set tRng = Nothing
                Else
                    Set rng = Union(rng, lstObj.ListColumns(listCols(idx)).DataBodyRange)
                End If
            End If
        Next idx
        ArrListCols = ArrRange(rng, flags)
    End If
    Set rng = Nothing
    
    
End Function


Private Function ArrArray(ByVal arr As Variant, flags As ArrayOptionFlags, Optional zeroBasedAsColumns As Boolean = False) As Variant
'   By default, a zero-based array will become multiple rows.  Set 'zeroBasedAsColumns' to create 1 row with multiple columns
    Dim retArray As Variant
    Dim Unique As Boolean
    
    '   CHECK TO DETERMINE IF 'ARR' IS FROM A PARAMARRAY -- WHICH MEANS WE SHOULD TAKE 'ARR(0)' AS THE INPUT ARRAY
    If ValidArray(arr) And LBound(arr) = UBound(arr) Then
        If EnumCompare(VarType(arr(LBound(arr))), VbVarType.vbArray) Then
            If UBound(arr) >= LBound(arr) Then
                arr = arr(LBound(arr))
            End If
        End If
    End If
    
    Unique = EnumCompare(flags, ArrayOptionFlags.aoUnique)

    If ArrDimensions(arr) = 1 Then
        retArray = ConvertArrToRCArr(arr, zeroBasedAsColumns)
    Else
        retArray = arr
    End If
    Dim ai As ArrInformation
    ai = ArrayInfo(retArray)
    If Unique Then
        If ai.Dimensions = 0 Then
            Err.Raise 427, Description:="Array not initialized"
        End If
        If ai.Rows = 1 And ai.Columns = 1 Then
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
Private Function IsArrInit(inpt As Variant) As Boolean
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
    

Private Function ArrayInfo(arr As Variant) As ArrInformation
'   Returns Information about array dimensions
'   Note: Use Arr[X] Functions in pbRangeArray (e.g. 'ArrRange', 'ArrArray', 'ArrListObject') to ensure all arrays
'       are 1-based, 2-dimensional - required for populating worksheet ranges in a 'table style rows/columns' convention
On Error Resume Next
    Dim TMP As ArrInformation
    TMP.IsArray = ValidArray(arr)
    If TMP.IsArray = False Then GoTo Finalize:
        
    If UBound(arr) = -1 Or LBound(arr) > UBound(arr) Then
        TMP.Dimensions = 0
    Else
        TMP.Dimensions = ArrDimensions(arr)
        If TMP.Dimensions > 0 Then
            TMP.LBound_first = LBound(arr, 1)
            TMP.Ubound_first = UBound(arr, 1)
            TMP.Rows = (TMP.Ubound_first - TMP.LBound_first) + 1
        End If
        If TMP.Dimensions = 1 Then
            TMP.Columns = 1
        Else
            If TMP.Dimensions = 2 Then
                TMP.Columns = (UBound(arr, 2) - LBound(arr, 2)) + 1
            End If
        End If
        If TMP.Dimensions >= 2 Then
            TMP.LBound_second = LBound(arr, 2)
            TMP.UBound_second = UBound(arr, 2)
        End If
    End If
    
Finalize:
    ArrayInfo = TMP
    If Err.number <> 0 Then Err.Clear
End Function

Private Function ArrDimensions(ByRef checkArr As Variant) As Long
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
        Dim TMP As Variant
        TMP = UBound(checkArr, dimCount + 1)
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
    
    With TempSheet
        Set tmpRng = .Range("A1")
        Set tmpRng = tmpRng.Resize(rowSize:=aInfo.Rows, ColumnSize:=aInfo.Columns)
        tmpRng.Value = arr
        If Not EnumCompare(flags, ArrayOptionFlags.aoUniqueNoSort) Then
            Dim sidx As Long, sRng As Range
            .Sort.SortFields.Clear
            For sidx = 1 To tmpRng.Columns.Count
                Set sRng = tmpRng.Resize(ColumnSize:=1).offSet(ColumnOffset:=sidx - 1)
                .Sort.SortFields.add2 key:=.Range(sRng.Address), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
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
        retArray = WorksheetFunction.Unique(tmpRng, False, False)
        retAI = ArrayInfo(retArray)
        If retAI.Columns < aInfo.Columns And retAI.Dimensions = 1 Then
            'got down to one column with > 1 value, so now it needs to be flipped
            Dim fixIDX As Long
            ReDim fixARR(1 To 1, 1 To retAI.Rows)
            For fixIDX = 1 To retAI.Rows
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



Private Function BuildRC1(rng As Range) As Variant
On Error GoTo E:
    Dim failed As Boolean
    Dim retArray As Variant
    Dim rgInfo As RngInfo
    rgInfo = RangeInfo(rng)
    If rgInfo.Areas = 1 Then
        If rgInfo.Columns = 1 And rgInfo.Rows = 1 Then
            ReDim retArray(1 To 1, 1 To 1)
            retArray(1, 1) = rng.Value
            GoTo Finalize:
        Else
            retArray = rng.Value
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
    ReDim retArray(1 To rgInfo.Rows, 1 To rgInfo.Columns)
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
E:
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
    With TempSheet
        .Cells.EntireColumn.ColumnWidth = .StandardWidth
        .Cells.EntireRow.RowHeight = .StandardHeight
        .Cells.Clear
    End With
End Function

' ~~~ ~~ FLAG ENUM COMPARE ~~~ ~~~
Private Function EnumCompare(theEnum As Variant, enumMember As Variant, Optional ByVal iType As ecComparisonType = ecComparisonType.ecOR) As Boolean
    Dim c As Long
    c = theEnum And enumMember
    EnumCompare = IIf(iType = ecOR, c <> 0, c = enumMember)
End Function

Private Function StringsMatch( _
    ByVal checkString As Variant, ByVal _
    validString As Variant, _
    Optional smEnum As strMatchEnum = strMatchEnum.smEqual, _
    Optional compMethod As VbCompareMethod = vbTextCompare) As Boolean
    Dim str1, str2
    str1 = CStr(checkString)
    str2 = CStr(validString)
    Select Case smEnum
        Case strMatchEnum.smEqual
            StringsMatch = StrComp(str1, str2, compMethod) = 0
        Case strMatchEnum.smNotEqualTo
            StringsMatch = StrComp(str1, str2, compMethod) <> 0
        Case strMatchEnum.smContains
            StringsMatch = InStr(1, str1, str2, compMethod) > 0
        Case strMatchEnum.smStartsWithStr
            StringsMatch = InStr(1, str1, str2, compMethod) = 1
        Case strMatchEnum.smEndWithStr
            If Len(str2) > Len(str1) Then
                StringsMatch = False
            Else
                StringsMatch = InStr(Len(str1) - Len(str2) + 1, str1, str2, compMethod) = Len(str1) - Len(str2) + 1
            End If
    End Select
End Function

Private Function RangeRowCount(ByVal rng As Range) As Long
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
Private Function RangeColCount(ByVal rng As Range) As Long
    Dim tmpCount As Long
    Dim colDict As Dictionary
    Dim firstCol As Long, areaIDX As Long, colidx As Long
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
            For colidx = 1 To rng.Areas(areaIDX).Columns.Count
                If colidx > 1 Then colDict(firstCol + (colidx - 1)) = firstCol + (colidx - 1)
            Next colidx
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

Private Function ContiguousRows(rng As Range) As Boolean
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

Private Function ContiguousColumns(rng As Range) As Boolean
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






