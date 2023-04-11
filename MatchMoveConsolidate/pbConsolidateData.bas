Attribute VB_Name = "pbConsolidateData"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  CONSOLIDATE DATA FROM LISTOBJECTS/TABLES OR
'  RANGES INTO A 'MASTER' LISTOBJECT/RANGE
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbConsolidateData.bas
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Option Explicit
Option Base 1
Option Compare Text

Public Enum DataFormatEnum
    [_dferror] = 0
    dfRange = 1
    dfListObject = 2
End Enum
Public Enum MapTypeEnum
    [_mterror] = 0
    mtWorksheet = 1
    mtRangeOrListObject = 2
End Enum

 Public Enum strMatchEnum
     smEqual = 0
     smNotEqualTo = 1
     smContains = 2
     smStartsWithStr = 3
     smEndWithStr = 4
 End Enum


Private lTargetDataType As DataFormatEnum
Private lTargetSheet As Worksheet
Private lTargetLO As ListObject
Private lTargetRNG As Range
Private lSourceDataType As DataFormatEnum
Private lSourceSheet As Worksheet
Private lSourceLO As ListObject
Private lSourceRNG As Range
Private lMap As Collection


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
' If Target is a ListObject, source can be ListObject name, or
'   any range that is part of the list object
' Calling this method RESETS everything
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function ConfigureTarget(dataType As DataFormatEnum, wksht As Worksheet, source)
    Reset
    lTargetDataType = dataType
    Set lTargetSheet = wksht
    If dataType = dfListObject Then
        If StringsMatch(TypeName(source), "String") Then
            If ValidListObject(wksht, source) Then
                Set lTargetLO = wksht.ListObjects(source)
            End If
        ElseIf StringsMatch(TypeName(source), "Range") Then
            If Not source.ListObject Is Nothing Then
                Set lTargetLO = source.ListObject
            End If
        End If
    ElseIf dataType = dfRange Then
        If StringsMatch(TypeName(source), "Range") Then
            Set lTargetRNG = source
        End If
    End If
    If lTargetLO Is Nothing And lTargetRNG Is Nothing Then
        Err.Raise 1004, source:="pbConsolidateData.ConfigureTarget", Description:="[source] is invalid"
    End If
End Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
' If Source is a ListObject, source can be ListObject name, or
'   any range that is part of the list object
' 'ConfigureTarget' must be run first
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function ConfigureSource(dataType As DataFormatEnum, wksht As Worksheet, source)
    If lTargetLO Is Nothing And lTargetRNG Is Nothing Then
        Err.Raise 1004, source:="pbConsolidateData.ConfigureSource", Description:="'ConfigureTarget' required before 'ConfigureSource'"
    End If
    
    lSourceDataType = dataType
    Set lSourceSheet = wksht
    If dataType = dfListObject Then
        If StringsMatch(TypeName(source), "String") Then
            If ValidListObject(wksht, source) Then
                Set lSourceLO = wksht.ListObjects(source)
            End If
        ElseIf StringsMatch(TypeName(source), "Range") Then
            If Not source.ListObject Is Nothing Then
                Set lSourceLO = source.ListObject
            End If
        End If
    ElseIf dataType = dfRange Then
        If StringsMatch(TypeName(source), "Range") Then
            Set lSourceRNG = source
        End If
    End If
    If lSourceLO Is Nothing And lSourceRNG Is Nothing Then
        Err.Raise 1004, source:="pbConsolidateData.ConfigureSource", Description:="[source] is invalid"
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
' If referring to ListObject, Provide ListObject Column Name OR
'   ListObject Column Index, OR ** Worksheet Column ** with
'   mapType = mtWorksheet, and the appropriate ListColumn
'   index will be calculated
' If referring to Range, provide the Range column index OR
'  ** Worksheet Column ** with mapType = mtWorksheet,
'  and the range column index will be calculated
'
' 'ConfigureMaster' and 'ConfigureSource' must be run first
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function AddDataMap(srcColumn, srcType As MapTypeEnum, destColumn, destType As MapTypeEnum)
    If lTargetLO Is Nothing And lTargetRNG Is Nothing Then
        Err.Raise 1004, source:="pbConsolidateData.AddDataMap", Description:="'ConfigureMaster' and 'ConfigureSource' not valid"
    End If
    If lSourceLO Is Nothing And lSourceRNG Is Nothing Then
        Err.Raise 1004, source:="pbConsolidateData.AddDataMap", Description:="'ConfigureMaster' and 'ConfigureSource' not valid"
    End If
    
    If lMap Is Nothing Then
        Set lMap = New Collection
    End If
    
    Dim srcIndex As Long, destIndex As Long
    If srcType = mtRangeOrListObject Then
        If StringsMatch(TypeName(srcColumn), "String") And lSourceDataType = dfListObject Then
            srcIndex = lSourceLO.ListColumns(srcColumn).Index
        ElseIf IsNumeric(srcColumn) Then
            srcIndex = srcColumn
        End If
    ElseIf srcType = mtWorksheet Then
        If lSourceDataType = dfListObject Then
            srcIndex = srcColumn - lSourceLO.Range.Column + 1
            If srcIndex < 1 Or srcIndex > lSourceLO.ListColumns.Count Then
                Err.Raise 1004, source:="pbConsolidateData.AddDataMap", Description:="'Worksheet Column Index (" & srcColumn & ") is not valid for " & lSourceLO.Name
            End If
        ElseIf lSourceDataType = dfRange Then
            srcIndex = srcColumn - lSourceRNG.Column + 1
            If srcIndex < 1 Or srcIndex > lSourceRNG.Columns.Count Then
                Err.Raise 1004, source:="pbConsolidateData.AddDataMap", Description:="'Worksheet Column Index (" & srcColumn & ") is not valid for " & lSourceSheet.Name & "!" & lSourceRNG.Address
            End If
        End If
    End If
    If destType = mtRangeOrListObject Then
        If StringsMatch(TypeName(destColumn), "String") And lTargetDataType = dfListObject Then
            destIndex = lTargetLO.ListColumns(destColumn).Index
        ElseIf IsNumeric(destColumn) Then
            destIndex = destColumn
        End If
    ElseIf destType = mtWorksheet Then
        If lTargetDataType = dfListObject Then
            destIndex = destColumn - lTargetLO.Range.Column + 1
            If destIndex < 1 Or destIndex > lTargetLO.ListColumns.Count Then
                Err.Raise 1004, source:="pbConsolidateData.AddDataMap", Description:="'Worksheet Column Index (" & destColumn & ") is not valid for " & lTargetLO.Name
            End If
        ElseIf lTargetDataType = dfRange Then
            destIndex = destColumn - lTargetRNG.Column + 1
            If destIndex < 1 Or destIndex > lTargetRNG.Columns.Count Then
                Err.Raise 1004, source:="pbConsolidateData.AddDataMap", Description:="'Worksheet Column Index (" & destColumn & ") is not valid for " & lTargetSheet.Name & "!" & lTargetRNG.Address
            End If
        End If
    End If
    lMap.Add Array(srcIndex, destIndex)

End Function

Public Function Execute()
    If lTargetLO Is Nothing And lTargetRNG Is Nothing Then
        Err.Raise 1004, source:="pbConsolidateData.Execute", Description:="'ConfigureMaster' and 'ConfigureSource' not valid"
    End If
    If lSourceLO Is Nothing And lSourceRNG Is Nothing Then
        Err.Raise 1004, source:="pbConsolidateData.Execute", Description:="'ConfigureMaster' and 'ConfigureSource' not valid"
    End If
    If lMap Is Nothing Then
        Err.Raise 1004, source:="pbConsolidateData.Execute", Description:="Column Mapping Not Configured ('AddDataMap')"
    End If
    
    Dim srcArray() As Variant
    Dim tmpCol As New Collection
    Dim tmpDestRow() As Variant
    Dim destArray() As Variant
    Dim srcRowIDX As Long
    Dim mapItem As Variant
    Dim destItem As Variant
    Dim newDestIdx As Long
    Dim destColIdx As Long
    Dim destRng As Range
    srcArray = GetSourceArray
    For srcRowIDX = LBound(srcArray, 1) To UBound(srcArray, 1)
        tmpDestRow = NewDestRow
        For Each mapItem In lMap
            tmpDestRow(1, mapItem(UBound(mapItem))) = srcArray(srcRowIDX, mapItem(LBound(mapItem)))
        Next mapItem
        tmpCol.Add tmpDestRow
    Next srcRowIDX
    ReDim destArray(1 To tmpCol.Count, 1 To UBound(NewDestRow, 2))
    For Each destItem In tmpCol
        newDestIdx = newDestIdx + 1
        For destColIdx = 1 To UBound(destArray, 2)
            destArray(newDestIdx, destColIdx) = destItem(1, destColIdx)
        Next destColIdx
    Next destItem
    Set destRng = GetNewTargetRange(tmpCol.Count)
    destRng.Value = destArray
    Reset
End Function

Private Function GetNewTargetRange(newRowCount As Long) As Range
    Dim rngStart As Range
    If lTargetDataType = dfListObject Then
        If lTargetLO.ListRows.Count = 0 Then
            Set rngStart = lTargetLO.HeaderRowRange.Offset(RowOffset:=1)
            lTargetLO.Resize lTargetLO.Range.Resize(RowSize:=lTargetLO.Range.Rows.Count + newRowCount - 1)

        Else
            Set rngStart = lTargetLO.ListRows(lTargetLO.ListRows.Count).Range.Offset(RowOffset:=1)
            lTargetLO.Resize lTargetLO.Range.Resize(RowSize:=lTargetLO.Range.Rows.Count + newRowCount)
        End If
        
        Set GetNewTargetRange = rngStart.Resize(RowSize:=newRowCount)
    Else
        Set rngStart = lTargetRNG.Rows(RowIndex:=lTargetRNG.Rows.Count).Offset(RowOffset:=1)
        Set GetNewTargetRange = rngStart.Resize(RowSize:=newRowCount)
    End If
End Function

Private Function NewDestRow() As Variant()
    Dim destColCount As Long
    Dim newRow() As Variant
    destColCount = TargetColsCount
    ReDim newRow(1 To 1, 1 To destColCount)
    NewDestRow = newRow
End Function

Private Function TargetColsCount() As Long
    If lTargetDataType = dfListObject Then
        TargetColsCount = lTargetLO.ListColumns.Count
    Else
        TargetColsCount = lTargetRNG.Columns.Count
    End If
End Function

Private Function GetSourceArray() As Variant()
    If lSourceDataType = dfListObject Then
        GetSourceArray = lSourceLO.DataBodyRange.Value
    Else
        GetSourceArray = lSourceRNG.Value
    End If
End Function

Private Function Reset()
    lTargetDataType = 0
    Set lTargetSheet = Nothing
    Set lTargetLO = Nothing
    Set lTargetRNG = Nothing
    lSourceDataType = 0
    Set lSourceSheet = Nothing
    Set lSourceLO = Nothing
    Set lSourceRNG = Nothing
    Set lMap = Nothing
End Function
Private Function ValidListObject(wksht As Worksheet, lstObjName) As Boolean
    Dim lo As ListObject
    For Each lo In wksht.ListObjects
        If StringsMatch(lo.Name, lstObjName) Then
            ValidListObject = True
            Exit For
        End If
    Next lo
End Function




' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
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


