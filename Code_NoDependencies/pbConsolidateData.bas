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
Public Enum StaticMapTypeEnum
    [_smtError] = 0
    smtWorkookName = 2 ^ 0
    smtWorksheetName = 2 ^ 1
    smtManualValuePrefix = 2 ^ 2
    smtManualValueSuffix = 2 ^ 3
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
'   Enums need for private implementation of methods
'   from pbCommon
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Enum strMatchEnum
    smEqual = 0
    smNotEqualTo = 1
    smContains = 2
    smStartsWithStr = 3
    smEndWithStr = 4
End Enum
Private Enum ecComparisonType
    ecOR = 0 'default
    ecAnd
End Enum


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
            srcIndex = lSourceLO.ListColumns(srcColumn).index
        ElseIf IsNumeric(srcColumn) Then
            srcIndex = srcColumn
        End If
    ElseIf srcType = mtWorksheet Then
        If lSourceDataType = dfListObject Then
            srcIndex = srcColumn - lSourceLO.Range.column + 1
            If srcIndex < 1 Or srcIndex > lSourceLO.ListColumns.Count Then
                Err.Raise 1004, source:="pbConsolidateData.AddDataMap", Description:="'Worksheet Column Index (" & srcColumn & ") is not valid for " & lSourceLO.Name
            End If
        ElseIf lSourceDataType = dfRange Then
            srcIndex = srcColumn - lSourceRNG.column + 1
            If srcIndex < 1 Or srcIndex > lSourceRNG.Columns.Count Then
                Err.Raise 1004, source:="pbConsolidateData.AddDataMap", Description:="'Worksheet Column Index (" & srcColumn & ") is not valid for " & lSourceSheet.Name & "!" & lSourceRNG.Address
            End If
        End If
    End If
    If destType = mtRangeOrListObject Then
        If StringsMatch(TypeName(destColumn), "String") And lTargetDataType = dfListObject Then
            destIndex = lTargetLO.ListColumns(destColumn).index
        ElseIf IsNumeric(destColumn) Then
            destIndex = destColumn
        End If
    ElseIf destType = mtWorksheet Then
        If lTargetDataType = dfListObject Then
            destIndex = destColumn - lTargetLO.Range.column + 1
            If destIndex < 1 Or destIndex > lTargetLO.ListColumns.Count Then
                Err.Raise 1004, source:="pbConsolidateData.AddDataMap", Description:="'Worksheet Column Index (" & destColumn & ") is not valid for " & lTargetLO.Name
            End If
        ElseIf lTargetDataType = dfRange Then
            destIndex = destColumn - lTargetRNG.column + 1
            If destIndex < 1 Or destIndex > lTargetRNG.Columns.Count Then
                Err.Raise 1004, source:="pbConsolidateData.AddDataMap", Description:="'Worksheet Column Index (" & destColumn & ") is not valid for " & lTargetSheet.Name & "!" & lTargetRNG.Address
            End If
        End If
    End If
    lMap.Add Array(srcIndex, destIndex)

End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Add Source Workbook Name and/or Source Worksheet Name,
'       and/or Manual Value (First, or Last) as a data map, instead
'       of mapping data from Source ListObject or Range
'   [staticDataType] can be 1 or more enum arguments
'   e.g
'       AddStaticMap smtManualValuePrefix + smtWorksheetName, _
'            manualValue:=Format(Now, "yyyymmddhhnnss")
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function AddStaticMap(destColumn, destType As MapTypeEnum, staticDataType As StaticMapTypeEnum, Optional manualValue, Optional delimiter As String = "_")
    If lTargetLO Is Nothing And lTargetRNG Is Nothing Then
        Err.Raise 1004, source:="pbConsolidateData.AddStaticMap", Description:="'ConfigureMaster' and 'ConfigureSource' not valid"
    End If
    If lSourceLO Is Nothing And lSourceRNG Is Nothing Then
        Err.Raise 1004, source:="pbConsolidateData.AddStaticMap", Description:="'ConfigureMaster' and 'ConfigureSource' not valid"
    End If
    
    If lMap Is Nothing Then
        Set lMap = New Collection
    End If
    
    If staticDataType > 0 Then
        Dim tmpVal
        If EnumCompare(staticDataType, StaticMapTypeEnum.smtManualValuePrefix) And Not IsMissing(manualValue) Then
            tmpVal = manualValue
        End If
        If EnumCompare(staticDataType, StaticMapTypeEnum.smtWorkookName) Then
            tmpVal = tmpVal & IIf(Len(tmpVal) > 0, delimiter & lSourceSheet.Parent.Name, lSourceSheet.Parent.Name)
        End If
        If EnumCompare(staticDataType, StaticMapTypeEnum.smtWorksheetName) Then
            tmpVal = tmpVal & IIf(Len(tmpVal) > 0, delimiter & lSourceSheet.Name, lSourceSheet.Name)
        End If
        If EnumCompare(staticDataType, StaticMapTypeEnum.smtManualValueSuffix) And Not IsMissing(manualValue) Then
            tmpVal = tmpVal & IIf(Len(tmpVal) > 0, delimiter & manualValue, manualValue)
        End If
        
        Dim destIndex As Long
        If destType = mtRangeOrListObject Then
            If StringsMatch(TypeName(destColumn), "String") And lTargetDataType = dfListObject Then
                destIndex = lTargetLO.ListColumns(destColumn).index
            ElseIf IsNumeric(destColumn) Then
                destIndex = destColumn
            End If
        ElseIf destType = mtWorksheet Then
            If lTargetDataType = dfListObject Then
                destIndex = destColumn - lTargetLO.Range.column + 1
                If destIndex < 1 Or destIndex > lTargetLO.ListColumns.Count Then
                    Err.Raise 1004, source:="pbConsolidateData.AddDataMap", Description:="'Worksheet Column Index (" & destColumn & ") is not valid for " & lTargetLO.Name
                End If
            ElseIf lTargetDataType = dfRange Then
                destIndex = destColumn - lTargetRNG.column + 1
                If destIndex < 1 Or destIndex > lTargetRNG.Columns.Count Then
                    Err.Raise 1004, source:="pbConsolidateData.AddDataMap", Description:="'Worksheet Column Index (" & destColumn & ") is not valid for " & lTargetSheet.Name & "!" & lTargetRNG.Address
                End If
            End If
        End If
        lMap.Add Array("STATICMAP", tmpVal, destIndex)
    End If
End Function

Public Function Execute(Optional perfMode As Boolean = True)
    On Error GoTo E:
    
    If lTargetLO Is Nothing And lTargetRNG Is Nothing Then
        Err.Raise 1004, source:="pbConsolidateData.Execute", Description:="'ConfigureMaster' and 'ConfigureSource' not valid"
    End If
    If lSourceLO Is Nothing And lSourceRNG Is Nothing Then
        Err.Raise 1004, source:="pbConsolidateData.Execute", Description:="'ConfigureMaster' and 'ConfigureSource' not valid"
    End If
    If lMap Is Nothing Then
        Err.Raise 1004, source:="pbConsolidateData.Execute", Description:="Column Mapping Not Configured ('AddDataMap')"
    End If
    
    Dim evts As Boolean, scrn As Boolean
    
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
            If StringsMatch(mapItem(LBound(mapItem)), "STATICMAP") Then
                tmpDestRow(1, mapItem(UBound(mapItem))) = mapItem(LBound(mapItem) + 1)
            Else
                tmpDestRow(1, mapItem(UBound(mapItem))) = srcArray(srcRowIDX, mapItem(LBound(mapItem)))
            End If
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
    
    If perfMode Then
        evts = Application.EnableEvents
        scrn = Application.ScreenUpdating
        Application.EnableEvents = False
        Application.ScreenUpdating = False
    End If
    destRng.Value = destArray
Finalize:
    On Error Resume Next
    Reset
    If perfMode Then
        Application.EnableEvents = evts
        Application.ScreenUpdating = scrn
    End If
    Exit Function
E:
    'Implement Your Error Handling Code
    Beep
    MsgBox "Error Occured in pbConsolidateData.Execute (" & Err.number & " - " & Err.Description & ")"
    Err.Clear
    Resume Finalize:

End Function

Private Function GetNewTargetRange(newRowCount As Long) As Range
    Dim rngStart As Range
    If lTargetDataType = dfListObject Then
        If lTargetLO.listRows.Count = 0 Then
            Set rngStart = lTargetLO.HeaderRowRange.offSet(rowOffset:=1)
            lTargetLO.Resize lTargetLO.Range.Resize(rowSize:=lTargetLO.Range.Rows.Count + newRowCount - 1)

        Else
            Set rngStart = lTargetLO.listRows(lTargetLO.listRows.Count).Range.offSet(rowOffset:=1)
            lTargetLO.Resize lTargetLO.Range.Resize(rowSize:=lTargetLO.Range.Rows.Count + newRowCount)
        End If
        
        Set GetNewTargetRange = rngStart.Resize(rowSize:=newRowCount)
    Else
        Set rngStart = lTargetRNG.Rows(RowIndex:=lTargetRNG.Rows.Count).offSet(rowOffset:=1)
        Set GetNewTargetRange = rngStart.Resize(rowSize:=newRowCount)
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
'   Private Version of pbCommon.EnumCompare
'   Included in some 'pb[XXXXX]' modules in order to remove
'   dependency on pbCommon
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function EnumCompare(theEnum As Variant, enumMember As Variant, Optional ByVal iType As ecComparisonType = ecComparisonType.ecOR) As Boolean
    Dim c As Long
    c = theEnum And enumMember
    EnumCompare = IIf(iType = ecOR, c <> 0, c = enumMember)
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Private Version of pbCommon.Concat
'   Included in some 'pb[XXXXX]' modules in order to remove
'   dependency on pbCommon
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function Concat(ParamArray Items() As Variant) As String
    Concat = Join(Items, "")
End Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Private Version of pbCommon.ConcatWithDelim
'   Included in some 'pb[XXXXX]' modules in order to remove
'   dependency on pbCommon
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function ConcatWithDelim(ByVal delimeter As String, ParamArray Items() As Variant) As String
    ConcatWithDelim = Join(Items, delimeter)
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Private Version of pbCommon.StringsMatch
'   Included in some 'pb[XXXXX]' modules in order to remove
'   dependency on pbCommon
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Function StringsMatch( _
    ByVal checkString As Variant, ByVal _
    validString As Variant, _
    Optional smEnum As strMatchEnum = strMatchEnum.smEqual, _
    Optional compMethod As VbCompareMethod = vbTextCompare) As Boolean
'    Private Enum strMatchEnum
'        smEqual = 0
'        smNotEqualTo = 1
'        smContains = 2
'        smStartsWithStr = 3
'        smEndWithStr = 4
'    End Enum
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



