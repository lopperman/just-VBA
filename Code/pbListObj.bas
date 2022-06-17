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

Public Function PopulateStaticFromFormula(lstObj As ListObject, lstColName As String, r1c1Formula As String, Optional createIfMissing As Boolean = True, Optional numberFormat As String = vbNullString) As Boolean
'   Create a formula in ListColumn, and then replace the ListColumn contents with the values from the Formula
'   Note: Sets the 'Formula2R1C1' Formula Property.
On Error GoTo E:
    Dim failed As Boolean
    If lstObj.listRows.Count = 0 Then
        Exit Function
    End If
    
    
    Dim tmpColArr As Variant
    If ListColumnExists(lstObj, lstColName) = False And createIfMissing = True Then
        AddColumnIfMissing lstObj, lstColName
    End If

    If Not ListColumnExists(lstObj, lstColName) Then
        failed = True
        GoTo Finalize:
    End If
    
    lstObj.ListColumns(lstColName).DataBodyRange.ClearContents
    lstObj.ListColumns(lstColName).DataBodyRange.numberFormat = "General"
    lstObj.ListColumns(lstColName).DataBodyRange.Formula2R1C1 = r1c1Formula
     
    ReplaceFormulasWithStatic lstObj
     
    tmpColArr = ListColumnAsArray(lstObj, lstColName)
    lstObj.ListColumns(lstColName).DataBodyRange.ClearContents
    If Len(numberFormat) > 0 Then
        lstObj.ListColumns(lstColName).DataBodyRange.numberFormat = numberFormat
    End If
    If ArrDimensions(tmpColArr) = 2 Then
        lstObj.ListColumns(lstColName).DataBodyRange.value = tmpColArr
    Else
        failed = True
    End If
    
Finalize:
    On Error Resume Next
    
    If ArrDimensions(tmpColArr) > 0 Then Erase tmpColArr
    PopulateStaticFromFormula = Not failed
    
    
    If Err.Number <> 0 Then Err.Clear
    Exit Function
E:
    failed = True
    ErrorCheck
    Resume Finalize:
End Function
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

Public Function ListColumnExists(ByRef lstObj As ListObject, lstColName As String) As Boolean
'   Return true If 'lstColName' is a valid ListColumn in 'lstObj'
    ListColumnExists = ListColumnIndex(lstObj, lstColName) > 0
End Function

Public Function ListColumnIndex(ByRef lstObj As ListObject, lstColName As String) As Long
    Dim retV As Long
    Dim colIDX As Long
    For colIDX = 1 To lstObj.ListColumns.Count
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
            Set nc = lstObj.ListColumns.add(position:=position)
        Else
            Set nc = lstObj.ListColumns.add
        End If
        nc.Name = colName
        If lstObj.listRows.Count > 0 And numberFormat <> vbNullString Then
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
    If HasData(lstObj) = False Then
        CountBlanks = 0
        Exit Function
    End If
    Set rng = lstObj.ListColumns(listColumn).DataBodyRange.SpecialCells(xlCellTypeBlanks)
    If Not rng Is Nothing Then
        CountBlanks = rng.Rows.Count
    End If
    Set rng = Nothing
    If Err.Number <> 0 Then Err.Clear

End Function

Public Function HasData(lstObj As Variant) As Boolean
On Error Resume Next
    If TypeName(lstObj) = "ListObject" Then
        HasData = lstObj.listRows.Count > 0
    ElseIf TypeName(lstObj) = "String" Then
        HasData = wt(CStr(lstObj)).listRows.Count > 0
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


