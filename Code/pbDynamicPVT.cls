VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbDynamicPVT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' pbDynamicPvt v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' Pivot Utilities
'
' @module pbDynamicPvt
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit
Option Base 1
Option Compare Text

Private tblName As String
Private lPvtTable As PivotTable
Private lSrcTbl As ListObject
Private lDataFieldConsolFunctions As Dictionary

Private Function AddConsolidationFunction(fldName As Variant, consolName As Variant, Optional resetList As Boolean = False)
    If resetList Then
        Set lDataFieldConsolFunctions = Nothing
    End If
    If lDataFieldConsolFunctions Is Nothing Then Set lDataFieldConsolFunctions = New Dictionary
    lDataFieldConsolFunctions.Add fldName, consolName
End Function

Public Property Get DynamicPvtTable() As PivotTable
    If lPvtTable Is Nothing Then
        If wsDynamicPivot.PivotTables.Count = 1 Then
            Set lPvtTable = wsDynamicPivot.PivotTables(1)
        End If
    End If
    Set DynamicPvtTable = lPvtTable
End Property

Public Function Reset()
    If Not DynamicPvtTable Is Nothing Then
        DynamicPvtTable.ClearTable
    End If
End Function

Public Property Get sourceTable() As ListObject
    Set sourceTable = lSrcTbl
End Property

Public Function SetSourceTable(lstObj As ListObject)
    Set lSrcTbl = lstObj
End Function

Public Function FieldNames() As Dictionary

    Dim d As Dictionary
    Set d = New Dictionary
    Dim tmpARR As Variant
    tmpARR = WorksheetFunction.Sort(sourceTable.HeaderRowRange.Value)
    
    Dim idx As Long
    For idx = 1 To sourceTable.ListColumns.Count
        d.Add sourceTable.ListColumns(idx).Name, sourceTable.ListColumns(idx).Name
    Next idx
    Set FieldNames = d

End Function

Public Function AddPageFields(Optional flds As Variant)

    If IsMissing(flds) Then
        If AskYesNo("Would you like to keep the existing page fields?", "Add Pivot Page/Filter Fields", True) = vbNo Then
            ClearPivotFields XlPivotFieldOrientation.xlPageField
        End If
            
        Dim pkg As Package
        Set pkg = BuildPackage(pbPivotTable.DynamicPivot, "AddPageFields", VbMethod)
        Wom.picklist.Initialize "Choose Page Field(s)", "Select 1 or more page fields (used for filtering only on Pivot Table)", FieldNames, plMultiple_MinimumOne, pkg
        Exit Function
    End If
    
    If IsArray(flds) = False Then
        pbPivotUtil.UpdatePivotField DynamicPvtTable, flds, xlPageField
    Else
        Dim fidx As Long
        For fidx = LBound(flds) To UBound(flds)
            If ArrDimensions(flds) = 1 Then
                pbPivotUtil.UpdatePivotField DynamicPvtTable, flds(fidx), xlPageField
            Else
                pbPivotUtil.UpdatePivotField DynamicPvtTable, flds(fidx, 1), xlPageField
            End If
        Next fidx
    End If

End Function


Public Function AddRowFields(Optional flds As Variant)
    If IsMissing(flds) Then
        If AskYesNo("Would you like to keep the existing Row fields?", "Add Pivot Row Fields", True) = vbNo Then
            ClearPivotFields XlPivotFieldOrientation.xlRowField
        End If
            
        Dim pkg As Package
        Set pkg = BuildPackage(pbPivotTable.DynamicPivot, "AddRowFields", VbMethod)
        Wom.picklist.Initialize "Choose Row Field(s)", "Select 1 or more Row fields", FieldNames, plMultiple_MinimumOne, pkg
        Exit Function
    End If
    
    If IsArray(flds) = False Then
        pbPivotUtil.UpdatePivotField DynamicPvtTable, flds, xlRowField, , , True
    Else
        Dim fidx As Long
        For fidx = LBound(flds) To UBound(flds)
            If ArrDimensions(flds) = 1 Then
                pbPivotUtil.UpdatePivotField DynamicPvtTable, flds(fidx), xlRowField, , , True
            Else
                pbPivotUtil.UpdatePivotField DynamicPvtTable, flds(fidx, 1), xlRowField, , , True
            End If
        Next fidx
    End If
End Function

Public Function AddDataFields(Optional flds As Variant)
    If IsMissing(flds) Then
        If AskYesNo("Would you like to keep the existing Data (summarized) fields?", "Add Pivot Data Fields", True) = vbNo Then
            ClearPivotFields XlPivotFieldOrientation.xlDataField
        End If
            
        Dim pkg As Package
        Set pkg = BuildPackage(pbPivotTable.DynamicPivot, "AddDataFields", VbMethod)
        Wom.picklist.Initialize "Choose Data Field(s)", "Select 1 or more Data (Summary) fields", FieldNames, plMultiple_MinimumOne, pkg
        Exit Function
    End If
    
    
    ConfigureDataFields flds
    Exit Function
End Function

Public Function ConfigureDataFields(flds As Variant, Optional summaryType As Variant)

    Dim pkg As Package
    Dim curFld As Variant
    
    Dim fldCount As Long
    fldCount = UBound(flds) - LBound(flds) + 1
    
    If IsMissing(summaryType) Then
        Set lDataFieldConsolFunctions = Nothing
        Set lDataFieldConsolFunctions = New Dictionary
        curFld = flds(LBound(flds))
        Set pkg = BuildPackage(pbPivotTable.DynamicPivot, "ConfigureDataFields", VbMethod)
        pkg.UpdateArgs flds
        Wom.picklist.Initialize "Choose Summary Type for: " & curFld, "Choose Function to Summarize Field: " & curFld, GetConsolidationNames, plSingle, pkg
        Exit Function
    End If
    
    If IsArray(flds) Then
        curFld = flds(LBound(flds) + lDataFieldConsolFunctions.Count)
    Else
        curFld = flds
    End If
    
    lDataFieldConsolFunctions.Add curFld, summaryType
    
    If lDataFieldConsolFunctions.Count < fldCount Then
        curFld = flds(LBound(flds) + lDataFieldConsolFunctions.Count)
        Set pkg = BuildPackage(pbPivotTable.DynamicPivot, "ConfigureDataFields", VbMethod)
        pkg.UpdateArgs flds
        Wom.picklist.Initialize "Choose Summary Type for: " & curFld, "Choose Function to Summarize Field: " & curFld, GetConsolidationNames, plSingle, pkg
        Exit Function
    End If
    
    Dim dKey As Variant
    For Each dKey In lDataFieldConsolFunctions.Keys
        pbPivotUtil.UpdatePivotField DynamicPvtTable, dKey, xlDataField, , , True, GetConsolidationFunction(lDataFieldConsolFunctions(dKey))
    Next dKey
    
    
End Function


Public Function GetConsolidationFunction(consolidationName As String) As XlConsolidationFunction

    Dim retV As XlConsolidationFunction

    Select Case consolidationName
        Case "Average"
            retV = xlAverage
        Case "Count"
            retV = xlCount
        Case "Count Numeric"
            retV = xlCountNums
        Case "Count Distinct"
            retV = xlDistinctCount
        Case "Max"
            retV = xlMax
        Case "Min"
            retV = xlMin
        Case "Product"
            retV = xlProduct
        Case "StdDev"
            retV = xlStDev
        Case "StdDevP"
            retV = xlStDevP
        Case "Sum"
            retV = xlSum
        Case "Var"
            retV = xlVar
        Case "VarP"
            retV = xlVarP
        Case Else
            retV = xlSum
    End Select

    GetConsolidationFunction = retV

End Function

Public Function GetConsolidationNames() As Dictionary


    Dim d As Dictionary
    Set d = New Dictionary
    d.Add "Average", "Average"
    d.Add "Count", "Count"
    d.Add "Count Numeric", "Count Numeric"
    d.Add "Count Distinct", "Count Distinct"
    d.Add "Max", "Max"
    d.Add "Min", "Min"
    d.Add "Product", "Product"
    d.Add "StdDev", "StdDev"
    d.Add "StdDevP", "StdDevP"
    d.Add "Sum", "Sum"
    d.Add "Var", "Var"
    d.Add "VarP", "VarP"
    
    Set GetConsolidationNames = d

End Function

Public Function Subtotals(Optional flds As Variant)


    Dim pf As PivotField
    If IsMissing(flds) Then
        If AskYesNo("Clear All Subtotals? (Answer NO to choose which fields to show subtotals)", "Subtotals", False) = vbYes Then
            ClearSubtotals
            Exit Function
        End If
        Dim d As New Dictionary
        For Each pf In DynamicPvtTable.rowFields
            d(pf.SourceName) = pf.SourceName
        Next pf
        If d.Count > 0 Then
            Dim pkg As Package
            Set pkg = BuildPackage(pbPivotTable.DynamicPivot, "Subtotals", VbMethod)
            Wom.picklist.Initialize "Add Subtotals", "Choose fields to display subtotals", d, plMultiple_MinimumNone, pkg
        End If
        Exit Function
    
    End If
    
    Dim idx As Long
    For Each pf In DynamicPvtTable.rowFields
        pf.Subtotals(1) = False
        For idx = LBound(flds) To UBound(flds)
            If StrComp(pf.SourceName, flds(idx), vbTextCompare) = 0 Then
                pf.Subtotals(1) = True
            End If
        Next idx
    Next pf
    

End Function

Public Function ClearSubtotals()
    Dim pf As PivotField

    For Each pf In DynamicPvtTable.rowFields
        pf.Subtotals(1) = False
    Next pf

End Function

Public Function AddColumnFields(Optional flds As Variant)
    If IsMissing(flds) Then
        If AskYesNo("Would you like to keep the existing Column fields?", "Add Pivot Column Fields", True) = vbNo Then
            ClearPivotFields XlPivotFieldOrientation.xlColumnField
        End If
            
        Dim pkg As Package
        Set pkg = BuildPackage(pbPivotTable.DynamicPivot, "AddColumnFields", VbMethod)
        Wom.picklist.Initialize "Choose Column Field(s)", "Select 1 or more Column fields", FieldNames, plMultiple_MinimumOne, pkg
        Exit Function
    End If
    
    If IsArray(flds) = False Then
        pbPivotUtil.UpdatePivotField DynamicPvtTable, flds, xlColumnField
    Else
        Dim fidx As Long
        For fidx = LBound(flds) To UBound(flds)
            If ArrDimensions(flds) = 1 Then
                pbPivotUtil.UpdatePivotField DynamicPvtTable, flds(fidx), xlColumnField
            Else
                pbPivotUtil.UpdatePivotField DynamicPvtTable, flds(fidx, 1), xlColumnField
            End If
        Next fidx
    End If
End Function

Public Function ClearPivotFields(fldType As XlPivotFieldOrientation)

    Select Case fldType
        Case xlPageField
            ClearPageFields
        Case xlDataField
            ClearDataFields
        Case xlRowField
            ClearRowFields
        Case xlColumnField
            ClearColumnFields
    End Select

End Function

Private Function ClearColumnFields()

    Dim pf As PivotField
    For Each pf In DynamicPvtTable.ColumnFields
        pf.orientation = xlHidden
    Next pf

End Function

Private Function ClearRowFields()

    Dim pf As PivotField
    For Each pf In DynamicPvtTable.rowFields
        pf.orientation = xlHidden
    Next pf

End Function

Private Function ClearDataFields()

    Dim pf As PivotField
    Dim df As PivotField
    For Each pf In DynamicPvtTable.CalculatedFields
        For Each df In DynamicPvtTable.dataFields
          If df.SourceName = pf.Name Then
            With df
              .Parent.PivotItems(.Name).visible = False
            End With
            Exit For
          End If
        Next df
    Next pf
    
    For Each pf In DynamicPvtTable.dataFields
      pf.orientation = xlHidden
    Next pf

End Function

Private Function ClearPageFields()

    Dim pf As PivotField
    For Each pf In DynamicPvtTable.pageFields
        pf.orientation = xlHidden
    Next pf

End Function

Public Function Refresh()
    If Not DynamicPvtTable Is Nothing Then
        DynamicPvtTable.PivotCache.Refresh
    End If
End Function



Public Function ClearTable()
    If Not DynamicPvtTable Is Nothing Then
        DynamicPvtTable.ClearTable
    End If
End Function

