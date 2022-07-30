Attribute VB_Name = "pbPivotUtil"
Option Explicit
Option Compare Text
Option Base 1

Public Function UpdateRowField(ByRef pvt As PivotTable, fldName As Variant, visible As Boolean, Optional position As Variant, Optional showSubTotal As Boolean = False)

    Dim pf As PivotField
    Set pf = pvt.PivotFields(fldName)
    If Not pf Is Nothing Then
        With pf
            If .orientation <> xlRowField Then
                position = pvt.rowFields.count + 1
                .orientation = xlRowField
                .position = position
            Else
                If Not IsMissing(position) Then
                    If position <> .position And position <= pvt.rowFields.count Then
                        .position = position
                    Else
                        .position = pvt.rowFields.count
                    End If
                End If
            End If
            If visible = False Then
                .orientation = xlHidden
            End If
            Dim nbrFormat As Variant: nbrFormat = FindNumberFormat(pvt, fldName)
            If Len(nbrFormat & vbNullString) > 0 Then
                .dataRange.numberFormat = nbrFormat
            End If
        End With
        pf.Subtotals(1) = showSubTotal
    Else
        Beeper
        Trace "Error 'UpdateRowField': was unable to find pivotfield: " & fldName
    End If

End Function

Public Function UpdateDataField(ByRef pvt As PivotTable, fldName As Variant, visible As Boolean, summaryFn As XlConsolidationFunction, Optional position As Variant)

    Dim pf As PivotField
    Set pf = pvt.PivotFields(fldName)
    If Not pf Is Nothing Then
        With pf
            If visible = True Then
            
                If .orientation <> xlDataField Then
                    position = pvt.dataFields.count + 1
                    .orientation = xlDataField
                    .position = position
                Else
                    If Not IsMissing(position) Then
                        If .position <> position And position <= pvt.dataFields.count Then
                            .position = position
                        End If
                    End If
                End If
                If .Function <> summaryFn Then
                    .Function = summaryFn
                End If
                
                Dim nbrFormat As Variant: nbrFormat = FindNumberFormat(pvt, fldName)
                If Len(nbrFormat & vbNullString) > 0 Then
                    .dataRange.numberFormat = nbrFormat
                End If
            Else
                .orientation = xlHidden
            End If
         End With
    Else
        Beeper
        Trace "Error 'UpdateDataField': was unable to find pivotfield: " & fldName
    End If

End Function

Private Function FindNumberFormat(ByRef pvt As PivotTable, fldName As Variant) As Variant
On Error Resume Next
    Dim lo As ListObject
    Set lo = wt(pvt.SourceData)
    If Not lo Is Nothing Then
        If lo.listRows.count > 0 Then
            If lo.Range.Worksheet.ProtectContents Then pbUnprotectSheet lo.Range.Worksheet
            FindNumberFormat = lo.ListColumns(fldName).DataBodyRange(1, 1).numberFormat
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function UpdatePageField(ByRef pvt As PivotTable, fldName As Variant, visible As Boolean, Optional position As Variant)
        
    Dim pf As PivotField
    Set pf = pvt.PivotFields(fldName)
    If Not pf Is Nothing Then
        DebugPrint fldName & " previous orientation: " & pf.orientation
        With pf
            If .orientation <> xlPageField Then
                position = pvt.pageFields.count + 1
                .orientation = xlPageField
                .position = position
            Else
                If Not IsMissing(position) Then
                    If position <> .position And position <= pvt.pageFields.count Then
                        .position = position
                    End If
                End If
            End If
            If visible = False Then
                .orientation = xlHidden
            End If
            
        End With
    Else
        Beeper
        Trace "Error updating pivot UpdatePageField: " & pvt.Name & ", " & fldName, True
    End If

End Function

Public Function UpdatePivotField(ByRef pvt As PivotTable, fldName As Variant, orientation As XlPivotFieldOrientation, Optional position As Variant, Optional showSubTotal As Boolean = False, Optional inheritFormat As Boolean = False, Optional consolFunction As XlConsolidationFunction)
On Error GoTo E:
    Dim fld As PivotField
    Set fld = GetPivotField(pvt, fldName)
    If Not fld Is Nothing Then
        With fld
            If .orientation <> orientation Then
                .orientation = orientation
            End If
            If Not IsMissing(position) Then
                .position = position
            End If
            If .orientation = xlRowField Then
                .Subtotals(1) = showSubTotal
            End If
            If .orientation = xlDataField Then
                If IsMissing(consolFunction) Then
                    .Function = xlSum
                Else
                    .Function = consolFunction
                End If
            End If
        End With
        If inheritFormat Then
            SetNumberFormat pvt, fld
        End If
    End If

    Exit Function
E:
    Beeper
    Trace "Error in PivotUtil.UpdatePivotfield: " & pvt.Name & " - " & fldName
    Err.Clear

End Function

Public Function SetNumberFormat(pvtTbl As PivotTable, pvtField As PivotField, Optional numFormat As String = vbNullString)
On Error Resume Next
    Dim tstRng As Range
    Set tstRng = wt(pvtTbl.SourceData).ListColumns(pvtField.SourceName).DataBodyRange(1, 1)
    If numFormat = vbNullString Then
        Select Case TypeName(tstRng.value)
            Case "Boolean"
                'nothing
            Case "String"
                'nothing
            Case Else
                pvtField.numberFormat = tstRng.numberFormat
        End Select
    Else
        pvtField.numberFormat = numFormat
    End If
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function GetPivotField(ByRef pvt As PivotTable, fldName As Variant) As PivotField
On Error GoTo E:
    Set GetPivotField = pvt.PivotFields(fldName)
    
    Exit Function
    
E:
    Beeper
    Err.Clear
End Function

Public Function ListPivotFields(ByRef pvt As PivotTable)

    Dim fld As PivotField
    For Each fld In pvt.PivotFields
        DebugPrint fld.Name & ": " & fld.orientation
        
    Next fld

End Function

Public Function DeletePivotTable(pvt As PivotTable)

    If pvt.TableRange2.Worksheet.ProtectContents Then
        pbUnprotectSheet pvt.TableRange2.Worksheet
    End If
    
    pvt.TableRange2.Clear

End Function

