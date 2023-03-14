Attribute VB_Name = "pbUpdateLstObj"
Option Explicit
Option Compare Text
Option Base 1



Public Enum strMatchEnum
    smEqual = 0
    smNotEqualTo = 1
    smContains = 2
    smStartsWithStr = 3
    smEndWithStr = 4
End Enum

Public Function FindAndReplaceMatchingCols(lstColumnName As String, oldVal, newVal, valType As VbVarType, Optional wkbk As Workbook, Optional strMatch As strMatchEnum = strMatchEnum.smEqual) As Boolean
    If wkbk Is Nothing Then
        Set wkbk = ThisWorkbook
    End If
    Dim tWS As Worksheet, tLO As ListObject
    For Each tWS In wkbk.Worksheets
        For Each tLO In tWS.ListObjects
            If HasData(tLO) And ListColumnIndex(tLO, lstColumnName) > 0 Then
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
    Dim ERR_REPROTECT_SHEET As Long: ERR_REPROTECT_SHEET = 1004
    
    Dim colIdx As Long
    Dim tmpARR, changedCount As Long, tmpIdx As Long, tmpValue As Variant
    Dim itemValid As Boolean

    If StringsMatch(TypeName(lstColIdxOrName), "String") Then
        colIdx = ListColumnIndex(lstObj, CStr(lstColIdxOrName))
    End If
    If colIdx > 0 And colIdx <= lstObj.ListColumns.Count And HasData(lstObj) Then
        Select Case valType
            Case VbVarType.vbArray, VbVarType.vbDataObject, VbVarType.vbEmpty, VbVarType.vbError, VbVarType.vbObject, VbVarType.vbUserDefinedType, VbVarType.vbVariant
                Err.Raise ERR_TYPE_MISMATCH, Source:="pbListObj.FindAndReplaceListCol", Description:="VbVarType: " & valType & " is not supported"
        End Select
    
        If lstObj.ListColumns(colIdx).DataBodyRange.Count = 1 Then
            ReDim tmpARR(1 To 1, 1 To 1)
            tmpARR(1, 1) = lstObj.ListColumns(colIdx).DataBodyRange.value
        Else
            tmpARR = lstObj.ListColumns(colIdx).DataBodyRange.value
        End If
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
            If lstObj.Range.Worksheet.ProtectContents Then
                ' You must call your code to PROTECT worksheet: lst.Range.Worksheet
                ' make sure UserInterfaceOnly:=True is included
                ProtectSht lstObj.Range.Worksheet
            End If
            On Error Resume Next
            pbSafeUpdate.UpdateListObjRange lstObj.ListColumns(colIdx).DataBodyRange, tmpARR
            If Err.number <> 0 Then
                Err.Clear
                'Try to reprotect sheet and try once more
                ProtectSht lstObj.Range.Worksheet
                pbSafeUpdate.UpdateListObjRange lstObj.ListColumns(colIdx).DataBodyRange, tmpARR
                If Err.number <> 0 Then
                    On Error GoTo 0
                    Err.Raise ERR_REPROTECT_SHEET, "pbUpdateLstObj.FindAndReplaceListCol", "Not able to change values in Worksheet.  Either Unprotect or Re-protect with UserInterfaceOnly:=True"
                End If
            End If
        End If
    End If
        
Finalize:
    On Error Resume Next
    If Not failed Then
        FindAndReplaceListCol = changedCount
    End If
    
    If Err.number <> 0 Then Err.Clear
    Exit Function
E:
    failed = True
    'ErrorCheck
    'Your custom error handling
    Resume Finalize:
End Function

Private Function ProtectSht(ws As Worksheet)

    'YOUR CODE TO REPROTECT WORKSHEET WITH 'UserInterfaceOnly:=True

End Function

Public Function HasData(lstObj As ListObject) As Boolean
    HasData = lstObj.ListRows.Count > 0
End Function

Public Function StringsMatch( _
    ByVal checkString As Variant, ByVal _
    validString As Variant, _
    Optional smEnum As strMatchEnum = strMatchEnum.smEqual, _
    Optional compMethod As VbCompareMethod = vbTextCompare) As Boolean
    
'       IF NEEDED, PUT THIS ENUM AT TOP OF A STANDARD MODULE
        'Public Enum strMatchEnum
        '    smEqual = 0
        '    smNotEqualTo = 1
        '    smContains = 2
        '    smStartsWithStr = 3
        '    smEndWithStr = 4
        'End Enum
        
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

Public Function ListColumnIndex(ByRef lstObj As ListObject, lstColName As String) As Long
    Dim lstCol As ListColumn
    For Each lstCol In lstObj.ListColumns
        If StringsMatch(lstCol.Name, lstColName) Then
            ListColumnIndex = lstCol.Index
            Exit For
        End If
    Next lstCol
End Function

Private Function ConcatDelim(ByVal delimeter As String, ParamArray Items() As Variant) As String
    ConcatDelim = Join(Items, delimeter)
End Function

