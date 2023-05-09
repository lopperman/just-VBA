Attribute VB_Name = "pbRangeUpdate"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Range Update Utility
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbRangeUpdate.bas
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Option Explicit
Option Compare Text
Option Base 1

Public Enum DblClickUpdateMode
    dcUnknown = 0
    dcFillDown = 1
    dcReplaceSelected = 2
    dcReplaceBlanks = 3
End Enum
Private Enum ecComparisonType
    ecOR = 0 'default
    ecAnd
End Enum
Private Enum ftInputBoxType
    ftibFormula = 0
    ftibNumber = 1
    ftibString = 2
    ftibLogicalValue = 4
    ftibCellReference = 8
    ftibErrorValue = 16
    ftibArrayOfValues = 64
End Enum
Private Enum strMatchEnum
    smEqual = 0
    smNotEqualTo = 1
    smContains = 2
    smStartsWithStr = 3
    smEndWithStr = 4
End Enum


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   If [clickedRange] is empty, find the first non-empty value
'       looking up, and fill all blanks from the found value,
'       to the cell the was clicked.
'       - Ignores header range if clicked cell is part of a
'         ListObject, or if [headersRow], if [clickedRange] is
'         not in ListObject and [headersRow] is > 0
'   If [clickeRange] is not empty, provides option to
'   manually enter a value and replace ALL values in column
'   that matched the [clickedRange] value, with the new
'   manually entered value
'   Returns Range of Updated Cells
'       * NOTE:  If the dcReplaceBlanks or
'         dcReplaceSelected modes were used, the returned
'         range may be Unioned (Contain multiple Areas)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'
'   USAGE
'
'       Call from a Worksheet_BeforeDoubleClick Event
'       e.g.:
'       in this hander: Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'           Cancel = True
'           Dim tmpUpdateRng As Range
'           If Target.Count = 1 Then
'                Set tmpUpdateRng = RangeUpd(Target)
'           End if
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Public Function RangeUpd(ByVal clickedRange As Range, Optional headersRow As Long, Optional countZeroAsEmpty As Boolean = True) As Range
On Error GoTo E:
    Dim failed As Boolean
    Dim evts As Boolean, scrn As Boolean
    evts = Application.EnableEvents
    scrn = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim updateMode As DblClickUpdateMode
    Dim firstValidWorksheetRow As Long, clickedRow As Long, clickedCol As Long
    Dim fillStartRow As Long, isInvalid As Boolean, lastRow As Long
    Dim cellEmpty As Boolean, cellValue As Variant
    Dim newValue As Variant, tmpVal As Variant
    Dim missingFillVal As Boolean, i As Long
    Dim updatedCell As Boolean

    
    updateMode = dcUnknown
    clickedRow = clickedRange.Row
    clickedCol = clickedRange.Column
    If clickedRange(1, 1).Value = 0 Then
        If countZeroAsEmpty = False Then
            cellValue = clickedRange(1, 1).Value
        End If
    Else
        cellValue = clickedRange(1, 1).Value
    End If
    cellEmpty = StringsMatch(cellValue, "")
    
    With clickedRange
        If Not .ListObject Is Nothing Then
            With .ListObject.HeaderRowRange
                headersRow = .Row + (1 - .Rows.Count)
                firstValidWorksheetRow = headersRow + 1
            End With
            lastRow = .ListObject.ListRows(.ListObject.ListRows.Count).Range.Row
        ElseIf headersRow > 0 Then
            firstValidWorksheetRow = headersRow + 1
        Else
            headersRow = 1
            firstValidWorksheetRow = headersRow + 1
        End If
        If lastRow = 0 Then
            lastRow = clickedRange.Worksheet.UsedRange.Row + clickedRange.Worksheet.UsedRange.Rows.Count - 1
        End If
    End With
    
    If clickedRow <= headersRow Then
        isInvalid = True
    Else
        If cellEmpty Then
            For i = clickedRow - 1 To headersRow + 1 Step -1
                tmpVal = clickedRange.Worksheet.Cells(i, clickedCol).Value
                If tmpVal = 0 And countZeroAsEmpty Then
                    tmpVal = ""
                End If
                If Len(tmpVal) > 0 Then
                    newValue = tmpVal
                    fillStartRow = i + 1
                    updateMode = dcFillDown
                    Exit For
                End If
            Next i
        Else
            updateMode = dcReplaceSelected
        End If
    End If
    
    If updateMode = dcUnknown Then
        MsgBox_FT "There aren't any values above that can be used to fill down.", vbOKOnly + vbInformation, "OOPS"
    ElseIf updateMode = dcFillDown Then
        If MsgBox_FT("Fill in blanks cells with '" & newValue & "'?", vbYesNo + vbDefaultButton1 + vbQuestion, "UPDATE COLUMN VALUES") = vbYes Then
            For i = fillStartRow To clickedRow
                clickedRange.Worksheet.Cells(i, clickedCol).Value = newValue
                'Worksheet_Change Me.Cells(i, clickedRange.column)
            Next i
            Set RangeUpd = clickedRange.Worksheet.Cells(fillStartRow, clickedCol)
            Set RangeUpd = RangeUpd.Resize(RowSize:=1 + (clickedRow - fillStartRow))
        End If
    ElseIf updateMode = dcReplaceSelected Then
        newValue = InputBox_FT("Enter Value to replace '" & cellValue & "' in current column", "REPLACE MATCHING COLUMN VALUES")
        If IsObject(newValue) Then
            'just in case user selects a range
            MsgBox "You entered an invalid value"
        ElseIf Len(newValue) = 0 Then
            If MsgBox_FT("You entered [blank].  Would you like to clear all of the cells in current column that contain '" & cellValue & "'?", vbYesNo + vbDefaultButton2 + vbQuestion, "Clear Values?") = vbYes Then
                For i = firstValidWorksheetRow To lastRow
                    updatedCell = False
                    If IsNumeric(clickedRange.Worksheet.Cells(i, clickedCol).Value) Then
                        If clickedRange.Worksheet.Cells(i, clickedCol).Value = cellValue Then
                           clickedRange.Worksheet.Cells(i, clickedCol).ClearContents
                           updatedCell = True
                            'TODO add to range
                        End If
                    ElseIf StringsMatch(clickedRange.Worksheet.Cells(i, clickedCol).Value, cellValue) Then
                           clickedRange.Worksheet.Cells(i, clickedCol).ClearContents
                           updatedCell = True
                            'TODO add to range
                    End If
                    If updatedCell Then
                        If RangeUpd Is Nothing Then
                            Set RangeUpd = clickedRange.Worksheet.Cells(i, clickedCol)
                        Else
                            Set RangeUpd = Union(RangeUpd, clickedRange.Worksheet.Cells(i, clickedCol))
                        End If
                    End If
                Next i
            End If
        Else
            If MsgBox_FT("Would you like to REPLACE '" & cellValue & "', with '" & newValue & "'? for ALL the matching cells in the current column?", vbYesNo + vbDefaultButton2 + vbQuestion, "Replace Values?") = vbYes Then
                For i = firstValidWorksheetRow To lastRow
                    updatedCell = False
                    If IsNumeric(clickedRange.Worksheet.Cells(i, clickedCol).Value) Then
                        If clickedRange.Worksheet.Cells(i, clickedCol).Value = cellValue Then
                           clickedRange.Worksheet.Cells(i, clickedCol).Value = newValue
                           updatedCell = True
                        End If
                    ElseIf StringsMatch(clickedRange.Worksheet.Cells(i, clickedCol).Value, cellValue) Then
                           clickedRange.Worksheet.Cells(i, clickedCol).Value = newValue
                           updatedCell = True
                    End If
                    If updatedCell Then
                        If RangeUpd Is Nothing Then
                            Set RangeUpd = clickedRange.Worksheet.Cells(i, clickedCol)
                        Else
                            Set RangeUpd = Union(RangeUpd, clickedRange.Worksheet.Cells(i, clickedCol))
                        End If
                    End If
                Next i
            End If
        End If
    End If

Finalize:
    On Error Resume Next
    Application.EnableEvents = evts
    Application.ScreenUpdating = scrn
    
    Exit Function
E:
    failed = True
    Application.EnableEvents = evts
    Application.ScreenUpdating = scrn
    'Implement Your Own Error Handler
    Err.Raise Err.Number
    Resume Finalize:
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'
'   UTILITY METHODS FROM pbCommon
'   https://github.com/lopperman/just-VBA/blob/main/Code_NoDependencies/pbCommon.bas
'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~


' ~~~~~~~~~~   INPUT BOX   ~~~~~~~~~~'
Private Function InputBox_FT(prompt As String, Optional title As String = "Financial Tool - Input Needed", Optional default As Variant, Optional inputType As ftInputBoxType, Optional useVBAInput As Boolean = False) As Variant
    Beep
    
    If useVBAInput Then
        InputBox_FT = VBA.InputBox(prompt, title:=title, default:=default)
    Else
        If inputType > 0 Then
            InputBox_FT = Application.InputBox(prompt, title:=title, default:=default, Type:=inputType)
        Else
            InputBox_FT = Application.InputBox(prompt, title:=title, default:=default)
        End If
    End If
    
    DoEvents
End Function
' ~~~~~~~~~~   MSG BOX   ~~~~~~~~~~'
Private Function MsgBox_FT(prompt As String, Optional buttons As VbMsgBoxStyle = vbOKOnly, Optional title As Variant) As Variant
    Dim evts As Boolean: evts = Application.EnableEvents
    Dim screenUpd As Boolean: screenUpd = Application.ScreenUpdating
    Application.EnableEvents = False
    If Not EnumCompare(buttons, vbSystemModal) Then buttons = buttons + vbSystemModal
    If Not EnumCompare(buttons, vbMsgBoxSetForeground) Then buttons = buttons + vbMsgBoxSetForeground
    Beep
    If Not ThisWorkbook.ActiveSheet Is Application.ActiveSheet Then
        Application.ScreenUpdating = True
        ThisWorkbook.Activate
        DoEvents
        Application.ScreenUpdating = screenUpd
    End If
    MsgBox_FT = MsgBox(prompt, buttons, title)
    Application.EnableEvents = evts
    DoEvents
End Function

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


