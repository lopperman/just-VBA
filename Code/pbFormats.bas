Attribute VB_Name = "pbFormats"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' pbFormats v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' Misc utilities related to Worksheet and Object Formats
'
' @module pbFormats
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'
'   Concept -- Produce Code needed to re-create specific formats of selected range
'   or object(s)
'
'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~


Option Explicit
Option Compare Text
Option Base 1

Public Type FormatDetail
    Id As Long
    WkshtName As String
    RangeAddr As String
    PropName As String
    PropType As String
    PropVal As Variant
    AsCode As String
    AddlInfo() As Variant
End Type

Public Enum FormatType
    fmtFont = 2 ^ 0
    fmtBorders = 2 ^ 1
    fmtInterior = 2 ^ 2
    fmtFormulas = 2 ^ 3
    fmtMergeAreas = 2 ^ 4
    fmtConditionalFormatting = 2 ^ 5
    fmtRangeInfo = 2 ^ 6
End Enum

Public Function AnalyzeRangeFormats(ByVal Target As Range, formats As FormatType)
    Dim tmpCol As New Collection
    If EnumCompare(formats, FormatType.fmtFont) Then AddFontInfo Target, tmpCol
    If EnumCompare(formats, FormatType.fmtFormulas) Then AddFormulasInfo Target, tmpCol
    If EnumCompare(formats, FormatType.fmtMergeAreas) Then AddMergeInfo Target, tmpCol
    If EnumCompare(formats, FormatType.fmtRangeInfo) Then AddRangeInfo Target, tmpCol

End Function

Public Function AddRangeInfo(ByVal Target As Range, Optional ByRef tCol As Collection)
    On Error Resume Next
    Dim rngArea As Range, cl As Range
    For Each rngArea In Target.Areas
        For Each cl In rngArea
            If cl.MergeCells = False Or cl.Address = cl.MergeArea(1, 1).Address Then
                With cl
                    If cl.MergeCells Then
                        Debug.Print Concat("* RANGE FORMAT * ", .Worksheet.Name, "!", .Address, " to ", .Worksheet.Name, "!", .MergeArea(.MergeArea.Rows.Count, .MergeArea.Columns.Count).Address)
                    Else
                        Debug.Print Concat("* RANGE FORMAT * ", .Worksheet.Name, "!", .Address)
                    End If
                    If Not IsNull(.HorizontalAlignment) Then Debug.Print Concat(vbTab, "TypeName: (xlHAlign) ", TypeName(.HorizontalAlignment), " .HorizontalAlignment=", .HorizontalAlignment)
                    If Not IsNull(.VerticalAlignment) Then Debug.Print Concat(vbTab, "TypeName: (xlVAlign) ", TypeName(.VerticalAlignment), " .VerticalAlignment=", .VerticalAlignment)
                    If Not IsNull(.IndentLevel) Then Debug.Print Concat(vbTab, "TypeName: ", TypeName(.IndentLevel), " .IndentLevel=", .IndentLevel)
                    If Not IsNull(.Interior) Then
                        If Not IsNull(.Interior.color) Then Debug.Print Concat(vbTab, "TypeName: ", TypeName(.Interior.color), " .Interior.Color=", .Interior.color)
                        If Not IsNull(.Interior.ColorIndex) Then Debug.Print Concat(vbTab, "TypeName: ", TypeName(.Interior.ColorIndex), " .Interior.ColorIndex=", .Interior.ColorIndex)
'                        If Not IsNull(.Interior.Gradient) Then
'                            'Dim tGrad As colorin
'                        End If
                        
                    End If
                End With
            End If
        Next
    Next
End Function

Public Function AddMergeInfo(ByVal Target As Range, Optional ByRef tCol As Collection)
    On Error Resume Next
    Dim rngArea As Range, cl As Range
    For Each rngArea In Target.Areas
        For Each cl In rngArea
            If cl.MergeCells Then
                If cl.Address = cl.MergeArea(1, 1).Address Then
                    With cl
                        Debug.Print Concat("* MERGE CELLS * ", .Worksheet.Name, "!", .Address, " to ", .Worksheet.Name, "!", .MergeArea(.MergeArea.Rows.Count, .MergeArea.Columns.Count).Address)
                        Debug.Print Concat(vbTab, "Merge Cols: ", .MergeArea.Columns.Count)
                        Debug.Print Concat(vbTab, "Merge Rows: ", .MergeArea.Rows.Count)
                    End With
                End If
            End If
        Next
    Next
End Function

Public Function AddFormulasInfo(ByVal Target As Range, Optional ByRef tCol As Collection)
    On Error Resume Next
    Dim rngArea As Range, cl As Range
    For Each rngArea In Target.Areas
        For Each cl In rngArea
            With cl
                If .HasFormula Then
                    Debug.Print Concat("* FORMULA * ", .Worksheet.Name, "!", .Address)
                    Debug.Print Concat(vbTab, "(A1 Style) .Formula", """", cl.formula, """")
                    If StringsMatch(cl.formula, """", smContains) Then
                        Debug.Print Concat(vbTab, "(A1 Style - VBA Code) .Formula", """", Replace(cl.formula, """", """"""), """")
                    End If
                    Debug.Print Concat(vbTab, "(R1C1 Style) .Formula2R1C1", """", cl.Formula2R1C1, """")
                    If StringsMatch(cl.Formula2R1C1, """", smContains) Then
                        Debug.Print Concat(vbTab, "(R1C1 Style - VBA Code) .Formula", """", Replace(cl.Formula2R1C1, """", """"""), """")
                    End If
                End If
            End With
        Next cl
    Next
End Function

Public Function AddFontInfo(ByVal Target As Range, Optional ByRef tCol As Collection)
    On Error Resume Next
    Dim rngArea As Range, cl As Range
    For Each rngArea In Target.Areas
        For Each cl In rngArea
            With cl
                If cl.MergeCells = False Or cl.Address = cl.MergeArea(1, 1).Address Then
                    Debug.Print Concat("* FONT * ", .Worksheet.Name, "!", .Address)
                    If Not IsNull(.Font.Name) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.Name), " .Font.Name=", cl.Font.Name)
                    End If
                    If Not IsNull(.Font.Size) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.Size), " .Font.Size=", cl.Font.Size)
                    End If
                    If Not IsNull(.Font.Bold) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.Bold), " .Font.Bold=", cl.Font.Bold)
                    End If
                    If Not IsNull(.Font.Background) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.Background), " ", ".Background=", cl.Font.Background)
                    End If
                    If Not IsNull(.Font.color) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.color), " ", ".Color=", cl.Font.color)
                    End If
                    If Not IsNull(.Font.ColorIndex) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.ColorIndex), " ", ".ColorIndex=", cl.Font.ColorIndex)
                    End If
                    If Not IsNull(.Font.FontStyle) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.FontStyle), " ", ".FontStyle=", cl.Font.FontStyle)
                    End If
                    If Not IsNull(.Font.Italic) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.Italic), " ", ".Italic=", cl.Font.Italic)
                    End If
                    If Not IsNull(.Font.Strikethrough) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.Strikethrough), " ", ".Strikethrough=", cl.Font.Strikethrough)
                    End If
                    If Not IsNull(.Font.Subscript) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.Subscript), " ", ".Subscript=", cl.Font.Subscript)
                    End If
                    If Not IsNull(.Font.Superscript) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.Superscript), " ", ".Superscript=", cl.Font.Superscript)
                    End If
                    If Not IsNull(.Font.ThemeColor) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.ThemeColor), " ", ".ThemeColor=", cl.Font.ThemeColor)
                    End If
                    If Not IsNull(.Font.ThemeFont) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.ThemeFont), " ", ".ThemeFont=", cl.Font.ThemeFont)
                    End If
                    If Not IsNull(.Font.TintAndShade) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.TintAndShade), " ", ".TintAndShade=", cl.Font.TintAndShade)
                    End If
                    If Not IsNull(.Font.Underline) Then
                        Debug.Print Concat(vbTab, "TypeName: ", TypeName(cl.Font.Underline), " ", ".Underline=", cl.Font.Underline)
                    End If
                End If
            End With
        Next
    Next rngArea
    
End Function
