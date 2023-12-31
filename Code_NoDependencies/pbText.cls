VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  A class for editing text in Excel cells
''
''   **  See: https://github.com/lopperman/just-VBA/
'' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  author (c) Paul Brower https://github.com/lopperman/
''  module pbText.cls
''  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  collection that stores list of formats and text used for writing to range
Private ltext As New Collection
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  collection that contains information about building borders
Private lborders As New Collection
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  property used to set Horizontal Text Alignmnet for Output Range
''  Defaults to 'xlHAlignGeneral', can be changed with 'Configure' method
''  anytime prior to calling 'WriteText' method
Private lHAlign As XlHAlign
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  property used to set Vertical Text Alignmnet for Output Range
''  Defaults to 'xlVAlignCenter', can be changed with 'Configure' method
''  anytime prior to calling 'WriteText' method
Private lVAlign As XlVAlign
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  property to store color ([Range].Interior.Color) for Output Range
''  background.  Defaults to 'xlNone'
Private lInteriorColor As Variant
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Property to store where Output Range cells will be merged.
''  Defaults to 'True', can be changed with 'Configure' method
''  anytime prior to calling 'WriteText' method
Private lMergeCells As Boolean
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Property to store 'WrapText' property for Output Range
''  Defaults to 'True', can be changed with 'Configure' method
''  anytime prior to calling 'WriteText' method
Private lWrapText As Boolean
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Property to store whether all formats will be removed from
''  Output Range, prior to writing out text/formats.
''  Defaults to 'True', can be changed with 'Configure' method
''  anytime prior to calling 'WriteText' method
Private lClearFormats As Boolean
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private Enum used to indicate type of information being added from
''  'AddLine' or 'AddLineFormatted' methods
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Enum itemType
    itTextNewLine = 0
    itTextAppend
    itFontColor
    itFontSize
    itFontName
    itFontStyle
End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private Bitwise Enum used to store information about how to format
''  a segment of text
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Enum FontItemStyle
    fisDefault = 0
    fisBold = 2 ^ 0
    fisItalic = 2 ^ 1
    fisUnderline = 2 ^ 2
    fisStrikethrough = 2 ^ 3
    fisSuperscript = 2 ^ 4
    fisSubscript = 2 ^ 5
End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private Enum used for determining comparison type when checking
''  for existence of bitwise enum values
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Enum tComparisonType
    tOR = 0 'default
    tAnd
End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private Enum used for determining validity of bitwise enum value
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Enum tFlagEnumModify
    tVerifyEnumExists
    tVerifyEnumRemoved
End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private Enum used by custom 'tStringsMatch' method
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Enum tstrMatchEnum
    tsmEqual = 0
    tsmNotEqualTo = 1
    tsmContains = 2
    tsmStartsWithStr = 3
    tsmEndWithStr = 4
End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private Type for simplifying properties for borders that that are stored
''  in an array
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Type pbTypeBorder
    item As XlBordersIndex
    borderColor As Double
    borderWeight As XlBorderWeight
    borderStyle As XlLineStyle
End Type

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private Type for simplifying properties for text/formatting information
''  that is stored in an array
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Type pbTypeItem
    tiType As itemType
    tiValue As Variant
End Type

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  OPTIONAL - Customize formatting options for output range
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function Configure( _
    Optional verticalAlign As XlVAlign = xlVAlignCenter _
    , Optional horizontalAlign As XlHAlign = xlHAlignGeneral _
    , Optional clearRangeFormatsOnWrite As Boolean = True _
    , Optional mergeCellsOnWrite As Boolean = True _
    , Optional wrapText As Boolean = True _
    , Optional rangeInteriorColor) As pbText
    
    lVAlign = verticalAlign
    lHAlign = horizontalAlign
    lClearFormats = clearRangeFormatsOnWrite
    lMergeCells = mergeCellsOnWrite
    If IsMissing(rangeInteriorColor) Then rangeInteriorColor = xlNone
    lInteriorColor = rangeInteriorColor
    lWrapText = wrapText
    
    Set Configure = Me
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  OPTIONAL - Configure format for a 'Border Around' Output Range
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function AddBorderAround( _
    Optional borderColor = xlColorIndexAutomatic _
    , Optional borderWeight As XlBorderWeight = XlBorderWeight.xlThin _
    , Optional borderStyle As XlLineStyle = XlLineStyle.xlContinuous _
    ) As pbText

    AddBorder xlEdgeTop, borderColor, borderWeight, borderStyle
    AddBorder xlEdgeBottom, borderColor, borderWeight, borderStyle
    AddBorder xlEdgeLeft, borderColor, borderWeight, borderStyle
    AddBorder xlEdgeRight, borderColor, borderWeight, borderStyle

    Set AddBorderAround = Me
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  OPTIONAL - Configure format for any border 'item' in Output Range
''  e.g. To add top border only:
''  AddBorder XlBordersIndex.xlEdgeTop, borderWeight:=XlBorderWeight.xlMedium
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function AddBorder( _
    item As XlBordersIndex _
    , Optional borderColor = xlColorIndexAutomatic _
    , Optional borderWeight As XlBorderWeight = XlBorderWeight.xlThin _
    , Optional borderStyle As XlLineStyle = XlLineStyle.xlContinuous _
    ) As pbText
    
    Dim bItem, b As pbTypeBorder, i As Long
    If lborders.Count > 0 Then
        For i = lborders.Count To 1 Step -1
            b = GetBorder(lborders(i))
            If b.item = item Then
                lborders.Remove (i)
                Exit For
            End If
        Next i
    End If
    lborders.Add Array(item, borderColor, borderWeight, borderStyle)
    Set AddBorder = Me
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Store (explicitly unformatted) text to be written
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function AddLine(ByVal txt As String) As pbText
    Set AddLine = AddLineFormatted(txt)
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Store (explicitly FORMATTED) text with formats to be written
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function AddLineFormatted( _
    ByVal txt As String _
    , Optional fontColor _
    , Optional fontName _
    , Optional fontSize _
    , Optional fontBold _
    , Optional fontItalic _
    , Optional fontUnderline _
    , Optional fontStrikethrough _
    , Optional fontSuperscript _
    , Optional fontSubscript) As pbText

    If HasText Then
        txt = vbNewLine & txt
    End If
    
    If Not IsMissing(fontColor) Then AddItem itFontColor, fontColor
    If Not IsMissing(fontName) Then AddItem itFontName, fontName
    If Not IsMissing(fontSize) Then AddItem itFontSize, fontSize
    Dim tStyle As FontItemStyle
    If Not IsMissing(fontBold) Then
        If fontBold = True Then tStyle = tEnumModify(tStyle, FontItemStyle.fisBold, tVerifyEnumExists)
    End If
    If Not IsMissing(fontItalic) Then
        If fontItalic = True Then tStyle = tEnumModify(tStyle, FontItemStyle.fisItalic, tVerifyEnumExists)
    End If
    If Not IsMissing(fontUnderline) Then
        If fontUnderline = True Then tStyle = tEnumModify(tStyle, FontItemStyle.fisUnderline, tVerifyEnumExists)
    End If
    If Not IsMissing(fontStrikethrough) Then
        If fontStrikethrough = True Then tStyle = tEnumModify(tStyle, FontItemStyle.fisStrikethrough, tVerifyEnumExists)
    End If
    If Not IsMissing(fontSuperscript) Then
        If fontSuperscript = True Then tStyle = tEnumModify(tStyle, FontItemStyle.fisSuperscript, tVerifyEnumExists)
    End If
    If Not IsMissing(fontSubscript) Then
        If fontSubscript = True Then tStyle = tEnumModify(tStyle, FontItemStyle.fisSubscript, tVerifyEnumExists)
    End If
    AddItem itFontStyle, tStyle
    AddItem itTextNewLine, txt
    Set AddLineFormatted = Me
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Write text with specified formation to [writeTo] range
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function WriteText(writeTo As Range) As pbText
    On Error Resume Next
    Dim rng As Range
    If lClearFormats Then
        writeTo.Clear
    Else
        writeTo.ClearContents
    End If
    If writeTo.Count > 1 And lMergeCells = False Then
        Set rng = writeTo(1, 1)
    ElseIf writeTo.Count > 1 And lMergeCells = True Then
        Set rng = writeTo
        If IsNull(rng.mergeCells) Then
            rng.Merge
        ElseIf rng.mergeCells = False Then
            rng.Merge
        End If
    Else
        Set rng = writeTo
    End If
    With rng
        .HorizontalAlignment = lHAlign
        .VerticalAlignment = lVAlign
        If ltext.Count > 0 Then
            RenderText rng(1, 1)
        End If
        If lInteriorColor <> xlNone Then
            .Interior.color = lInteriorColor
        End If
        If lborders.Count > 0 Then
            ApplyBorders writeTo
        End If
        If lWrapText Then
            rng.wrapText = True
        End If
    End With
    Clear
    Set WriteText = Me
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  private method used to build 'pbTypeBorder' Type from Array
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function GetBorder(v As Variant) As pbTypeBorder
    Dim resp As pbTypeBorder
    resp.item = v(LBound(v))
    resp.borderColor = v(LBound(v) + 1)
    resp.borderWeight = v(LBound(v) + 2)
    resp.borderStyle = v(LBound(v) + 3)
    GetBorder = resp
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  private method used to build 'pbTypeItem' Type from Array
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function GetItem(v As Variant) As pbTypeItem
    Dim resp As pbTypeItem
    resp.tiType = v(LBound(v))
    resp.tiValue = v(LBound(v) + 1)
    GetItem = resp
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  private method used to add text/format information to 'lText'
''  collection
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function AddItem(iKeyType As itemType, iItemValue As Variant)
    Dim ti() As Variant
    ReDim ti(1 To 2)
    ti(1) = iKeyType
    ti(2) = iItemValue
    ltext.Add ti
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  private method used to clear previously cached text and formatting
''  information, AFTER is has been written to Output Range
''  (Enables re-use of class without have to re-create and re-configure)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function Clear()
    Set ltext = Nothing
    Set ltext = New Collection
    Set lborders = Nothing
    Set lborders = New Collection
    SetDefaults
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  private 'helper' Property used to return whether or not previous
''  text has been cached -- used to help determine new line should be
''  added
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Property Get HasText() As Boolean
    Dim tArr As Variant
    For Each tArr In ltext
        If tArr(LBound(tArr)) = itTextNewLine Then
            HasText = True
            Exit For
        End If
    Next
End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private Method used to render Output Range border/s
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function ApplyBorders(rng As Range)
    Dim bArr As Variant, b As pbTypeBorder
    For Each bArr In lborders
        b = GetBorder(bArr)
        With rng.borders(b.item)
            .color = b.borderColor
            .LineStyle = b.borderStyle
            .Weight = b.borderWeight
        End With
    Next
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private method -- call from 'WriteText' that organizes and renders
''  formatting and text
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function RenderText(toCell As Range)
    On Error Resume Next
    '' get indexes and count of continuous text items
    '' for each continous groups of text, determine format that should be applied based on previous format and style entries
    Dim tArr() As Variant, tType As itemType, tVal As Variant
    Dim outputs As New Collection
    '' outputs structure (1)=formatStart, (2) = formatEnd, (3) = text
    Dim i As Long
    
    Dim workingText As String
    Dim workingFormatStart As Long
    Dim workingFormatEnd As Long
    Dim lastType As itemType
    Dim tOutput() As Variant

    For i = 1 To ltext.Count
        tArr = ltext(i)
        tType = tArr(1)
        tVal = tArr(2)
        If i = 1 Then
            If tType = itTextNewLine Or tType = itTextAppend Then
                workingFormatStart = -1
                workingFormatEnd = -1
                workingText = CStr(tVal)
            Else
                workingFormatStart = 1
                workingFormatEnd = 1
            End If
        Else
            If tType = itTextNewLine Then
                If lastType = itTextNewLine Then
                    workingText = workingText & CStr(tVal)
                Else
                    workingText = CStr(tVal)
                End If
            Else
                '' current item is formatting
                If lastType = itTextNewLine Then
                    ReDim tOutput(1 To 3)
                    tOutput(1) = workingFormatStart
                    tOutput(2) = workingFormatEnd
                    tOutput(3) = workingText
                    outputs.Add tOutput
                    workingFormatStart = i
                    workingFormatEnd = i
                    workingText = ""
                Else
                    workingFormatEnd = i
                End If
            End If
        End If
        lastType = tType
        If i = ltext.Count Then
            '' if last thing is formatting,ignore it
            If tType = itTextNewLine Then
                ReDim tOutput(1 To 3)
                tOutput(1) = workingFormatStart
                tOutput(2) = workingFormatEnd
                tOutput(3) = workingText
                outputs.Add tOutput
            End If
        End If
    Next i
    Dim tOut As Variant
    Dim defStyle As Style
    Set defStyle = ThisWorkbook.Styles("Normal")
    For Each tOut In outputs
        If Len(tOut(3)) > 0 Then
            Dim nextPos As Long, nextLen As Long
            nextPos = toCell.Characters().Count + 1
            nextLen = Len(tOut(3))
            toCell.Characters(nextPos).Insert CStr(tOut(3))
            With toCell.Characters(nextPos, nextLen)
                .font.size = defStyle.font.size
                .font.Name = defStyle.font.Name
                .font.color = defStyle.font.color
                .font.Bold = defStyle.font.Bold
                .font.Italic = defStyle.font.Italic
                .font.Underline = defStyle.font.Underline
                .font.Subscript = defStyle.font.Subscript
                .font.Strikethrough = defStyle.font.Strikethrough
                .font.Superscript = defStyle.font.Superscript
                
                If tOut(1) > 0 And tOut(2) > 0 Then
                    
                    For i = CLng(tOut(1)) To CLng(tOut(2))
                        Dim tK, tV
                        tK = ltext(i)(1)
                        tV = ltext(i)(2)
                        If tK = itemType.itFontColor Then
                            .font.color = tV
                        ElseIf tK = itemType.itFontName Then
                            .font.Name = tV
                        ElseIf tK = itemType.itFontSize Then
                            .font.size = tV
                        ElseIf tK = itemType.itFontStyle Then
                            If tEnumCompare(tV, FontItemStyle.fisBold) Then .font.Bold = True
                            If tEnumCompare(tV, FontItemStyle.fisItalic) Then .font.Italic = True
                            If tEnumCompare(tV, FontItemStyle.fisStrikethrough) Then .font.Strikethrough = True
                            If tEnumCompare(tV, FontItemStyle.fisUnderline) Then .font.Underline = True
                            If tEnumCompare(tV, FontItemStyle.fisSuperscript) Then .font.Superscript = True
                            If tEnumCompare(tV, FontItemStyle.fisSubscript) Then .font.Subscript = True
                        End If
                    Next i
                End If
            End With
        End If
    Next
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  FLAG ENUM COMPARE
''  (Private Local copy of 'EnumCompare' from just-VBA pbCommon module
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function tEnumCompare(theEnum As Variant, enumMember As Variant, Optional ByVal iType As tComparisonType = tComparisonType.tOR) As Boolean
    Dim c As Long
    c = theEnum And enumMember
    tEnumCompare = IIf(iType = tOR, c <> 0, c = enumMember)
End Function
Private Function tEnumModify(theEnum, enumMember, modifyType As tFlagEnumModify) As Long
    Dim Exists As Boolean
    Exists = tEnumCompare(theEnum, enumMember)
    If Exists And modifyType = tVerifyEnumRemoved Then
        theEnum = theEnum - enumMember
    ElseIf Exists = False And modifyType = tVerifyEnumExists Then
        theEnum = theEnum + enumMember
    End If
    tEnumModify = theEnum
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  STRING COMPARISON UTILITY
''  (Private Local copy of 'StringsMatch' from just-VBA pbCommon module
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function tStringsMatch( _
    ByVal checkString As Variant, ByVal _
    validString As Variant, _
    Optional smEnum As tstrMatchEnum = tstrMatchEnum.tsmEqual, _
    Optional compMethod As VbCompareMethod = vbTextCompare) As Boolean
    Dim str1, str2
    str1 = CStr(checkString)
    str2 = CStr(validString)
    Select Case smEnum
        Case tstrMatchEnum.tsmEqual
            tStringsMatch = StrComp(str1, str2, compMethod) = 0
        Case tstrMatchEnum.tsmNotEqualTo
            tStringsMatch = StrComp(str1, str2, compMethod) <> 0
        Case tstrMatchEnum.tsmContains
            tStringsMatch = InStr(1, str1, str2, compMethod) > 0
        Case tstrMatchEnum.tsmStartsWithStr
            tStringsMatch = InStr(1, str1, str2, compMethod) = 1
        Case tstrMatchEnum.tsmEndWithStr
            If Len(str2) > Len(str1) Then
                tStringsMatch = False
            Else
                tStringsMatch = InStr(Len(str1) - Len(str2) + 1, str1, str2, compMethod) = Len(str1) - Len(str2) + 1
            End If
    End Select
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Used to set (or reset) Output Range properties
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function SetDefaults()
    lHAlign = xlHAlignGeneral
    lVAlign = xlVAlignCenter
    lInteriorColor = xlNone
    lMergeCells = True
    lWrapText = True
    lClearFormats = True
End Function

Private Sub Class_Initialize()
    SetDefaults
End Sub