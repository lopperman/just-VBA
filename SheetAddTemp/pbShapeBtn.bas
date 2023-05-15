Attribute VB_Name = "pbShapeBtn"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' pbShapeBtn v1.0.1
' (c) Paul Brower - https://github.com/lopperman/just-VBA
'
' Alternative to dull boring built-in buttons
'
' @module pbButtonShape
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1
Option Private Module

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   UPDPATE TO V1.0.1
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'
'   I did a bit of research on color schemes, so this version
'   has colors that go better together for the default
'   button styles
'   Also added the option to have a main set of
'   'default navigation buttons' -- these would be the
'   same on all your pages, then the second set
'   would be additional buttons for different pages
'   if you don't want to use the 'Nav' (navigation0)
'   button area, just change BTN_NAV_FIRST_LEFT below
'   to be '3' instead of '275
'
'   Added 'Add2ColorGradient' Function -- for quick
'   and easy 2-color gradients - just pass in the
'   reference to your Shape object, and the 2 colors you want
'   to use for gradients
'
'   If you have a small graphic that you'd like to be placed
'   between the primary navigation buttons, and the rest
'   just add the image to your worksheet, and name it
'   [worksheet code name]_graphic
'   If your sheet is passed in to 'AddPrimaryNavigation'
'   function, the image will be found, resized, and used
'   as a 'spacer' between the navigation buttons and
'   the rest of your screen buttons


Private Const BTN_UNIT_WIDTH As Single = 68
Private Const NAV_BTN_UNIT_WIDTH As Single = 53
Private Const BTN_UNIT_HEIGHT As Single = 25
Private Const BTN_DEFAULT_FONT_SIZE As Long = 14
Private Const BTN_PADDING As Single = 2
Private Const BTN_LINE_WEIGHT As Single = 1
Private Const BTN_DEFAULT_FONT_COLOR As Long = 11232280
Private Const BTN_DEFAULT_FILL_COLOR As Long = 13431551
Private Const BTN_DEFAULT_LINE_COLOR As Long = 6740479
'   ~~~ ~~~
Private Const BTN_FIRST_LEFT As Single = 275
Private Const BTN_NAV_FIRST_LEFT As Single = 3
Private Const BTN_FIRST_TOP As Single = 3


Public Enum bsStyle
    bsNavigation = 1
    bsdefault = 1
    bsutility
    bsReport
    bsFilter
    bsAddEdit
    bsdelete
    bshelp
    bsCustom
End Enum

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

Public Function SampleBeep()
''   IF YOU HAVE A LEGEND ANYWHERE FOR BUTTON STYLES, YOU CAN HAVE THE 'Action' set to 'SampleBeep' in case the User clicks the sample buttons
    Beep
End Function

Public Function BuildPrimaryNavBtn(ws As Worksheet, shpName As String, shpCaption As String, rowPos As Long, colPos As Long, Optional btnStyle As bsStyle = bsStyle.bsNavigation) As Shape

    Dim navLeft As Variant, navtop As Variant
    navLeft = BTN_NAV_FIRST_LEFT
    navLeft = navLeft + ((colPos - 1) * (NAV_BTN_UNIT_WIDTH + BTN_PADDING))
    navtop = BTN_FIRST_TOP + ((rowPos - 1) * BTN_PADDING) + ((rowPos - 1) * BTN_UNIT_HEIGHT)
    
    Set BuildPrimaryNavBtn = BuildShapeBtn(ws, shpName, shpCaption, 0, 0, btnStyle, "ButtonAction", navLeft, navtop, 2, 1, forceWidth:=(NAV_BTN_UNIT_WIDTH * 2))

    Select Case shpName
        Case "btnNavHome"
            ChangeBtnFill BuildPrimaryNavBtn, 16750080
        Case "btnNavConfig"
            ChangeBtnFont BuildPrimaryNavBtn, rgbClr:=1137094
        Case "btnNavTeam"
            ChangeBtnFill BuildPrimaryNavBtn, rgbClr:=11159552
        Case "btnNavCostHours"
            ChangeBtnFill BuildPrimaryNavBtn, rgbClr:=14179072
        Case "btnNavForecast"
            ChangeBtnFill BuildPrimaryNavBtn, rgbClr:=16739072
        Case ""
    End Select

End Function


Public Function BuildShapeBtn(ws As Worksheet _
    , shpName As String, shpCaption As String _
    , rowPos As Long, colPos As Long _
    , Optional btnStyle As bsStyle = bsStyle.bsdefault _
    , Optional shpOnAction As String = "ButtonAction" _
    , Optional forceLeft As Variant _
    , Optional forceTop As Variant _
    , Optional unitsWide As Long = 2 _
    , Optional unitsTall As Long = 1 _
    , Optional fontClr As Variant _
    , Optional fillClr As Variant _
    , Optional lineClr As Variant _
    , Optional lineWt As Variant _
    , Optional fontSize As Long = 14 _
    , Optional fontBold As Boolean = True _
    , Optional fontUnderLine As Boolean = False _
    , Optional shpPlacement As XlPlacement = XlPlacement.xlFreeFloating _
    , Optional forceWidth As Variant _
    , Optional forceHeight As Variant) As Shape
    

    Dim adjLeft As Single, adjTop As Single, adjWidth As Single, adjHeight As Single
    If Not IsMissing(forceLeft) Then adjLeft = CSng(forceLeft)
    If Not IsMissing(forceTop) Then adjTop = CSng(forceTop)
    If Not IsMissing(forceHeight) Then adjHeight = CSng(forceHeight)
    If Not IsMissing(forceWidth) Then adjWidth = CSng(forceWidth)
    
    Dim lPadCnt As Long, tPadCnt As Long
    lPadCnt = colPos - 1
    tPadCnt = rowPos - 1
    If adjLeft = 0 Then adjLeft = BTN_FIRST_LEFT + (lPadCnt * BTN_UNIT_WIDTH) + (lPadCnt * BTN_PADDING)
    If adjTop = 0 Then adjTop = BTN_FIRST_TOP + (tPadCnt * BTN_UNIT_HEIGHT) + (tPadCnt * BTN_PADDING)
    If adjWidth = 0 Then adjWidth = (BTN_UNIT_WIDTH * unitsWide) + ((unitsWide - 1) * BTN_PADDING)
    If adjHeight = 0 Then adjHeight = (BTN_UNIT_HEIGHT * unitsTall) + ((unitsTall - 1) * BTN_PADDING)
    
    Dim shp As Shape
    Set shp = FindShapeButton(ws, shpName)
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, adjLeft, adjTop, adjWidth, adjHeight)
    Else
        If Not shp.Left = adjLeft Then shp.Left = adjLeft
        If Not shp.Top = adjTop Then shp.Top = adjTop
        If Not shp.width = adjWidth Then shp.width = adjWidth
        If Not shp.height = adjHeight Then shp.height = adjHeight
    End If
    
    With shp
        If Not .Placement = shpPlacement Then .Placement = shpPlacement
        If Not .Name = shpName Then .Name = shpName
        If .Locked = False Then .Locked = True
        If StringsMatch(.onAction, shpOnAction) = False Then .onAction = shpOnAction
        .ZOrder msoBringToFront
        If Not StringsMatch(.TextFrame2.TextRange.Characters.Text, shpCaption) Then .TextFrame2.TextRange.Characters.Text = shpCaption
        If Not .TextFrame2.VerticalAnchor = msoAnchorMiddle Then .TextFrame2.VerticalAnchor = msoAnchorMiddle
        If Not .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter Then .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        If Not .TextFrame2.TextRange.Font.Size = fontSize Then .TextFrame2.TextRange.Font.Size = fontSize
        If Not .TextFrame2.TextRange.Font.Bold = IIf(fontBold, msoTrue, msoFalse) Then .TextFrame2.TextRange.Font.Bold = IIf(fontBold, msoTrue, msoFalse)
    End With
    With shp.TextFrame2.TextRange.Characters(1, Len(shp.TextFrame2.TextRange.Characters.Text)). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignCenter
    End With
    With shp.TextFrame2.TextRange.Characters(1, Len(shp.TextFrame2.TextRange.Characters.Text)).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Name = "+mn-lt"
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent4
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Weight = 1
    End With
    
    If btnStyle = bsCustom Then
        ChangeBtnFont shp, IIf(IsMissing(fontClr), BTN_DEFAULT_FONT_COLOR, CLng(fontClr))
        ChangeBtnFill shp, IIf(IsMissing(fillClr), BTN_DEFAULT_FONT_COLOR, CLng(fillClr))
        ChangeBtnLine shp, IIf(IsMissing(lineClr), BTN_DEFAULT_LINE_COLOR, CLng(lineClr)), IIf(IsMissing(lineWt), BTN_LINE_WEIGHT, CSng(lineWt))
    Else
        FormatDefinedStyle shp, btnStyle
        If Not IsMissing(fontClr) Then ChangeBtnFont shp, rgbClr:=fontClr
        If Not IsMissing(fillClr) Then ChangeBtnFill shp, rgbClr:=fillClr
        If Not IsMissing(lineClr) Then ChangeBtnLine shp, rgbClr:=lineClr
        If Not IsMissing(lineWt) Then ChangeBtnLine shp, lineWidth:=lineWt
        If (fontSize <> 14 Or fontBold = False Or fontUnderLine = True) Then ChangeBtnFont shp, fntSize:=fontSize, fntBold:=fontBold, fntUnderline:=fontUnderLine
        If Not IsMissing(forceWidth) Then shp.width = forceWidth
    End If
    
    
    Set BuildShapeBtn = shp
    Set shp = Nothing
    
End Function



Private Function FormatDefinedStyle(ByRef shpBtn As Shape, ByVal bStyle As bsStyle)

    Select Case bStyle
        Case bsStyle.bsFilter
            ChangeBtnFont shpBtn, 7884319, fntSize:=14
            ChangeBtnFill shpBtn, 14348258
            ChangeBtnLine shpBtn, 7884319, lineWidth:=1
        Case bsStyle.bsAddEdit
            ChangeBtnFont shpBtn, 16724484, fntSize:=14
            ChangeBtnFill shpBtn, 16247773
            ChangeBtnLine shpBtn, 16724484, lineWidth:=1
        Case bsStyle.bsdelete
            ChangeBtnFont shpBtn, 255, fntSize:=14
            ChangeBtnFill shpBtn, 14083324
            ChangeBtnLine shpBtn, 255, lineWidth:=1
        Case bsStyle.bsNavigation
            ChangeBtnFont shpBtn, 16777215, fntSize:=14, fntBold:=True
            'ChangeBtnFill shpBtn, 11159552
            ChangeBtnFill shpBtn, 14179072
            ChangeBtnLine shpBtn, 4291544, lineWidth:=1
        Case bsStyle.bshelp
            ChangeBtnFont shpBtn, 7884319, fntSize:=14, fntBold:=True
            ChangeBtnFill shpBtn, 16777215
            ChangeBtnLine shpBtn, 7884319, lineWidth:=1
            'Add2ColorGradient shpBtn, 13431551, 13431539
            
        Case bsStyle.bsutility
            ChangeBtnFont shpBtn, 7884319, fntSize:=12
            'ChangeBtnFill shpBtn, 14998742
            ChangeBtnFill shpBtn, 15461355
            ChangeBtnLine shpBtn, 37887, lineWidth:=1
        Case bsStyle.bsReport
            ChangeBtnFont shpBtn, 7884319, fntSize:=14, fntBold:=True
            ChangeBtnFill shpBtn, 7884319
            ChangeBtnLine shpBtn, 37887, lineWidth:=1
            Add2ColorGradient shpBtn, 16120062, 12703994
        
    End Select

End Function

Public Function Add2ColorGradient(ByRef shp As Shape, foreClr As Long, backClr As Long, Optional grdStyle As MsoGradientColorType = msoGradientHorizontal)
        With shp.Fill
            .Solid 'reset previous gradient
            .TwoColorGradient grdStyle, 2
            .ForeColor.RGB = foreClr '(shows on botton)
            .backColor.RGB = backClr
            .RotateWithObject = msoTrue
        End With
End Function

'Private Function AddReportGradient(ByRef shp As Shape)
'
'        With shp.Fill
'            .Solid
'            .TwoColorGradient msoGradientHorizontal, 2
'            .ForeColor.RGB = 16247773 '(shows on botton)
'            .backColor.RGB = 15123099
'            .RotateWithObject = msoTrue
'        End With
'
'End Function

Public Function ChangeBtnFont(ByRef shpBtn As Shape, Optional ByVal rgbClr As Variant, _
    Optional fntSize As Variant, Optional fntBold As Boolean = True, Optional fntUnderline As Boolean = False)
    
    Dim cf As ColorFormat
    If Not IsMissing(rgbClr) Then
        Set cf = shpBtn.TextFrame2.TextRange.Font.Fill.ForeColor
        cf.RGB = rgbClr
    End If
    shpBtn.TextFrame2.TextRange.Font.Bold = fntBold
    If Not IsMissing(fntSize) Then shpBtn.TextFrame2.TextRange.Font.Size = fntSize
    If fntUnderline Then
        shpBtn.TextFrame2.TextRange.Font.UnderlineStyle = msoUnderlineSingleLine
        shpBtn.TextFrame2.TextRange.Font.UnderlineColor = cf.RGB
    Else
        shpBtn.TextFrame2.TextRange.Font.UnderlineStyle = msoNoUnderline
    End If
    
    Set cf = Nothing
    
    
    
End Function
Public Function ChangeBtnFill(ByRef shpBtn As Shape, ByVal rgbClr As Long)
    Dim cf As ColorFormat
    shpBtn.Fill.Solid
    Set cf = shpBtn.Fill.ForeColor
    cf.RGB = rgbClr
    Set cf = Nothing
End Function
Public Function ChangeBtnLine(ByRef shpBtn As Shape, Optional ByVal rgbClr As Variant, Optional lineWidth As Variant)
    If Not IsMissing(rgbClr) Then
        Dim cf As ColorFormat
        Set cf = shpBtn.Line.ForeColor
        cf.RGB = rgbClr
    End If
    If Not IsMissing(lineWidth) Then shpBtn.Line.Weight = lineWidth
    Set cf = Nothing
End Function

Public Function FindChart(ws As Worksheet, chtName As String) As Chart
    Set FindChart = FindShape(ws, chtName, msoChart)
End Function
    Public Function FindShape(ws As Worksheet, shpName As String, shpType As MsoShapeType, Optional autoShpType As MsoAutoShapeType = msoShapeRoundedRectangle) As Shape
        If shpType = msoAutoShape Then
            Set FindShape = FindShapeAsObj(ws, shpName, shpType, autoShpType)
        Else
            Set FindShape = FindShapeAsObj(ws, shpName, shpType)
        End If
    End Function
    'FIND AND RETURN ANY msoShapeType by Worksheet, ByName
    Public Function FindShapeAsObj(ws As Worksheet, shpName As String, shpType As MsoShapeType, Optional autoShpType As Variant) As Object
        If ws.Shapes.Count > 0 Then
            Dim shp As Shape
            For Each shp In ws.Shapes
                If shp.Type = shpType Then
                    If Not IsMissing(autoShpType) Then
                        If shp.AutoShapeType = autoShpType Then
                            If StringsMatch(shp.Name, shpName) Then
                                Set FindShapeAsObj = shp
                                Exit For
                            End If
                        End If
                    Else
                        If StringsMatch(shp.Name, shpName) Then
                            Set FindShapeAsObj = shp
                            Exit For
                        End If
                    End If
                End If
            Next shp
        End If
    End Function

Public Function FindShapeButton(ws As Worksheet, btnName As String) As Shape
    Set FindShapeButton = FindShapeAsObj(ws, btnName, msoAutoShape, msoShapeRoundedRectangle)
End Function

Public Function DeleteButtonShp(ws As Worksheet, shpName As String, Optional deleteAllMsAutoShapes As Boolean = False)
        Dim shp As Shape
        For Each shp In ws.Shapes
            If shp.Type = msoAutoShape Then
                If deleteAllMsAutoShapes Then
                    shp.Delete
                ElseIf StringsMatch(shp.Name, shpName) Then
                    shp.Delete
                    Exit For
                End If
            End If
        Next shp
End Function

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'   LOOK FOR [shapeName] on all worksheets, delete when found
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Public Function DeleteFoundShapes(shapeName As String, Optional countOnly As Boolean = False)
    Dim ws As Worksheet, shp As Shape, shpIndex As Long, foundCount As Long
    For Each ws In ThisWorkbook.Worksheets
        shpIndex = 1
        For Each shp In ws.Shapes
            If StringsMatch(shp.Name, shapeName) Then
                foundCount = foundCount + 1
                If Not countOnly Then
                    ws.Shapes(shpIndex).Delete
                End If
                Exit For
            End If
            shpIndex = shpIndex + 1
        Next shp
    Next ws
    DeleteFoundShapes = foundCount
End Function

Public Function AddPrimaryNavigation(ws As Worksheet)
    Select Case ws.CodeName
    Case "wsDashboard"
        BuildShapeBtn ws, "btnExit", "E X I T", 1, 1, bsCustom, "QuitApp", BTN_NAV_FIRST_LEFT, BTN_FIRST_TOP, 2, 1, 16724484, 14998742, 14395790, 1, 18, forceWidth:=106
        BuildPrimaryNavBtn ws, "btnNavConfig", "CONFIG", 3, 1, btnStyle:=bshelp
        BuildPrimaryNavBtn ws, "btnNavSupport", "SUPPORT", 2, 1, btnStyle:=bshelp
        BuildPrimaryNavBtn ws, "btnNavTeam", "TEAM", 1, 3
        BuildPrimaryNavBtn ws, "btnNavCostHours", "COST-HOURS", 2, 3
        BuildPrimaryNavBtn ws, "btnNavForecast", "FORECAST", 3, 3
    Case Else
        BuildPrimaryNavBtn ws, "btnNavHome", "DASHBOARD", 1, 1
        BuildPrimaryNavBtn ws, "btnNavConfig", "CONFIG", 3, 1, btnStyle:=bshelp
        BuildPrimaryNavBtn ws, "btnNavSupport", "SUPPORT", 2, 1, btnStyle:=bshelp
        BuildPrimaryNavBtn ws, "btnNavTeam", "TEAM", 1, 3
        BuildPrimaryNavBtn ws, "btnNavCostHours", "COST-HOURS", 2, 3
        BuildPrimaryNavBtn ws, "btnNavForecast", "FORECAST", 3, 3
    End Select
    
    Dim sheetGraphic As Shape
    Set sheetGraphic = FindSheetGraphic(ws)
    If Not sheetGraphic Is Nothing Then
        Dim newLeft As Variant, newTop As Variant
        newLeft = BTN_NAV_FIRST_LEFT + (BTN_PADDING * 2) + (106 * 2) + 8
        newTop = 25
        If Not sheetGraphic.Top = newTop Then sheetGraphic.Top = newTop
        If Not sheetGraphic.Left = newLeft Then sheetGraphic.Left = newLeft
        If Not sheetGraphic.width = 40 Then sheetGraphic.width = 40
        If Not sheetGraphic.height = 40 Then sheetGraphic.height = 40
        Set sheetGraphic = Nothing
        AdjustEye ws
    End If
    
End Function

Public Function AdjustEye(ws As Worksheet)
    Dim shtGraphic As Shape
    Set shtGraphic = FindSheetGraphic(ws)
    If Not shtGraphic Is Nothing Then
        If Not FindShape(ws, "changeFont", msoGraphic) Is Nothing Then
            With FindShape(ws, "changeFont", msoGraphic)
                .width = 28
                .height = 28
                .Top = 0
                .Left = shtGraphic.Left + (shtGraphic.width / 2) - (.width / 2)
                .Placement = xlFreeFloating
            End With
        End If
    End If
End Function

Public Function FindSheetGraphic(fndInWS As Worksheet) As Shape
    Dim shp As Shape
    For Each shp In fndInWS.Shapes
        If StringsMatch(Concat(fndInWS.CodeName, "_", "graphic"), shp.Name) Then
            Set FindSheetGraphic = shp
            Exit For
        End If
    Next shp
End Function






' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Private Version of pbCommon.EnumCompare
'   Included in some 'pb[XXXXX]' modules in order to remove
'   dependency on pbCommon
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function EnumCompare(theEnum As Variant, enumMember As Variant, Optional ByVal iType As ecComparisonType = ecComparisonType.ecOR) As Boolean
    Dim c As Long
    c = theEnum And enumMember
    EnumCompare = IIf(iType = ecOR, c <> 0, c = enumMember)
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Private Version of pbCommon.Concat
'   Included in some 'pb[XXXXX]' modules in order to remove
'   dependency on pbCommon
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function Concat(ParamArray Items() As Variant) As String
    Concat = Join(Items, "")
End Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Private Version of pbCommon.ConcatWithDelim
'   Included in some 'pb[XXXXX]' modules in order to remove
'   dependency on pbCommon
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function ConcatWithDelim(ByVal delimeter As String, ParamArray Items() As Variant) As String
    ConcatWithDelim = Join(Items, delimeter)
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Private Version of pbCommon.StringsMatch
'   Included in some 'pb[XXXXX]' modules in order to remove
'   dependency on pbCommon
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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



