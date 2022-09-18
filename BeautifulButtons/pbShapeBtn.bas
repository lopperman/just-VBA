Attribute VB_Name = "pbShapeBtn"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' pbShapeBtn v1.0.1
' (c) Paul Brower - https://github.com/lopperman/just-VBA
'
' Alternative to dull boring built-in buttons
'
' @module pbButtonShape
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1

' ~~~ ~~~ ~~~ ~~~ UPDPATE TO V1.0.1  ~~~ ~~~ ~~~ ~~~
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


'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   UNCOMMENT NEXT 3 LINES BEFORE USING
'   Private Const BTN_FIRST_LEFT As Single = 275
'   Private Const BTN_NAV_FIRST_LEFT As Single = 3
'   Private Const BTN_FIRST_TOP As Single = 3

'   COMMENT (OR DELETE) NEXT 3 LINES BEFORE USING
    Private Const BTN_FIRST_LEFT As Single = 275 + 750
    Private Const BTN_NAV_FIRST_LEFT As Single = 3 + 750
    Private Const BTN_FIRST_TOP As Single = 3 + 258
'   ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~



Public Enum bsStyle
    bsNavigation = 1
    bsdefault = 1
    bsutility
    bsReport
    bsFilter
    bsAddEdit
    bsdelete
    bsHelp
    bsCustom
End Enum

Public Function BuildPrimaryNavBtn(ws As Worksheet, shpName As String, shpCaption As String, rowPos As Long, colPos As Long, Optional btnStyle As bsStyle = bsStyle.bsNavigation)

    Dim navLeft As Variant, navtop As Variant
    navLeft = BTN_NAV_FIRST_LEFT
    navLeft = navLeft + ((colPos - 1) * (NAV_BTN_UNIT_WIDTH + BTN_PADDING))
    navtop = BTN_FIRST_TOP + ((rowPos - 1) * BTN_PADDING) + ((rowPos - 1) * BTN_UNIT_HEIGHT)
    
    BuildShapeBtn ws, shpName, shpCaption, 0, 0, btnStyle, "ButtonAction", navLeft, navtop, 2, 1, forceWidth:=(NAV_BTN_UNIT_WIDTH * 2)

End Function

Public Function BuildShapeBtn(ws As Worksheet, shpName As String, shpCaption As String, rowPos As Long, colPos As Long, _
    Optional btnStyle As bsStyle = bsStyle.bsdefault, Optional shpOnAction As String = "ButtonAction", _
    Optional forceLeft As Variant, Optional forceTop As Variant, _
    Optional unitsWide As Long = 2, Optional unitsTall As Long = 1, _
    Optional fontClr As Variant, Optional fillClr As Variant, Optional lineClr As Variant, Optional lineWt As Variant, _
    Optional fontSize As Long = 14, Optional fontBold As Boolean = True, Optional fontUnderLine As Boolean = False, Optional shpPlacement As XlPlacement = XlPlacement.xlFreeFloating, Optional forceWidth As Variant)
    

    Dim adjLeft As Single, adjTop As Single, adjWidth As Single, adjHeight As Single
    If Not IsMissing(forceLeft) And Not IsMissing(forceTop) Then
        adjLeft = CSng(forceLeft)
        adjTop = CSng(forceTop)
    Else
        Dim lPadCnt As Long, tPadCnt As Long
        lPadCnt = colPos - 1
        tPadCnt = rowPos - 1
        adjLeft = BTN_FIRST_LEFT + (lPadCnt * BTN_UNIT_WIDTH) + (lPadCnt * BTN_PADDING)
        adjTop = BTN_FIRST_TOP + (tPadCnt * BTN_UNIT_HEIGHT) + (tPadCnt * BTN_PADDING)
    End If
    adjWidth = (BTN_UNIT_WIDTH * unitsWide) + ((unitsWide - 1) * BTN_PADDING)
    adjHeight = (BTN_UNIT_HEIGHT * unitsTall) + ((unitsTall - 1) * BTN_PADDING)
    
    
    If Not IsMissing(forceWidth) Then adjWidth = forceWidth
    
    Dim shp As Shape
    Set shp = FindShapeButton(ws, shpName)
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, adjLeft, adjTop, adjWidth, adjHeight)
    Else
        If Not shp.Left = adjLeft Then shp.Left = adjLeft
        If Not shp.Top = adjTop Then shp.Top = adjTop
        If Not shp.Width = adjWidth Then shp.Width = adjWidth
        If Not shp.Height = adjHeight Then shp.Height = adjHeight
    End If
    
    With shp
        .Placement = shpPlacement
        If Not .Name = shpName Then .Name = shpName
        .Locked = True
        .OnAction = shpOnAction
        .ZOrder msoBringToFront
        .TextFrame2.TextRange.Characters.Text = shpCaption
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.TextRange.Font.Size = fontSize
        .TextFrame2.TextRange.Font.Bold = IIf(fontBold, msoTrue, msoFalse)
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
    End If
    
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
            ChangeBtnFill shpBtn, 7884319
            ChangeBtnLine shpBtn, 4291544, lineWidth:=1
        Case bsStyle.bsHelp
            ChangeBtnFont shpBtn, 7884319, fntSize:=14, fntBold:=True
            ChangeBtnFill shpBtn, 16777215
            ChangeBtnLine shpBtn, 7884319, lineWidth:=1
            'Add2ColorGradient shpBtn, 13431551, 13431539
            
        Case bsStyle.bsutility
            ChangeBtnFont shpBtn, 7884319, fntSize:=12
            ChangeBtnFill shpBtn, 14998742
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
            .BackColor.RGB = backClr
'            .backColor.Brightness = 0.15
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
Public Function ChangeBtnLine(ByRef shpBtn As Shape, ByVal rgbClr As Long, Optional lineWidth As Variant)
    Dim cf As ColorFormat
    Set cf = shpBtn.Line.ForeColor
    cf.RGB = rgbClr
    If Not IsMissing(lineWidth) Then shpBtn.Line.Weight = lineWidth
    Set cf = Nothing
End Function


Public Function FindShapeButton(ws As Worksheet, shpName As String) As Shape
    If ws.Shapes.Count > 0 Then
        Dim shp As Shape
        For Each shp In ws.Shapes
            If shp.Type = msoAutoShape Then
                If StringsMatch(shp.Name, shpName) Then
                    Set FindShapeButton = shp
                    Exit For
                End If
            End If
        Next shp
    End If
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

Public Function AddPrimaryNavigation(ws As Worksheet)
        
        BuildPrimaryNavBtn ws, "btnNavHome", "DASHBOARD", 1, 1
        BuildPrimaryNavBtn ws, "btnNavTeam", "TEAM", 1, 3
        BuildPrimaryNavBtn ws, "btnNavForecast", "FORECAST", 2, 1
        BuildPrimaryNavBtn ws, "btnNavCostHours", "COST-HOURS", 2, 3
        BuildPrimaryNavBtn ws, "btnNavConfig", "CONFIG", 3, 1
        BuildPrimaryNavBtn ws, "btnNavSupport", "SUPPORT", 3, 3, btnStyle:=bsHelp

    Dim sheetGraphic As Shape
    Set sheetGraphic = FindSheetGraphic(ws)
    If Not sheetGraphic Is Nothing Then
        Dim newLeft As Variant
        newLeft = BTN_NAV_FIRST_LEFT + (BTN_PADDING * 2) + (106 * 2) + 3
        If Not sheetGraphic.Top = BTN_FIRST_TOP Then sheetGraphic.Top = BTN_FIRST_TOP
        If Not sheetGraphic.Left = newLeft Then sheetGraphic.Left = newLeft
        If Not sheetGraphic.Width = 50 Then sheetGraphic.Width = 50
        If Not sheetGraphic.Height = 50 Then sheetGraphic.Height = 50
        Set sheetGraphic = Nothing
    End If
    
End Function

Public Function FindSheetGraphic(fndInWS As Worksheet) As Shape
    Dim shp As Shape
    For Each shp In fndInWS.Shapes
        If StringsMatch(fndInWS.CodeName & "_graphic", shp.Name) Then
            Set FindSheetGraphic = shp
            Exit For
        End If
    Next shp
End Function




