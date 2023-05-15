Attribute VB_Name = "mdlGeneral"
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    '  Demo/Helper - Build Sheets
    '   ** Dependencies **
    '    - pbCommonUtil.bas
    '    - pbShapeBtn.bas
    '    - pbCommonEvents.cls
    '      NOTE:  pbCommonEvents.cls MUST be imported into
    '      any project. (Copy/Paste will now work)
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    '  author (c) Paul Brower https://github.com/lopperman/just-VBA
    '  module mdlGeneral.bas (One-Off Demo)
    '  license GNU General Public License v3.0
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Option Explicit
    Option Compare Text
    Option Base 1
    
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    '   Auto_Open is private, to prevent users from triggering
    '   the macro.  (it will still run when workbook is opened)
    '   This Function (which is not visible to the macro viewer)
    '   Can be used to manually call Auto_Open for
    '   development purposes
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function DevRunAutoOpen()
        Auto_Open
    End Function
    
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    '   FIRE WHEN WORKBOOK IS OPENED, REGARDLESS OF
    '   APPLICATION EVENTS STATE
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Sub Auto_Open()
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        pbShapeBtn.BuildShapeBtn wsDashboard _
            , "btnAddDateSheets" _
            , "SELECT WORKBOOK TO VERIFY DATE SHEETS" _
            , 2, -2, btnStyle:=bsAddEdit _
            , shpOnAction:="AddDateSheets" _
            , unitsWide:=3 _
            , unitsTall:=2
            
        If Not wsDashboard.Visible = xlSheetVisible Then
            wsDashboard.Visible = xlSheetVisible
        End If
        wsDashboard.Activate
        
        
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End Sub
    
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    '   Called when the button on the Dashboard is clicked
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function AddDateSheets()
    
        Dim wkbk As Workbook
        Dim msg As String
        Dim willChooseFile As Boolean
        msg = "Is the workbook you need to check currently open?"
        If MsgBox_FT(msg, vbYesNo + vbQuestion + vbDefaultButton2, "Find Workbook") = vbNo Then
            msg = "Are you able to select the file from your computer.  If not, answer NO and then open the workbook and start this process again."
            If MsgBox_FT(msg, vbYesNo + vbQuestion + vbDefaultButton2, "Find Workbook") = vbNo Then
                Exit Function
            Else
                willChooseFile = True
            End If
        End If
    
        If willChooseFile Then
            Dim openPath As String
            openPath = chooseFile("Select Workbook")
            If Len(openPath) > 0 Then
                If MsgBox_FT("Open file '" & FileNameFromFullPath(openPath) & "'?", vbYesNo + vbDefaultButton1 + vbQuestion, "Find Workbook") = vbNo Then
                    Exit Function
                Else
                    Set wkbk = Workbooks.Open(openPath)
                    DoEvents
                End If
            End If
        Else
            Dim tWB As Workbook
            For Each tWB In Application.Workbooks
                If Not StringsMatch(tWB.Name, ThisWorkbook.Name) Then
                    If MsgBox_FT("Is '" & tWB.Name & "' the workbook you wish to check?", vbYesNo + vbDefaultButton2 + vbQuestion, "Find Workbook") = vbYes Then
                        Set wkbk = tWB
                        Exit For
                    End If
                End If
            Next tWB
        End If
        
        If Not wkbk Is Nothing Then
            CheckDateSheets wkbk
        End If
    End Function
    
    Private Function CheckDateSheets(ByRef wkbk As Workbook)
        Dim processDt
        Dim updateRangeAddresses As New Collection
        'use this to set format of how to name sheets.
        'For this demo I'll use "MMM-DD-YYYY"
        Dim sheetNameFormat As String
        sheetNameFormat = "MMM-DD-YYYY"
        processDt = InputBox_FT("Enter any date of the month you wish to check", title:="Enter Date", default:=Date, inputType:=ftibString)
        If IsDate(processDt) Then
            processDt = DateSerial(DatePart("yyyy", processDt), DatePart("m", processDt), 1)
        End If
        
        'use this to put the cell references (e.g. "A10") that need to have the date updated on each sheet
        'For this demo, I'll use "A1", "A2", "A3"
        With updateRangeAddresses
            .Add "A1"
            .Add "A2"
            .Add "A3"
        End With
        
        Dim ws As Worksheet
        Dim includeWeekends As Boolean
        'change if you want to include weekends
        includeWeekends = False
        Dim workingDt, addDays As Long
        Dim tmpWorksheet As Worksheet, tmpSheetName As String, tmpExists As Boolean
        Dim tmpAddress
        workingDt = processDt
        
        Do While DatePart("m", workingDt) = DatePart("m", processDt)
            'don't change the weekstart = vbMonday -- that makes it easier to check for weekends
            If DatePart("w", workingDt, firstDayOfWeek:=vbMonday) <= 5 Or includeWeekends Then
                tmpSheetName = Format(workingDt, sheetNameFormat)
                If Not WorksheetExists(tmpSheetName, wbk:=wkbk) Then
                    Set tmpWorksheet = wkbk.Worksheets.Add(After:=wkbk.Worksheets(wkbk.Worksheets.Count))
                    tmpWorksheet.Name = tmpSheetName
                    For Each tmpAddress In updateRangeAddresses
                        tmpWorksheet.Range(CStr(tmpAddress)).Value = workingDt
                    Next tmpAddress
                End If
            End If
            
            addDays = addDays + 1
            workingDt = DateAdd("d", addDays, processDt)
        Loop
    End Function
    
    
