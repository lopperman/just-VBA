Attribute VB_Name = "autoPB"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  ** NOT INTENDED TO BE REUSED OUTSIDE ORIGINAL WORKBOOK **
''
'  Utililities And Start Options For Current Workbook
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  author (c) Paul Brower https://github.com/lopperman
'  module autoPB.bas
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  CONSTANTS, TYPES, ENUMS EXCLUSIVE TO VBEUTIL WORKBOOK
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Const COLOR_BLUEBERRY As Long = 16724484

Public Function UDFTest()
    Debug.Print Application.ThisCell.Worksheet.Name & " - " & Application.ThisCell.Address
    UDFTest = Application.ThisCell.Address
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  AUTO_OPEN
''  AWAYS RUNS ON WORKBOOK OPEN, EVEN WITH DISABLED EVENTS
''  (Private to prevent launching from Macro Browser)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Sub Auto_Open()
    On Error Resume Next
    AppMode = appStatusStarting
    ''Initialize Settings
    ThisWorkbook.EnableConnections
    Application.EnableEvents = False
    If stg.ValidConfig Then
        DoEvents
        Logger.LogFORCED "Auto_Open - Initializing Settings "
        CheckDevKey
    Else
        Beep
        Application.EnableEvents = True
        MsgBox "Error on Auto_Open - Unable to initialize Settings", vbOKOnly + vbCritical, "OOPS"
        AppMode = appStatusRunning
        
        Exit Sub
    End If

    If IsDev Then
        Logger.LogFORCED " *** Current User ISDEV *** "
    End If
    Logger.LogFORCED "Auto_Open: " & ThisWorkbook.FullName
    Logger.LogFORCED "Auto_Open - CurrentUser: " & ENV_LogName
    DoEvents
    CheckButtons
    wsDashboard.visible = xlSheetVisible
    wsDashboard.Activate
    wsDashboard.OnFormat
    
    pbCommonUtil.pbCheckDefaultSettings
    
    Application.EnableEvents = True
    AppMode = appStatusRunning
    
    If Err.number <> 0 Then
        Debug.Print Err.number, Err.Description
        If IsDev Then
            Beep
            Stop
        End If
    End If
End Sub

    Public Function CheckButtons(Optional wksht As Worksheet, Optional cfgValid)
        Dim tws As Worksheet, cfgIsValid As Boolean
        If wksht Is Nothing Then
            cfgIsValid = stg.ValidConfig
            For Each tws In ThisWorkbook.Worksheets
                CheckButtons tws, cfgValid:=cfgIsValid
            Next tws
            Exit Function
        End If
        If IsMissing(cfgValid) Then cfgValid = stg.ValidConfig
        If cfgValid = True Then
            CheckPBShapeButtons wksht
        End If
        Select Case wksht.CodeName
            Case wsDashboard.CodeName
                pbShapeBtn.BuildShapeBtn wksht _
                    , "btnKillGridlines", "ALL UR GRIDLINE BELONG TO US", 1, 5 _
                    , btnStyle:=bsutility, shpOnAction:="AllUrGridlineBelongToUs" _
                    , unitsWide:=2, unitsTall:=2, fontSize:=10
                pbShapeBtn.BuildShapeBtn wksht _
                    , "btnCodeUtility", "CODE UTILITY", 1, 1 _
                    , btnStyle:=bsNavigation, shpOnAction:="NavCodeUtil" _
                    , unitsWide:=2, unitsTall:=1
            Case wsCodeUtil.CodeName
                pbShapeBtn.BuildShapeBtn wksht _
                    , "btnExportCode", "EXPORT CODE", 1, 1 _
                    , btnStyle:=bsutility, shpOnAction:="ExportSupported" _
                    , unitsWide:=2, unitsTall:=1, fontSize:=12
            Case wsDevTestSheet.CodeName
            
            
        End Select
    
    End Function

''  (DEV) Enable Calling Auto_Open from Immediate Window
Public Function DevAuto()
    If IsDev Then
        Auto_Open
    End If
End Function

Public Function CheckPBShapeButtons(wksht As Worksheet)
    Dim tmpBtn As Shape, validSettings As Boolean, stgSheetCodeName As String
    If Not stg.ValidConfig Then
        Exit Function
    End If
    stgSheetCodeName = stg.pbSettingsSheet.CodeName
    Select Case wksht.CodeName
        ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
        ''  CONFIGURE/CHECK BUTTONS FOR wsDashboard
        ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
        Case wsDashboard.CodeName
            '' CREATE OR VERIFY 'btnExit'
            Set tmpBtn = pbShapeBtn.BuildPrimaryNavBtn(wksht _
                , "btnExit", "EXIT", 1, 1 _
                , btnStyle:=bshelp, _
                fontColor:=COLOR_BLUEBERRY)
            If Not StringsMatch(tmpBtn.OnAction, "pbSettingButtonAction") Then tmpBtn.OnAction = "pbSettingButtonAction"
            
            
            '' CREATE OR VERIFY 'btnToggleAutoHidePBSTG'
            pbShapeBtn.BuildShapeBtn wksht _
                , "btnToggleAutoHidePBSTG", "TOGGLE SETTINGS AUTOHIDE", 2, -1 _
                , btnStyle:=bsutility, shpOnAction:="ToggleSettingsAutoHide" _
                , unitsWide:=2, unitsTall:=1, fontSize:=10
            pbShapeBtn.BuildShapeBtn wksht _
                , "btnNAVpbSettings", "SETTINGS", 1, -1 _
                , btnStyle:=bsNavigation, shpOnAction:="NAVSettings" _
                , unitsWide:=2, unitsTall:=1
                
            ''  AllUrGridlineBelongToUs
            
        
        ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
        ''  CONFIGURE/CHECK BUTTONS FOR pbSettings Worksheet
        ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
        Case stgSheetCodeName
    End Select
    
    If Not StringsMatch(wksht.CodeName, wsDashboard.CodeName) Then
            Dim navDashboardCmd As String, btnNameNavHome As String
            navDashboardCmd = "NavDashboard"
            btnNameNavHome = "btnNavHome"
            Set tmpBtn = pbShapeBtn.BuildPrimaryNavBtn(wksht _
                , btnNameNavHome, "DASHBOARD", 1, 1 _
                , btnStyle:=bsNavigation)
            If Not StringsMatch(tmpBtn.OnAction, "pbSettingButtonAction") Then tmpBtn.OnAction = "pbSettingButtonAction"
            If Not StringsMatch(stg.stdButtonAction(wksht, btnNameNavHome), navDashboardCmd) Then
                stg.stdButtonAction(wksht, btnNameNavHome) = navDashboardCmd
            End If
    End If
    
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  AUTO HIDE SHEETS
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function CheckSheetsVisible()
        Dim doNotHide As Variant
        doNotHide = stg.stdArrayDoNotHideSheets
        Dim tws As Worksheet, dnhSheet
        Dim keepVisible As Boolean
        For Each tws In ThisWorkbook.Worksheets
            keepVisible = False
            If tws Is ThisWorkbook.activeSheet Then
                keepVisible = True
            ElseIf tws.visible = xlSheetVisible Then
                For Each dnhSheet In doNotHide
                    If tws.visible <> xlSheetVisible And StringsMatch(tws.CodeName, dnhSheet) Then
                        keepVisible = True
                        Exit For
                    End If
                Next dnhSheet
                If Not keepVisible Then tws.visible = xlSheetVeryHidden
            End If
        Next tws
    End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  pbShapeBtn OnAction Handlers
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~ '
    ''  FOR ANY pbShapeBtn BUTTON, that has the 'ONACTION' Property set to:
    ''      'pbSettingButtonAction', the btnName and worksheet CodeName will be
    ''      passed to pbSettings.
    ''  pbSetting will look to determine if the [worksheet.codename] +[button name]
    ''      combination have a public 'Function' or 'Sub' Name set to be called.
    ''      The method to be called must exist in a standard module (not class module),
    ''      and must be able to be called without arguments.
    ''  Example of where this might be useful would be something like a feature flag
    ''      e.g. You have a Worksheet called "Invoices" and another called "Invoices 2.0"
    ''      You can develop the 'Invoices 2.0' sheet, and even deploy unfinished code
    ''      if you needed to fix a bug in another area. No users would be able to
    ''      interact with the "Invoices 2.0" worksheet until it was ready.  At that point,
    ''      you could modify the local pbSetting value by synchronizing the setting value
    ''      with a SharePoint list.  Changing the value of a SharePoint setting ("OnClickBtnInvoices")
    ''      from something like "OpenInvoices" to "OpenNewInvoices"
    ''      "OpenInvoices" and "OpenNewInvoices" would be public function in a standard
    ''      Module, and once the setting was changed, any user clicking the button would
    ''      be taken to the 'Invoices 2.0' area.
    ''
    ''      AN EXAMPLE OF THIS 'SETTING DRIVEN BUTTON ACTION' CAN BE VIEWED BY
    ''      DOWNLOADING THE LATEST 'VBEUTIL [version].xlsm' from:
    ''          https://github.com/lopperman/just-VBA/tree/main/VBE-Tools
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~ '
    Public Function pbSettingButtonAction()
        On Error Resume Next
        If StringsMatch(TypeName(Application.Caller), "String") Then
            '' Confirm Application.Caller Is a Shape On ActiveSheet
            Dim btnCaller As Shape, btnName As String
            btnName = CStr(Application.Caller)
            Set btnCaller = pbShapeBtn.FindShapeButton(ThisWorkbook.activeSheet, btnName)
            If Not btnCaller Is Nothing Then
                stg.ExecuteButtonAction btnName, ThisWorkbook.activeSheet
            End If
        End If
    End Function


    Public Function QuitOrClose()
        Beep
        Debug.Print "TODO: IMPLEMENT 'QuitOrClose'"
    End Function
    
    Public Function ToggleSettingsAutoHide()
        stg.stdAutoHide = Not stg.stdAutoHide
        If stg.pbSettingsSheet.visible = xlSheetVisible Then
            stg.pbSettingsSheet.Activate
        End If
        ftBeep btButton
    End Function

    Public Function NavDashboard()
        NavigateTo wsDashboard
    End Function
    
    Public Function NavCodeUtil()
        NavigateTo wsCodeUtil
    End Function
    
    
    Public Function NAVSettings()
        If stg.stdAutoHide = True And Not stg.IsDeveloper Then
            MsgBox_FT "Settings Sheet is not available to view because it is hidden and requires developer access.", vbExclamation + vbOKOnly, "OOPS"
        Else
            NavigateTo stg.pbSettingsSheet
        End If
    End Function
    
    Private Function NavigateTo(toWksht As Worksheet, Optional pauseEvents As Boolean = False)
        On Error Resume Next
        Dim evts As Boolean: evts = Events
        Dim objSheet As Object
        If pauseEvents Then Events = False
        If Not toWksht.visible = xlSheetVisible Then toWksht.visible = xlSheetVisible
        toWksht.Activate
        Set objSheet = toWksht
        objSheet.OnFormat
        If Err.number <> 0 Then
            Beep
            LogERROR "Worksheet: " & toWksht.CodeName & " missing 'OnFormat' Method, called from autoPB.NavigateTo"
        End If
        Events = evts
    End Function

