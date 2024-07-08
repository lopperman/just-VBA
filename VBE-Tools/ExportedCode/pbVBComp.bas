Attribute VB_Name = "pbVBComp"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  HELPER CODE FOR MANAGING OBJECTS THAT NEED TO BE
''  PART OF SOURCE CODE CONTROL (E.G. GIT)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  author (c) Paul Brower https://github.com/lopperman/just-VBA
''  module pbVBComp.bas
''  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  DEPENDENCIES
''  pbCommonUtil.bas
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1

Public Const CODE_LINE_SEPARATOR As String = "' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '"
Public Const CODE_LINE_DOUBLE_TICK As String = "''"


Private lSavePath As String
Private Const SETTING_APP_NAME As String = "justVBA"
Private Const SETTING_SECTION_NAME As String = "pbVBComp"
Private Const SETTING_SAVE_KEY As String = "SavePath"

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'' BECAUSE THIS WON'T WORK AS A CONSTANT'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Get CODE_TAB_SPACES() As String
    CODE_TAB_SPACES = Strings.Space$(4)
End Property

Public Function ExportSupported(Optional vbaFile As Workbook, Optional lbl As String)
    If vbaFile Is Nothing Then
        Set vbaFile = ThisWorkbook
    End If
    ExportCode vbaFile, vbext_ct_Document, lbl
    ExportCode vbaFile, vbext_ct_ClassModule, lbl
    ExportCode vbaFile, vbext_ct_StdModule, lbl
    ExportCode vbaFile, vbext_ct_MSForm, lbl
    ExportCode vbaFile, vbext_ct_ActiveXDesigner, lbl
End Function

Public Function ShowWBIndexes(Optional showHasVBProjectOnly As Boolean = False)
    Dim thisIDX As Long
    Dim wb As Workbook
    Dim i As Long, validOutput As Boolean
    For i = 1 To Application.Workbooks.Count
        If Application.Workbooks(i) Is ThisWorkbook Then
            thisIDX = i
            Exit For
        End If
    Next i
    Debug.Print CODE_LINE_SEPARATOR
    Debug.Print Concat(CODE_TAB_SPACES, "pbVBComp running in: ", ThisWorkbook.Name, " (Index: " & thisIDX & ")")
    Debug.Print CODE_LINE_SEPARATOR
    For i = 1 To Application.Workbooks.Count
        validOutput = False
        If Application.Workbooks(i) Is ThisWorkbook Then
            validOutput = False
        ElseIf Application.Workbooks(i).HasVBProject = False Then
            If showHasVBProjectOnly = False Then validOutput = True
        Else
            validOutput = True
        End If
        If validOutput Then
            Debug.Print Concat(CODE_TAB_SPACES, Application.Workbooks(i).Name, " (Index: " & i & ")")
        End If
    Next i
    Debug.Print CODE_LINE_SEPARATOR
    
End Function

Private Function GetExtension(item As VBComponent) As String
    Select Case item.Type
        Case vbext_ct_ClassModule
            GetExtension = "cls"
        Case vbext_ct_StdModule, vbext_ct_Document, vbext_ct_MSForm
            GetExtension = "bas"
        Case Else
            GetExtension = "txt"
        End Select
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  EXPORT
''  @vbaFile - Workbook from which to export code items
''  @exportType - vb Component Type to export
''  @lbl - name to use for directory code will be exported to
''      If no value provide, will use yyyyMMMdd for folder name
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function ExportCode(vbaFile As Workbook, exportType As vbext_ComponentType, Optional lbl As String, Optional componentName)
    If Not IsExportSupported(exportType) Then
        MsgBox "Sorry, that component type is not yet supported for exporting"
        Exit Function
    End If
    If Len(lbl) = 0 Then
        lbl = FileNameWithoutExtension(vbaFile.Name) & UCase(Format(Now(), "-yyyyMMMdd-hhnnss"))
    End If
    If CheckDirectory(PathCombine(True, SavePath, lbl)) = False Then Exit Function
    Dim exportToDir As String, fileExt As String
    exportToDir = PathCombine(True, SavePath, lbl)
    If exportType = vbext_ct_ClassModule Then
        fileExt = "cls"
    ElseIf exportType = vbext_ct_StdModule Then
        fileExt = "bas"
    ElseIf exportType = vbext_ct_Document Then
        fileExt = "bas"
    ElseIf exportType = vbext_ct_MSForm Then
        fileExt = "bas"
    Else
        fileExt = "txt"
    End If
    Dim VBItem As VBComponent, compFileName As String, includeItem As Boolean
    For Each VBItem In vbaFile.VBProject.VBComponents
        If VBItem.Type = exportType Then
            includeItem = True
            If Not IsMissing(componentName) Then
                includeItem = StringsMatch(VBItem.Name, componentName)
            End If
            If includeItem Then
                compFileName = VBItem.Name
                VBItem.Export PathCombine(False, exportToDir, Concat(compFileName, ".", fileExt))
                LogTRACE ConcatWithDelim(" ", stg.UserNameOrLogin, "Exported:", compFileName & "." & fileExt, ", to", exportToDir)
                
            End If
        End If
    Next
End Function
Private Function IsExportSupported(exportType As vbext_ComponentType) As Boolean
    Select Case exportType
        Case _
            vbext_ComponentType.vbext_ct_ClassModule _
            , vbext_ComponentType.vbext_ct_StdModule _
            , vbext_ComponentType.vbext_ct_Document _
            , vbext_ComponentType.vbext_ct_MSForm _
            , vbext_ComponentType.vbext_ct_ActiveXDesigner
            IsExportSupported = True
    End Select
End Function



' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  SAVE PATH (Setter)
''  Set Path to Directory Where Files will be saved
''  The value persists when set (via 'SaveSetting')
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Let SavePath(directoryFullPath As String)
    SaveSetting SETTING_APP_NAME, SETTING_SECTION_NAME, SETTING_SAVE_KEY, directoryFullPath
    lSavePath = vbNullString
End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  SAVE PATH (Getter)
''  Get Path to Directory Where Files will be saved
''  If set previously, does not need to be set again unless path has changed
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Get SavePath() As String
    Dim resp As String
    If Len(lSavePath) > 0 Then
        resp = lSavePath
    Else
        Dim tmpPath As String
        tmpPath = GetSetting(SETTING_APP_NAME, SETTING_SECTION_NAME, SETTING_SAVE_KEY)
        If Len(tmpPath) > 0 Then
            lSavePath = tmpPath
            resp = lSavePath
        End If
    End If
    If CheckDirectory(resp) = False Then
        If CreateDirectory(resp) Then
            SavePath = resp
        End If
    Else
        SavePath = resp
    End If
End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  CHECK DIRECTORY
''  Verifies the full path ([path]) is valid
''  If directory does not exist, will attempt to create it
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function CheckDirectory(path As String) As Boolean
    If Len(path) = 0 Then
        CheckDirectory = False
    ElseIf DirectoryExists(path) Then
        CheckDirectory = True
    Else
        On Error Resume Next
        CheckDirectory = CreateDirectory(path)
        If Err.number <> 0 Then
            Err.Clear
        End If
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  SAVE PATH VALID
''  Returns true if 'SavePath' is a valid directory
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function SavePathValid() As Boolean
    If Len(SavePath) > 0 Then
        SavePathValid = CheckDirectory(SavePath)
    End If
End Function

Public Function IsCodePaneOpen(wkbk As Workbook, componentName As String) As Boolean
    Dim cp As CodePane
    
    For Each cp In wkbk.VBProject.VBE.CodePanes
        Debug.Print cp.CodeModule & " Window Visible: "; cp.Window.visible
    Next
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  CLOSE ALL OPEN VBCOMPONENT WINDOWS, EXCEPT FOR [visibleComp]
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function devVBECloseWindows(ParamArray visibleCompNames() As Variant)
    Dim tWB As Workbook, tVBComp As VBComponent
    Dim tVisible As Boolean
    For Each tWB In Application.Workbooks
        If tWB.HasVBProject Then
            For Each tVBComp In tWB.VBProject.VBComponents
                tVisible = False
                Dim vComp
                For Each vComp In visibleCompNames
                    If StringsMatch(vComp, tVBComp.CodeModule) Then
                        tVisible = True
                        Exit For
                    End If
                Next
                tVBComp.CodeModule.CodePane.Window.visible = tVisible
            Next
        End If
    Next
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  REPLACE VB COMPONENT IN OPEN WORKBOOK, WITH VERSION
''      FROM ANOTHER OPEN WORKBOOK
''  Returns true if SUCCESS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function VBEdevUpdateComponent( _
    srcComponentName As String _
    , Optional sourceWB As pbWKBK _
    , Optional targetWB As Workbook _
    , Optional canOverwriteDestination As Boolean = True _
    , Optional addIfMissing As Boolean = True) As Boolean
    On Error Resume Next
    If sourceWB Is Nothing Then Set sourceWB = pbWKBK
    Dim srcComp As VBComponent
    Dim saveFullPath As String
    Set srcComp = sourceWB.VBProject.VBComponents(srcComponentName)
    If Err.number <> 0 Then
        Err.Clear
        Debug.Print "Unable to find: '" & srcComponentName & "' in " & sourceWB.Name
        Exit Function
    End If
    saveFullPath = PathCombine(True, Application.DefaultFilePath, "pvVBCompTemp")
    If Not DirectoryExists(saveFullPath) Then CreateDirectory (saveFullPath)
    If DirectoryExists(saveFullPath) Then
        If targetWB Is Nothing Then
            Dim tWB As Workbook
            For Each tWB In Application.Workbooks
                If Not tWB Is sourceWB And tWB.HasVBProject Then
                    VBEdevUpdateComponent = VBEdevUpdateComponent(srcComponentName, sourceWB, tWB, canOverwriteDestination, addIfMissing)
                End If
            Next tWB
            Exit Function
        End If
        Dim targetExists As Boolean
        Dim existTarget As VBComponent
        Set existTarget = targetWB.VBProject.VBComponents(srcComponentName)
        targetExists = Not existTarget Is Nothing
        If Err.number <> 0 Then
            Err.Clear
        End If
        If targetExists Then
            sourceWB.VBProject.VBComponents("aaDELETE").Export PathCombine(False, saveFullPath, "aaDELETE.bas")
            targetWB.VBProject.VBComponents.Import PathCombine(False, saveFullPath, "aaDELETE.bas")
            DoEvents
            targetWB.VBProject.VBComponents.remove targetWB.VBProject.VBComponents("aaDELETE")
            DoEvents
            
            
            If existTarget.Saved = False Then targetWB.Save
            existTarget.CodeModule.CodePane.Window.visible = False
        End If
        If targetExists And canOverwriteDestination = False Then
            Debug.Print "Stopped - Not able to overwrite"
            Exit Function
        End If
        If targetExists = False And addIfMissing = False Then
            Debug.Print "Stopped - Not able to create if missing"
            Exit Function
        End If
        
        Application.DisplayAlerts = True
        If srcComp.Saved = False Then sourceWB.Save
        Dim exportPath As String
        exportPath = PathCombine(False, saveFullPath, srcComp.Name & "." & GetExtension(srcComp))
        srcComp.Export exportPath
        If Err.number = 0 Then
            If Not existTarget Is Nothing Then
                targetWB.VBProject.VBComponents.remove existTarget
            End If
            targetWB.VBProject.VBComponents.Import exportPath
            With targetWB.VBProject.VBComponents(srcComponentName).CodeModule
                .InsertLines 1, "' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '"
                .InsertLines 2, "''     *** DO NOT EDIT "
                .InsertLines 3, "''     *** AUTHOR: PAUL BROWER - CREATED: " & CStr(Now())
                .InsertLines 3, "''     *** INSERTED AUTOMATICALLY FROM: " & sourceWB.Name
                .InsertLines 5, "' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '"
               .InsertLines 6, "''"
                 
            End With
            VBEdevUpdateComponent = (Err.number = 0)
        End If
    End If
    
    If Err.number <> 0 Then
        Debug.Print " *** FAILED ***"
        Debug.Print Err.number, Err.Description
        Err.Clear
    Else
        Debug.Print " *** success ***"
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  EXPORT
''  @vbaFile - Workbook from which to find VBComponents to document
''      If missing, Uses Workbook where this code is being executued
''  @componentType - If provided will ignore all non matchine component
''      types
''  @componentName - if provided, will only document component
''      matching name
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function ShowComponentInfo(Optional vbaFile As Workbook, Optional ComponentType As vbext_ComponentType, Optional componentName)
    Dim VBItem As VBComponent, compFileName As String, includeItem As Boolean
    If vbaFile Is Nothing Then
        Set vbaFile = ThisWorkbook
    End If
    Dim invalidComp As Boolean
    For Each VBItem In vbaFile.VBProject.VBComponents
        invalidComp = False
        If ComponentType <> 0 And VBItem.Type <> ComponentType Then
            invalidComp = True
        End If
        If Not IsMissing(componentName) And Not StringsMatch(componentName, VBItem.Name) Then
            invalidComp = True
        End If
        If Not invalidComp Then
            Dim iLine As Long, totLines As Long
            totLines = VBItem.CodeModule.CountOfLines
            Dim codeLine As String
            If VBItem.CodeModule.CountOfLines > 0 Then
                For iLine = 1 To totLines
                    codeLine = VBItem.CodeModule.lines(iLine, 1)
                    If Not StringsMatch("'", codeLine, smStartsWithStr) Then
                        If StringsMatch(codeLine, "Public", smContains) Then
                            Debug.Print VBItem.Name, "Line " & iLine, codeLine
                        End If
                    End If
                Next iLine
            End If
        End If
    Next
End Function

Public Function pbCompareCode(wkbk1 As Workbook, item1, wkbk2 As Workbook, item2)
    Dim vbc As New VBItems
    vbc.CompareVBComp wkbk1, CStr(item1), wkbk2, CStr(item2)
End Function

