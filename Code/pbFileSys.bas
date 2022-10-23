Attribute VB_Name = "pbFileSys"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' pbFileSys v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' General File Utilities for MAC or PC
'
' @module pbFileSys
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1


Public Const TEMP_DIRECTORY_NAME2 As String = "VBATemp"

Public Function CopySheetToNewWB(ByVal ws As Worksheet, Optional filePath As Variant, Optional fileName As Variant)
On Error Resume Next
    Application.EnableEvents = False
    Dim newWB As Workbook
    Set newWB = Application.Workbooks.Add
    With ws
        .Copy Before:=newWB.Worksheets(1)
        DoEvents
    End With
    If IsMissing(filePath) Then filePath = Application.DefaultFilePath
    If IsMissing(fileName) Then fileName = ReplaceIllegalCharacters2(ws.Name, vbEmpty) & ".xlsx"
    newWB.SaveAs fileName:=PathCombine(False, filePath, fileName), FileFormat:=xlOpenXMLStrictWorkbook
    Application.EnableEvents = True
    If Not Err.number = 0 Then
        MsgBox "CopySheetToNewWB Error: " & Err.number & ", " & Err.Description
        Err.Clear
    End If

End Function

Function ReplaceIllegalCharacters2(strIn As String, strChar As String, Optional padSingleQuote As Boolean = True) As String
    Dim strSpecialChars As String
    Dim i As Long
    strSpecialChars = "~""#%&*:<>?{|}/\[]" & Chr(10) & Chr(13)

    For i = 1 To Len(strSpecialChars)
        strIn = Replace(strIn, Mid$(strSpecialChars, i, 1), strChar)
    Next
    
    If padSingleQuote And InStr(1, strIn, "''") = 0 Then
        strIn = CleanSingleTicks(strIn)
    End If
    
    ReplaceIllegalCharacters2 = strIn
End Function

    ' ~~~~~~~~~~   CLEAN SINGLE TICKS ~~~~~~~~~~'
    Public Function CleanSingleTicks(ByVal wbName As String) As String
        Dim retV As String
        If InStr(wbName, "'") > 0 And InStr(wbName, "''") = 0 Then
            retV = Replace(wbName, "'", "''")
        Else
            retV = wbName
        End If
        CleanSingleTicks = retV
    End Function


    Public Function SaveCopyToUserDocFolder(ByVal wb As Workbook, Optional fileName As Variant)
        SaveWBCopy wb, Application.DefaultFilePath, IIf(IsMissing(fileName), wb.Name, CStr(fileName))
    End Function

Public Function SaveWBCopy(ByVal wb As Workbook, dirPath As String, fileName As String)
On Error Resume Next
    Application.EnableEvents = False
    wb.SaveCopyAs PathCombine(False, dirPath, fileName)
    Application.EnableEvents = True
    If Not Err.number = 0 Then
        MsgBox "SaveWBCopy Error: " & Err.number & ", " & Err.Description
        Err.Clear
    End If
End Function

Public Function OpenPath(fldrPath As String)
'   Open Folder (MAC and PC Supported)
On Error Resume Next
    ftBeep btMsgBoxChoice
    Dim retV As Variant

    #If Mac Then
        Dim scriptStr As String
        scriptStr = "do shell script " & Chr(34) & "open " & fldrPath & Chr(34)
        MacScript (scriptStr)
    #Else
        Call Shell("explorer.exe " & fldrPath, vbNormalFocus)
    #End If
    
    If Err.number <> 0 Then
        LogError "pbFileSys.OpenFolder - Error Opening: (" & fldrPath & ") - " & ErrString
        Err.Clear
    End If
End Function

' ~~~~~~~~~~   Create Valid File or Directory Path (for PC or Mac, local, or internet) from 1 or more arguments  ~~~~~~~~~~'
Public Function PathCombine(includeEndSeparator As Boolean, ParamArray vals() As Variant) As String
' COMBINE PATH AND/OR FILENAME SEGMENTS
' WORKS FOR MAC OR PC ('/' vs '\'), and for web url's
'
'   DebugPrint PathCombine(True, "/usr", "\\what", "/a//", "mess")
'      outputs:  /usr/what/a/mess/
'   DebugPrint PathCombine(False, "/usr", "\\what", "/a//", "mess", "word.docx/")
'      outputs: /usr/what/a/mess/word.docx
'   DebugPrint PathCombine(true,"https://www.google.com\badurl","gmail")
'       outputs:  https://www.google.com/badurl/gmail/
    
    Dim tDelim As String, isHTTP As Boolean
    Dim i As Long
    Dim retV As String
    Dim dblPS As String
    Dim wrongPS As String
    For i = LBound(vals) To UBound(vals)
        If LCase(vals(i)) Like "*http*" Then
            isHTTP = True
            tDelim = "/"
            wrongPS = "\"
        End If
    Next i
    If Not isHTTP Then
        tDelim = Application.PathSeparator
        If InStr(1, "/", Application.PathSeparator) > 0 Then
            wrongPS = "\"
        Else
            wrongPS = "/"
        End If
    End If
    dblPS = tDelim & tDelim
    For i = LBound(vals) To UBound(vals)
        If i = LBound(vals) Then
            retV = CStr(vals(i))
            If Len(retV) = 0 Then retV = tDelim
        Else
            If Mid(retV, Len(retV)) = tDelim Then
                retV = retV & vals(i)
            Else
                retV = retV & tDelim & vals(i)
            End If
        End If
    Next i
    retV = Replace(retV, wrongPS, tDelim)
    If isHTTP Then
        retV = Replace(retV, "://", ":::")
        Do While InStr(1, retV, dblPS) > 0
            retV = Replace(retV, dblPS, tDelim)
        Loop
        retV = Replace(retV, ":::", "://")
    Else
        Do While InStr(1, retV, dblPS) > 0
            retV = Replace(retV, dblPS, tDelim)
        Loop
    End If
    If includeEndSeparator Then
        If Not Mid(retV, Len(retV)) = tDelim Then
            retV = retV & Application.PathSeparator
        End If
    Else
        'Remove it if it's there
        If Mid(retV, Len(retV)) = Application.PathSeparator Then
            retV = Mid(retV, 1, Len(retV) - 1)
        End If
    End If
    PathCombine = retV

End Function

Public Function FullPathExcludingFileName(fullFileName As String) As String
On Error Resume Next
    Dim tPath As String, tfileName As String, fNameStarts As Long
    tfileName = FileNameFromFullPath(fullFileName)
    fNameStarts = InStr(fullFileName, tfileName)
    tPath = Mid(fullFileName, 1, fNameStarts - 1)
    FullPathExcludingFileName = tPath
    If Err.number <> 0 Then Err.Clear
End Function

Public Function FileNameFromFullPath(fullFileName As String) As String
On Error Resume Next
    Dim sepChar As String
    sepChar = Application.PathSeparator
    If LCase(fullFileName) Like "*http*" Then
        sepChar = "/"
    End If
    Dim lastSep As Long: lastSep = Strings.InStrRev(fullFileName, sepChar)
    Dim shortFName As String:  shortFName = Mid(fullFileName, lastSep + 1)
    FileNameFromFullPath = shortFName
    If Err.number <> 0 Then Err.Clear
End Function
Public Function ChooseFolder(choosePrompt As String) As String
'   Get User-Selected Directory name (MAC and PC Supported)
On Error Resume Next
    ftBeep btMsgBoxChoice
    Dim retV As Variant

    #If Mac Then
        retV = MacScript("choose folder with prompt """ & choosePrompt & """ as string")
        If Len(retV) > 0 Then
            retV = MacScript("POSIX path of """ & retV & """")
        End If
    #Else
        Dim fldr As FileDialog
        Dim sItem As String
        Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
        With fldr
            .title = choosePrompt
            .AllowMultiSelect = False
            .InitialFileName = Application.DefaultFilePath
            If .Show <> -1 Then GoTo NextCode
            retV = .SelectedItems(1)
        End With
NextCode:
        Set fldr = Nothing
    #End If
    
    ChooseFolder = retV
    If Err.number <> 0 Then Err.Clear
End Function

Public Function RequestFileAccess(ParamArray files() As Variant)

    #If Mac Then
        'Declare Variables?
        Dim fileAccessGranted As Boolean
        Dim filePermissionCandidates
    
        'Create an array with file paths for the permissions that are needed.?
    '    filePermissionCandidates = Array("/Users//Desktop/test1.txt", "/Users//Desktop/test2.txt")
        filePermissionCandidates = files
    
        'Request access from user.?
        fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)
        'Returns true if access is granted; otherwise, false.
    #End If
End Function

Public Function FileNameWithoutExtension(ByVal fileName As String) As String
    If InStrRev(fileName, ".") > 0 Then
        Dim tmpExt As String
        tmpExt = Mid(fileName, InStrRev(fileName, "."))
        If Len(tmpExt) >= 2 Then
            fileName = Replace(fileName, tmpExt, vbNullString)
        End If
    End If
    FileNameWithoutExtension = fileName
End Function

Public Function SaveFileAs(savePrompt As String, Optional ByVal defaultFileName, Optional ByVal fileExt) As String
On Error Resume Next
    ftBeep btMsgBoxChoice
    Dim retV As Variant
        
    #If Mac Then
        If Len(fileExt) > 0 Then
            
            fileExt = Replace(Replace(fileExt, "*", ""), ".", "")
            retV = Application.GetSaveAsFilename(InitialFileName:=IIf(IsMissing(defaultFileName), "", defaultFileName), FileFilter:=IIf(IsMissing(fileExt), "", fileExt), ButtonText:="USE THIS NAME")
        Else
            retV = Application.GetSaveAsFilename(InitialFileName:=IIf(IsMissing(defaultFileName), "", defaultFileName), FileFilter:=IIf(IsMissing(fileExt), "", fileExt), ButtonText:="USE THIS NAME")
        End If
    #Else
NextCode:
        If Len(fileExt) > 0 Then
            fileExt = Replace(Replace(fileExt, "*", ""), ".", "")
            fileExt = Concat("*.", fileExt, "*")
            fileExt = "Files (" & fileExt & "), " & fileExt
            If Len(defaultFileName) > 0 Then
                retV = Application.GetSaveAsFilename(InitialFileName:=defaultFileName, FileFilter:=fileExt, title:=savePrompt, ButtonText:="USE THIS NAME")
            Else
                retV = Application.GetSaveAsFilename(FileFilter:=fileExt, title:=savePrompt, ButtonText:="USE THIS NAME")
            End If
            'retV = Application.GetOpenFilename(InitialFileName:=IIf(IsMissing(defaultFileName), "", defaultFileName), FileFilter:=fileExt, title:=choosePrompt, ButtonText:="USE THIS NAME")
        End If
    
    #End If
    
    If Err.number = 0 Then
        SaveFileAs = CStr(retV)
    Else
        LogError "pbFileSys.ChooseFile " & ErrString
        Err.Clear
    End If

End Function

Public Function ChooseFile(choosePrompt As String, Optional ByVal fileExt As String = vbNullString) As String
'TODO:  Also check out Application.GetSaveAsFileName
'   Get User-Select File Name (MAC and PC Supported)
On Error Resume Next
    ftBeep btMsgBoxChoice
    Dim retV As Variant
        
    #If Mac Then
        If Len(fileExt) > 0 Then
            fileExt = Replace(Replace(fileExt, "*", ""), ".", "")
            retV = Application.GetOpenFilename(FileFilter:=fileExt, ButtonText:="CHOOSE FILE")
        Else
            retV = Application.GetOpenFilename(title:=choosePrompt)
        End If
    #Else
NextCode:
        If Len(fileExt) > 0 Then
            fileExt = Replace(Replace(fileExt, "*", ""), ".", "")
            fileExt = Concat("*.", fileExt, "*")
            fileExt = "Files (" & fileExt & "), " & fileExt
            retV = Application.GetOpenFilename(FileFilter:=fileExt, title:=choosePrompt, ButtonText:="CHOOSE FILE")
        Else
            retV = Application.GetOpenFilename(title:=choosePrompt, ButtonText:="CHOOSE FILE")
        End If
    
    #End If
    
    If Err.number = 0 Then
        ChooseFile = CStr(retV)
    Else
        LogError "pbFileSys.ChooseFile " & ErrString
        Err.Clear
    End If

End Function



' ~~~~~~~~~~   CREATE THE ** LAST ** DIRECTORY IN 'fullPath' ~~~~~~~~~~'
Public Function CreateDirectory(fullPath As String) As Boolean
' IF 'fullPath' is not a valid directory but the '1 level back' IS a valid directory, then the last directory in 'fullPath' will be created
' Example: CreateDirectory("/Users/paul/Library/Containers/com.microsoft.Excel/Data/Documents/FinToolTemp/Logs")
    'If the 'FinToolTemp' directory exists, the Logs will be created if it is not already there.
'   Primary reason for not creating multiple directories in the path is issues with both PC and Mac for File System changes.
    
    LogTrace ConcatWithDelim(", ", "pbMiscUtil.CreateDirectory", "CHECKING", fullPath)
    
    Dim retV As Boolean

    If DirectoryExists(fullPath) Then
        DebugPrint ConcatWithDelim(", ", "pbMiscUtil.CreateDirectory", fullPath, "aready exists")
        retV = True
    Else
        Dim lastDirName As String, pathArr As Variant, checkFldrName As String
        fullPath = PathCombine(False, fullPath)
        If InStrRev(fullPath, Application.PathSeparator, compare:=vbTextCompare) > InStr(1, fullPath, Application.PathSeparator, vbTextCompare) Then
            lastDirName = Left(fullPath, InStrRev(fullPath, Application.PathSeparator, compare:=vbTextCompare) - 1)
            If DirectoryExists(lastDirName) Then
                On Error Resume Next
                DebugPrint ConcatWithDelim(", ", "pbMiscUtil.CreateDirectory", "Creating directory: ", fullPath)
                MkDir fullPath
                If Err.number = 0 Then
                    DebugPrint ConcatWithDelim(", ", "pbMiscUtil.CreateDirectory", "Created: ", fullPath)
                
                    retV = DirectoryExists(fullPath)
                End If
            End If
        End If
    End If
    CreateDirectory = retV
    If Err.number <> 0 Then Err.Clear
End Function

' ~~~~~~~~~~   Returns true if filePth Exists and is not a directory  ~~~~~~~~~~'
Public Function FileExists(filePth As String, Optional allowWildcardsForFile As Boolean = False) As Boolean
On Error Resume Next
    Dim retV As Boolean
    Dim lastDirName As String, pathArr As Variant, checkFlName As String
    filePth = PathCombine(False, filePth)

    If InStr(filePth, Application.PathSeparator) > 0 Then
        pathArr = Split(filePth, Application.PathSeparator)
        checkFlName = CStr(pathArr(UBound(pathArr)))
        Dim tmpReturnedFileName As String
        tmpReturnedFileName = Dir(filePth & "*", vbNormal)
        If allowWildcardsForFile = True And Len(tmpReturnedFileName) > 0 Then
            retV = True
        Else
            retV = StrComp(Dir(filePth & "*"), LCase(checkFlName), vbTextCompare) = 0
        End If
        If Err.number <> 0 Then DebugPrint "DirectoryExists: Err Getting Path: " & filePth & ", " & Err.number & " - " & Err.Description
    End If
    FileExists = retV
    If Err.number <> 0 Then Err.Clear
End Function

' ~~~~~~~~~~   Returns true if DIRECTORY path dirPath) Exists ~~~~~~~~~~'
Public Function DirectoryExists(dirPath As String) As Boolean
On Error Resume Next
    Dim retV As Boolean
    Dim lastDirName As String, pathArr As Variant, checkFldrName As String
    dirPath = PathCombine(False, dirPath)

    If InStr(dirPath, Application.PathSeparator) > 0 Then
        pathArr = Split(dirPath, Application.PathSeparator)
        checkFldrName = CStr(pathArr(UBound(pathArr)))
        retV = StrComp(Dir(dirPath & "*", vbDirectory), LCase(checkFldrName), vbTextCompare) = 0
        If Err.number <> 0 Then
            DebugPrint "DirectoryExists: Err Getting Path: " & dirPath & ", " & Err.number & " - " & Err.Description
        End If
    End If
    DirectoryExists = retV
    If Err.number <> 0 Then Err.Clear
End Function

Public Function DeleteFolderFiles(folderPath As String, Optional patternMatch As String = vbNullString)
On Error Resume Next
    folderPath = PathCombine(True, folderPath)
    
    If DirectoryFileCount(folderPath) > 0 Then
        Dim myPath As Variant
        myPath = PathCombine(True, folderPath)
        ChDir folderPath
        Dim myFile, MyName As String
        MyName = Dir(myPath, vbNormal)
        Do While MyName <> ""
            If (GetAttr(PathCombine(False, myPath, MyName)) And vbNormal) = vbNormal Then
                If patternMatch = vbNullString Then
                    Kill PathCombine(False, myPath, MyName)
                Else
                    If LCase(MyName) Like LCase(patternMatch) Then
                        Kill PathCombine(False, myPath, MyName)
                    End If
                End If
            End If
            MyName = Dir()
        Loop
    End If
    If Err.number <> 0 Then Err.Clear
End Function



Public Function DirectoryFileCount(tmpDirPath As String) As Long
On Error Resume Next

    Dim myFile, myPath, MyName As String, retV As Long
    myPath = PathCombine(True, tmpDirPath)
    MyName = Dir(myPath, vbNormal)
    Do While MyName <> ""
        If (GetAttr(PathCombine(False, myPath, MyName)) And vbNormal) = vbNormal Then
            retV = retV + 1
        End If
        MyName = Dir()
    Loop
    DirectoryFileCount = retV
    If Err.number <> 0 Then Err.Clear
End Function

Public Function DirectoryDirectoryCount(tmpDirPath As String) As Long
On Error Resume Next

    Dim myFile, myPath, MyName As String, retV As Long
    myPath = PathCombine(True, tmpDirPath)
    MyName = Dir(myPath, vbDirectory)
    Do While MyName <> ""
        If (GetAttr(PathCombine(False, myPath, MyName)) And vbDirectory) = vbDirectory Then
            retV = retV + 1
        End If
        MyName = Dir()
    Loop
    DirectoryDirectoryCount = retV
    If Err.number <> 0 Then Err.Clear
End Function

Public Function DeleteFile(filePath As String) As Boolean
    On Error Resume Next
    If FileExists(filePath) Then
        Kill filePath
        DoEvents
    End If
    DeleteFile = FileExists(filePath) = False
End Function



Public Function GetFiles(dirPath As String) As Variant()
On Error Resume Next
    
    Dim cl As New Collection

    Dim myFile, myPath, MyName As String
    myPath = PathCombine(True, dirPath)
    MyName = Dir(myPath, vbNormal)
    Do While MyName <> ""
        If (GetAttr(PathCombine(False, myPath, MyName)) And vbNormal) = vbNormal Then
            cl.Add MyName
        End If
        MyName = Dir()
    Loop
    
    If cl.Count > 0 Then
        Dim retV() As Variant
        ReDim retV(1 To cl.Count, 1 To 1)
        Dim fidx As Long
        For fidx = 1 To cl.Count
            retV(fidx, 1) = cl(fidx)
        Next fidx
        GetFiles = retV
    End If
    
    If Err.number <> 0 Then Err.Clear
    
End Function

Public Function DirectoryFileCount2(tmpDirPath As String) As Long
On Error Resume Next

    Dim myFile, myPath, MyName As String, retV As Long
    myPath = PathCombine(True, tmpDirPath)
    MyName = Dir(myPath, vbNormal)
    Do While MyName <> ""
        If (GetAttr(PathCombine(False, myPath, MyName)) And vbNormal) = vbNormal Then
            retV = retV + 1
        End If
        MyName = Dir()
    Loop
    DirectoryFileCount2 = retV
    If Err.number <> 0 Then Err.Clear
End Function

Public Function DirectoryDirectoryCount2(tmpDirPath As String) As Long
On Error Resume Next

    Dim myFile, myPath, MyName As String, retV As Long
    myPath = PathCombine(True, tmpDirPath)
    MyName = Dir(myPath, vbDirectory)
    Do While MyName <> ""
        If (GetAttr(PathCombine(False, myPath, MyName)) And vbDirectory) = vbDirectory Then
            retV = retV + 1
        End If
        MyName = Dir()
    Loop
    DirectoryDirectoryCount2 = retV
    If Err.number <> 0 Then Err.Clear
End Function


Public Function TempDirName2(Optional dirName As String = vbNullString) As String
    TempDirName2 = IIf(Not dirName = vbNullString, dirName, TEMP_DIRECTORY_NAME2)
End Function




