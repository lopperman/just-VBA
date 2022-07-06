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

Public Function SaveCopyToUserDocFolder(ByVal wb As Workbook, Optional fileName As Variant)

    SaveWBCopy wb, Application.DefaultFilePath, IIf(IsMissing(fileName), wb.Name, CStr(fileName))

End Function

Public Function SaveWBCopy(ByVal wb As Workbook, dirPath As String, fileName As String)
On Error Resume Next
    Application.EnableEvents = False
    wb.SaveCopyAs PathCombine(False, dirPath, fileName)
    Application.EnableEvents = True
    If Not Err.Number = 0 Then
        MsgBox "SaveWBCopy Error: " & Err.Number & ", " & Err.Description
        Err.Clear
    End If
End Function

Public Property Get StartupPath2() As String
    StartupPath2 = PathCombine(True, Application.StartupPath)
End Property

Public Function FullPathExcludingFileName2(fullFileName As String) As String
On Error Resume Next
    Dim tPath As String, tFileName As String, fNameStarts As Long
    tFileName = FileNameFromFullPath(fullFileName)
    fNameStarts = InStr(fullFileName, tFileName)
    tPath = Mid(fullFileName, 1, fNameStarts - 1)
    FullPathExcludingFileName2 = tPath
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function FileNameFromFullPath2(fullFileName As String) As String
On Error Resume Next
    Dim sepChar As String
    sepChar = Application.PathSeparator
    If LCase(fullFileName) Like "*http*" Then
        sepChar = "/"
    End If
    Dim lastSep As Long: lastSep = Strings.InStrRev(fullFileName, sepChar)
    Dim shortFName As String:  shortFName = Mid(fullFileName, lastSep + 1)
    FileNameFromFullPath2 = shortFName
    If Err.Number <> 0 Then Err.Clear
End Function
Public Function ChooseFolder2(choosePrompt As String) As String
'   Get User-Selected Directory name (MAC and PC Supported)
On Error Resume Next
    Beep
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
    
    ChooseFolder2 = retV
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function ChooseFile2(choosePrompt As String, Optional fileExt As String = vbNullString) As String
'   Get User-Select File Name (MAC and PC Supported)
On Error Resume Next
    Beep
    Dim retV As Variant

    #If Mac Then
        retV = MacScript("choose file with prompt """ & choosePrompt & """ as string")
        If Len(retV) > 0 Then
            retV = MacScript("POSIX path of """ & retV & """")
        End If
    #Else
        Dim fldr As FileDialog
        Dim sItem As String
        Set fldr = Application.FileDialog(msoFileDialogFilePicker)
        With fldr
            .title = choosePrompt
            If Not fileExt = vbNullString Then
                .Filters.Clear
                .Filters.add "Files", fileExt & "?", 1
            End If
            .AllowMultiSelect = False
            If .Show <> -1 Then GoTo NextCode
            retV = .SelectedItems(1)
        End With
NextCode:
        Set fldr = Nothing
    #End If
    
    ChooseFile2 = retV
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function URLEncode2(ByRef txt As String) As String
    Dim buffer As String, i As Long, c As Long, n As Long
    buffer = String$(Len(txt) * 12, "%")
 
    For i = 1 To Len(txt)
        c = AscW(Mid$(txt, i, 1)) And 65535
 
        Select Case c
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95  ' Unescaped 0-9A-Za-z-._ '
                n = n + 1
                Mid$(buffer, n) = ChrW(c)
            Case Is <= 127            ' Escaped UTF-8 1 bytes U+0000 to U+007F '
                n = n + 3
                Mid$(buffer, n - 1) = Right$(Hex$(256 + c), 2)
            Case Is <= 2047           ' Escaped UTF-8 2 bytes U+0080 to U+07FF '
                n = n + 6
                Mid$(buffer, n - 4) = Hex$(192 + (c \ 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case 55296 To 57343       ' Escaped UTF-8 4 bytes U+010000 to U+10FFFF '
                i = i + 1
                c = 65536 + (c Mod 1024) * 1024 + (AscW(Mid$(txt, i, 1)) And 1023)
                n = n + 12
                Mid$(buffer, n - 10) = Hex$(240 + (c \ 262144))
                Mid$(buffer, n - 7) = Hex$(128 + ((c \ 4096) Mod 64))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case Else                 ' Escaped UTF-8 3 bytes U+0800 to U+FFFF '
                n = n + 9
                Mid$(buffer, n - 7) = Hex$(224 + (c \ 4096))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
        End Select
    Next
    URLEncode2 = left$(buffer, n)
End Function

' ~~~~~~~~~~   Create Valid File or Directory Path (for PC or Mac, local, or internet) from 1 or more arguments  ~~~~~~~~~~'
Public Function PathCombine2(includeEndSeparator As Boolean, ParamArray vals() As Variant) As String
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
    PathCombine2 = retV

End Function

Public Function DeleteFolderFiles2(folderPath As String, Optional patternMatch As String = vbNullString)
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
    If Err.Number <> 0 Then Err.Clear
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
    If Err.Number <> 0 Then Err.Clear
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
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function TempDirName2(Optional dirName As String = vbNullString) As String
    TempDirName2 = IIf(Not dirName = vbNullString, dirName, TEMP_DIRECTORY_NAME2)
End Function




