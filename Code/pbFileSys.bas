Option Explicit
Option Compare Text
Option Base 1


Public Const TEMP_DIRECTORY_NAME2 As String = "VBATemp"

Public Function CopySheetToNewWB(ByVal ws As Worksheet, Optional filepath As Variant, Optional fileName As Variant)
On Error Resume Next
    Application.EnableEvents = False
    Dim newWB As Workbook
    Set newWB = Application.Workbooks.Add
    With ws
        .Copy Before:=newWB.Worksheets(1)
        DoEvents
    End With
    If IsMissing(filepath) Then filepath = Application.DefaultFilePath
    If IsMissing(fileName) Then fileName = ReplaceIllegalCharacters2(ws.Name, vbEmpty) & ".xlsx"
    newWB.SaveAs fileName:=PathCombine2(False, filepath, fileName), FileFormat:=xlOpenXMLStrictWorkbook
    Application.EnableEvents = True
    If Not Err.Number = 0 Then
        MsgBox "CopySheetToNewWB Error: " & Err.Number & ", " & Err.Description
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
        strIn = CleanSingleTicks2(strIn)
    End If
    
    ReplaceIllegalCharacters2 = strIn
End Function

Public Function CleanSingleTicks2(wbName As String) As String
    Dim retV As String
    If InStr(wbName, "'") > 0 And InStr(wbName, "''") = 0 Then
        retV = Replace(wbName, "'", "''")
    Else
        retV = wbName
    End If
    CleanSingleTicks2 = retV
End Function


    Public Function SaveCopyToUserDocFolder(ByVal wb As Workbook, Optional fileName As Variant)
        SaveWBCopy wb, Application.DefaultFilePath, IIf(IsMissing(fileName), wb.Name, CStr(fileName))
    End Function

Public Function SaveWBCopy(ByVal wb As Workbook, dirPath As String, fileName As String)
On Error Resume Next
    Application.EnableEvents = False
    wb.SaveCopyAs PathCombine2(False, dirPath, fileName)
    Application.EnableEvents = True
    If Not Err.Number = 0 Then
        MsgBox "SaveWBCopy Error: " & Err.Number & ", " & Err.Description
        Err.Clear
    End If
End Function

Public Property Get StartupPath2() As String
    StartupPath2 = PathCombine2(True, Application.StartupPath)
End Property

Public Function FullPathExcludingFileName2(fullFileName As String) As String
On Error Resume Next
    Dim tPath As String, tFileName As String, fNameStarts As Long
    tFileName = FileNameFromFullPath2(fullFileName)
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
            .Title = choosePrompt
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
            .Title = choosePrompt
            If Not fileExt = vbNullString Then
                .Filters.Clear
                .Filters.Add "Files", fileExt & "?", 1
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
    URLEncode2 = Left$(buffer, n)
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
    folderPath = PathCombine2(True, folderPath)
    
    If DirectoryFileCount2(folderPath) > 0 Then
        Dim myPath As Variant
        myPath = PathCombine2(True, folderPath)
        ChDir folderPath
        Dim myFile, MyName As String
        MyName = Dir(myPath, vbNormal)
        Do While MyName <> ""
            If (GetAttr(PathCombine2(False, myPath, MyName)) And vbNormal) = vbNormal Then
                If patternMatch = vbNullString Then
                    Kill PathCombine2(False, myPath, MyName)
                Else
                    If LCase(MyName) Like LCase(patternMatch) Then
                        Kill PathCombine2(False, myPath, MyName)
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
    myPath = PathCombine2(True, tmpDirPath)
    MyName = Dir(myPath, vbNormal)
    Do While MyName <> ""
        If (GetAttr(PathCombine2(False, myPath, MyName)) And vbNormal) = vbNormal Then
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
    myPath = PathCombine2(True, tmpDirPath)
    MyName = Dir(myPath, vbDirectory)
    Do While MyName <> ""
        If (GetAttr(PathCombine2(False, myPath, MyName)) And vbDirectory) = vbDirectory Then
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
