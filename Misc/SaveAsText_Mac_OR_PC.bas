Attribute VB_Name = "SaveAsText_MAC_OR_PC"
Option Explicit
Option Compare Text
Option Base 1

'' There is no need to do anything like use "ChDir"

Public Function SaveWKBKAs(Optional wkbk As Workbook, Optional fileName)
    '1 - Get Active Workbook
    If wkbk Is Nothing Then Set wkbk = ActiveWorkbook
    
    '1 - Get Active Workbook Full Path, without filename
    Dim saveFolder As String
    '' this will get 'saveable path' for local/network file,
    ''  as well as onedrive/sharepoint open files
    saveFolder = FullPathExcludingFileName2(wkbk.FullNameURLEncoded)

    '2 - set filename
    If IsMissing(fileName) Then fileName = "testFile.Txt"
    
    '3 - create full save path
    Dim saveAsPath: saveAsPath = CombinePath(saveFolder, fileName)
    
    '4 -usually a good idea to save before doing a save as, if you don't want to lose changes in source file
        ' see below 'SaveAs' vs. 'SaveCopyAs'-- you may not want to save befare saving out a copy
        
        ' if you want to save first
        wkbk.Save
    
    '5 - define type of 'Text Save As' to use
    Dim saveFormat As XlFileFormat
        '' These aren't all the possible formats, but I'm guessing you want one of these
                'saveFormat = xlCurrentPlatformText
                'saveFormat = xlTextMac
                'saveFormat = xlTextMSDOS
                'saveFormat = xlTextWindows
                'saveFormat = xlUnicodeText
    saveFormat = xlTextMac
    
''''     IF YOU WANT TO MAKE A 'REGULAR' BACKUP FILE , UNCOMMENT THIS SECTION
''''    6 - Perform SaveAs Strategy
''''        SaveCopyAs: Saves a copy of the workbook to a file but doesn't modify the open workbook in memory.
''''        SaveAs: Saves changes to the workbook in a different file.
''''
''''         USE ** SaveCopyAs ** Option
''''         I think you want 'SaveCopyAs', since you mentioned 'backup' -- however SaveCopyAs retains the same format
''''          as the workbook.  I can't imagine why you'd want to save a backup as a text file, but I"ll show you both ways
''''         If using save as, you could define a pattern of naming, something like:
''''
''''        Dim backupFileName As String
''''        backupFileName = "Backup" & Format(Now(), "_yyyyMMMdd-hhnnss_") & FileNameFromFullPath(wkbk.FullNameURLEncoded)
''''        saveAsPath = CombinePath(saveFolder, backupFileName)
''''        wkbk.SaveCopyAs fileName:=saveAsPath
''''
''''         at this point, activeWB still references the workbook that you backed up, and there is now a new file created
''''         at the same location as 'activeWB', but the filename starts with something like: Backup_2024Jan27-050212_[origFileName]
    
    
    
        '' TO MAKE YOUR MAC DO WHAT (I THINK) YOU'RE ASKING FOR
        Dim safeFilePath As String
        safeFilePath = CombinePath(Application.DefaultFilePath, "tmpTextFile1.txt")
        
        Dim desiredFilePath As String
        Dim textFileName As String
        textFileName = "TEXT_" & FileNameWithoutExtension2(FileNameFromFullPath2(wkbk.FullNameURLEncoded)) & ".txt"
        desiredFilePath = CombinePath(FullPathExcludingFileName2(wkbk.FullNameURLEncoded), textFileName)
        
        '' save as txt to 'safe' directory
        Debug.Print "Saving to " & safeFilePath
        ''Here's where the file is actually saved as text.
        wkbk.SaveAs fileName:=safeFilePath, FileFormat:=XlFileFormat.xlTextMac, AddToMru:=False
        
        
        Debug.Print "Opening " & safeFilePath & " into new workbook "
        Dim textWB As Workbook
        ''disable warning, but user still may see a security message they have to answer YES toi
        Application.DisplayAlerts = False
        Set textWB = Workbooks.Open(safeFilePath)
        Debug.Print "Saving copy of " & textWB.FullNameURLEncoded & " TO " & desiredFilePath
        textWB.SaveCopyAs desiredFilePath
        textWB.Close SaveChanges:=False
        Debug.Print "deleting " & safeFilePath
        DeleteFile2 safeFilePath
        Application.DisplayAlerts = True


End Function


Public Function FileNameWithoutExtension2(ByVal fileName As String) As String
    If InStrRev(fileName, ".") > 0 Then
        Dim tmpExt As String
        tmpExt = Mid(fileName, InStrRev(fileName, "."))
        If Len(tmpExt) >= 2 Then
            fileName = replace(fileName, tmpExt, vbNullString)
        End If
    End If
    FileNameWithoutExtension2 = fileName
End Function

Public Function DeleteFile2(filePath As String) As Boolean
    On Error Resume Next
    If FileExists2(filePath) Then
        Kill filePath
        DoEvents
    End If
    DeleteFile2 = FileExists2(filePath) = False
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Returns true if filePth Exists and is not a directory
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function FileExists2(filePth As String, Optional allowWildcardsForFile As Boolean = False) As Boolean
On Error Resume Next
    Dim retV As Boolean
    Dim lastDirName As String, pathArr As Variant, checkFlName As String
    filePth = CombinePath(filePth)

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
        If Err.number <> 0 Then Debug.Print "DirectoryExists: Err Getting Path: " & filePth & ", " & Err.number & " - " & Err.Description
    End If
    FileExists2 = retV
    If Err.number <> 0 Then Err.Clear
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Get File full path, excluding file name
''  Usage:  Dim myPath as String
''
''      myPath = FullPathExcludingFileName(ActiveWorkbook.FullNameURLEncoded)
''      Output: [full directory path of activeworkbook]
''          e.g. on Mac:  "/Users/UserName/Downloads/"
''          e.g. on PC: "C:\Users\UserName\Downloads\"
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function FullPathExcludingFileName2(fullFileName As String) As String
On Error Resume Next
    Dim tPath As String, tfileName As String, fNameStarts As Long
    tfileName = FileNameFromFullPath2(fullFileName)
    fNameStarts = InStr(fullFileName, tfileName)
    tPath = Mid(fullFileName, 1, fNameStarts - 1)
    FullPathExcludingFileName2 = tPath
    If Err.number <> 0 Then Err.Clear
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Get File Name from full path
''  Usage:  Dim myFile as String
''
''      myFile = FileNameFromFullPath(ActiveWorkbook.FullNameURLEncoded)
''      Output: [name of active workbook, with extension]
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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
    If Err.number <> 0 Then Err.Clear
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  MAC and PC Compatible Path Concatenation
''  Usage:  Dim myPath as String
''
''      Ex 1: myPath = CombinePath(Application.DefaultFilePath,"folder1","folder2")
''      Output (Mac): "/Users/paulbrower/Library/Containers/com.microsoft.Excel/Data/Documents/folder1/folder2/"
''
''      Ex 2: myPath = CombinePath(application.StartupPath ,"newFile.txt")
''      Output (Mac) = /Users/paulbrower/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Excel/newFile.txt
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function CombinePath(ParamArray vals() As Variant) As String
        
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
    retV = replace(retV, wrongPS, tDelim)
    If isHTTP Then
        retV = replace(retV, "://", ":::")
        Do While InStr(1, retV, dblPS) > 0
            retV = replace(retV, dblPS, tDelim)
        Loop
        retV = replace(retV, ":::", "://")
    Else
        Do While InStr(1, retV, dblPS) > 0
            retV = replace(retV, dblPS, tDelim)
        Loop
    End If
    If Not isHTTP Then
        If Not Mid(retV, Len(retV)) = Application.PathSeparator Then
            ''get last segment in path
            Dim lastSeg
            lastSeg = Mid(retV, InStrRev(retV, Application.PathSeparator) + 1)
            If Not InStr(lastSeg, ".") > 0 Then
                retV = retV & Application.PathSeparator
            End If
        End If
    End If
    CombinePath = retV
End Function
