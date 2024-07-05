Attribute VB_Name = "aaSandbox"
Option Explicit
Option Compare Text
Option Base 1



Public Function testSort()

    Dim sw As New StopWatch
    Dim maxLen As Long, ascStart As Long, ascEnd As Long
    ascStart = 65
    ascEnd = 122
    maxLen = 50
    
    Dim wb As Workbook, ws As Worksheet, lo As ListObject
    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)
    With ws
        .Range("A1") = "Col1"
        .Range("B1") = "Col2"
        .Range("A2") = "testing"
        .Range("B2") = Now()
        Set lo = .ListObjects.Add(SourceType:=xlSrcRange, Source:=.Range("A1:B2"), XLListObjectHasHeaders:=xlYes)
    End With
    Dim i As Long, tARR() As Variant
    Dim tVal
    ReDim tARR(1 To 10000, 1 To 2)
    For i = LBound(tARR) To UBound(tARR)
        tVal = ""
        Dim tLen As Long
        tLen = CDbl(maxLen) * Rnd
        If tLen < 1 Then tLen = 1
        tLen = CLng(tLen)
        Dim j As Long
        For j = 1 To tLen
            tVal = Concat(tVal, GetLetter())
        Next j
        tARR(i, 1) = tVal
        tARR(i, 2) = Now()
    Next i
    pbListObj.NewRowsRange(lo, 10000).value = tARR
    
    sw.StartTimer
    pbRange.AddSort lo, 1, , True
    sw.StopTimer
    Debug.Print "Sorting took " & sw.Result
        
    sw.resetTimer
    sw.StartTimer
    pbRange.AddSort lo, 1, , True
    sw.StopTimer
    Debug.Print "Second Sorting took " & sw.Result
    
    

End Function
Private Function GetLetter()
    GetLetter = Chr(CLng(58 * Rnd) + 65)
End Function

Public Function upw()
    Dim ws As Worksheet
    For Each ws In nypa.Worksheets
        If ws.protectContents Then
            Debug.Print ws.Name & " is currently protected"
        End If
        'ws.EnablePivotTable = True
        If ws.EnablePivotTable = False Then
            Debug.Print "Enable Pivot Table = " & ws.EnablePivotTable & " on " & ws.Name
        End If
        Debug.Print "Enable AutoFilter = " & ws.EnableAutoFilter & " on " & ws.Name
        
    Next ws
End Function

Public Function nypa() As Workbook
    Set nypa = Workbooks("VMs Inventory 20230417.xlsx")
End Function

Public Sub CallTestSub1()
    Beep
    Debug.Print "CallTestSub1 called  in workbook: " & ThisWorkbook.Name
End Sub
Public Sub CallTestSub2(Optional input1)
    Debug.Print "CallTestSub2 called with 1 arg (" & IIf(IsMissing(input1), "[not inclulded]", input1) & ") in workbook: " & ThisWorkbook.Name
End Sub
Public Function CallTestFunction1() As Long
    Debug.Print "CallTestFunction1 called  in workbook: " & ThisWorkbook.Name
End Function
Public Function CallTestFunction2(Optional input1 As String) As String
    Debug.Print "CallTestFunction2 called with 1 arg (" & IIf(IsMissing(input1), "[not inclulded]", input1) & ") in workbook: " & ThisWorkbook.Name
    CallTestFunction2 = TypeName(input1)
End Function


'Public Function ReCodeName(wkbk As Workbook, currentCodeName, newCodeName)
'On Error Resume Next
'    If wkbk.HasVBProject = True Then
'        Dim vbComp As VBComponent
'        Set vbComp = wkbk.VBProject.VBComponents(currentCodeName)
'        If Not vbComp Is Nothing Then
'            vbComp.Properties("_CodeName") = CStr(newCodeName)
'            If StringsMatch(vbComp.Properties("_CodeName"), newCodeName) Then
'                Beep
'                Debug.Print "Successfully Change Code Name to: "; newCodeName
'            End If
'        End If
'    End If
'End Function

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
        textFileName = "TEXT_" & FileNameWithoutExtension(FileNameFromFullPath2(wkbk.FullNameURLEncoded)) & ".txt"
        desiredFilePath = CombinePath(FullPathExcludingFileName2(wkbk.FullNameURLEncoded), textFileName)
        
        '' save as txt to 'safe' directory
        Debug.Print "Saving to " & safeFilePath
        wkbk.SaveCopyAs safeFilePath
        
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
            fileName = Replace(fileName, tmpExt, vbNullString)
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
