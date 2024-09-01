VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbUpdateLO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Find and Update Utility for List Objects
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      author (c) Paul Brower https://github.com/lopperman/just-VBA
''      class module pbUpdateLO.cls
''      license GNU General Public License v3.0
''      Created Aug-2024
''
''      Download pbUpdateLO.cls with NO DEPENDENCIES at
''      https://github.com/lopperman/just-VBA/blob/main/Code_NoDependencies/pbUpdateLO.cls
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1

    Private Enum strMatchEnum
        smEqual = 0
        smNotEqualTo = 1
        smContains = 2
        smStartsWithStr = 3
        smEndWithStr = 4
    End Enum

    Public Enum SearchTypeOperator
        stIncludeMatched = 1
        stExcludeMatched = 2
    End Enum
    
    Public Enum SearchResultsFormat
        ArrayMatchedIndexesAllColumns
        ArrayMatchedIndexes
        ArrayMatchedRows
        ArrayUnmatchedIndexes
        ArrayUnmatchedIndexesAllColumns
        ArrayUnmatchedRows
        CollectionMatchedIndexes
        CollectionUnmatchedIndexes
        AllListIndexesWithTrueFalse
    End Enum

    Public Enum ListObjUpdateOptions
        sheetProtectionPWD = 1
        disableConfirmation = 2
        failIfRowCountNotEqualTo = 3
        failIfRowCountGreaterThan = 4
        failIfRowCountLessThan = 5
        ignoreRaiseEvents = 6
    End Enum

Private l_options As Collection
Private l_newSearch  As Boolean
Private l_searchListObject As ListObject
Private l_IndexResults() As Variant
Private l_arr
Private l_matchCount As Long
Private l_logMessages As Collection

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      ROADMAP
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  TODO: Enable Update Operation to be stopped after any given row update


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      EVENTS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Raises Event that can be used for logging
''      [logLevel] will be one of the following:
''              "TRACE", "ERROR"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
        Public Event OnLogMessage(logLevel As String, msg As String)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Raised before changes are made
''      If handled, can be canceled by setting Cancel = True
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
        Public Event BeforeUpdate(updateCount As Long, pendingChanges As Collection, ByRef Cancel As Boolean)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Raised after changes are made
''      Use for logging / reporting
''      Raised for each row that has been updated
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
        Public Event ListRowChanged(ByVal rowIndex As Long, searchCols() As Variant, updateCols() As Variant)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Raised after error occurs
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Event OnError(ByVal errNumber, ByVal errDescription)


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      PUBLIC METHODS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''       Returns all log messages as collection of strings
''      (Does not require 'OnLogMessage' events to be handled)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Get LogMessages() As Collection
    Set LogMessages = l_logMessages
End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Adds a configuration option
''  Options must be set before calling 'NewSearch'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function AddOption(ByVal updateOption As ListObjUpdateOptions, ByVal optionValue) As pbUpdateLO
    If l_newSearch Then
        OnErrorRaise 1004, "pbUpdateLO.AddOption - Options cannot be added after 'NewSearch' has been called"
    End If
    OnLog "TRACE", "AddOption (" & updateOption & ") = " & optionValue
    If CollectionKeyExists(l_options, CStr(updateOption)) Then
        l_options.Remove CStr(updateOption)
    End If
    l_options.Add optionValue, key:=CStr(updateOption)
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Gets value for option [updateOption], if it has been set
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function GetOption(ByVal updateOption As ListObjUpdateOptions) As Variant
    If CollectionKeyExists(l_options, CStr(updateOption)) Then
        GetOption = l_options(CStr(updateOption))
    Else
        OnLog "TRACE", "pbUpdateLO.GetOption - Cannot get option [" & updateOption & "] because it does not exist"
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Return current number of matched rows from searched ListObject
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Get MatchedCount() As Long
    MatchedCount = l_matchCount
End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Return current number of unmatched rows from searched ListObject
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Get UnmatchedCount() As Long
    UnmatchedCount = UBound(l_IndexResults) - MatchedCount
End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns True if [listRowIdx] of searched ListObject was matched at this
''  point of search
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function WasMatched(listRowIdx As Long) As Boolean
    WasMatched = l_IndexResults(listRowIdx, 2)
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Get Results(toFormat As SearchResultsFormat)
''  (Can be called as many times as needed after search)
''
''  Returns matched or unmatched information, depending on value of
''  [toFormat] argument
''
''  [toFormat] = AllListIndexesWithTrueFalse
''      Returns array (1 to N, 1 to 2) that contains ** all ** listrow indexes and
''      True/False (True if matched search criteria)
''      e.g.
''          Dim r: r=[instance].Results(AllListIndexesWithTrueFalse)
''          Dim row1Matched: row1Matched = r(1,2)
''          if row1Matched Then ... (ListRow 1 matched search)
''
''  [toFormat] = ArrayMatchedIndexesAllColumns
''       Returns array (1 to N,0 to X) where N = number of matched rows
''      and X = number of columns in searched list object
''      Note: array(1,0) contains the row index of matched row, while
''          array(1,[1 to X]) contains the column value from the searched list object
''          for the row index that was matched
''          (this saves a step if you need todo something with ListRow values)
''
''  [toFormat] = ArrayMatchedRows
''      Returns array (1 to N,1 to X) where N = number of matched rows
''      and X = number of columns in searched list object
''      Similar to [ArrayMatchedIndexesAllColumns], however the row
''      index is omitted.
''
''  [toFormat] = ArrayMatchedIndexes
''      Returns array (1 to N, 1 to 1) of matched listobject listrow indexes
''      where N = number of matched rows.
''
''  [toFormat] = ArrayUnmatchedIndexesAllColumns
''       Returns array (1 to N,0 to X) where N = number of ** unmatched ** rows
''      and X = number of columns in searched list object
''      Note: array(1,0) contains the row index of unmatched row, while
''          array(1,[1 to X]) contains the column value from the searched list object
''          for the row index that was not matched
''          (this saves a step if you need todo something with ListRow values)
''
''  [toFormat] = ArrayUnmatchedIndexes
''      Returns array (1 to N, 1 to 1) of ** unmatched ** listobject listrow indexes
''      where N = number of unmatched rows.
''
''  [toFormat] = CollectionMatchedIndexes
''      Returns Collection containing matched row indexes
''
''  [toFormat] = CollectionUnmatchedIndexes
''      Returns Collection containing ** unmatched ** row indexes
''
''  [toFormat] = ArrayUnmatchedRows
''      Returns array (1 to N,1 to X) where N = number of ** unmatched  **rows
''      and X = number of columns in searched list object
''      Similar to [ArrayUnmatchedIndexesAllColumns], however the row
''      index is omitted.

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Get Results(toFormat As SearchResultsFormat)
    Select Case toFormat
        Case SearchResultsFormat.AllListIndexesWithTrueFalse
            Results = l_IndexResults
        Case SearchResultsFormat.ArrayMatchedIndexesAllColumns
            Results = BuildResultsArray(True, True)
        Case SearchResultsFormat.ArrayMatchedIndexes
            Results = BuildResultsArray(True, False)
        Case SearchResultsFormat.ArrayUnmatchedIndexesAllColumns
            Results = BuildResultsArray(False, True)
        Case SearchResultsFormat.ArrayUnmatchedIndexes
            Results = BuildResultsArray(False, False)
        Case SearchResultsFormat.ArrayMatchedRows
            Results = BuildResultsArray(True, True, True)
        Case SearchResultsFormat.ArrayUnmatchedRows
            Results = BuildResultsArray(False, True, True)
        Case SearchResultsFormat.CollectionMatchedIndexes
            Set Results = BuildResultsCollection(True)
        Case SearchResultsFormat.CollectionUnmatchedIndexes
            Set Results = BuildResultsCollection(False)
    End Select
End Property


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      PRIVATE METHODS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Builds Collection of row indexes that were matched ([matchedVal] = true)
''      or unmatched ([matchedVal]=false)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function BuildResultsCollection(matchedVal As Boolean) As Collection
    Dim resp As New Collection
    Dim i As Long
    For i = 1 To UBound(l_IndexResults)
        If l_IndexResults(i, 2) = matchedVal Then
            resp.Add (i)
        End If
    Next i
    Set BuildResultsCollection = resp
    Set resp = Nothing
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Builds array of matched/unmatchefd list row indexes
''      if [allColumns]=True, will also includes listrow values for all columns
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function BuildResultsArray(addIfMatched As Boolean, allColumns As Boolean, Optional excludeRowIndex As Boolean = False) As Variant()
    OnLog "TRACE", "BuildResultsArray AddIfMatched = " & addIfMatched & ", AllColumns = " & allColumns
    
    Dim arrDim1 As Long, arrDim2 As Long
    Dim i As Long, c As Long
    If addIfMatched Then
        arrDim1 = l_matchCount
    Else
        arrDim1 = UBound(l_IndexResults) - l_matchCount
    End If
    If allColumns Then
        arrDim2 = UBound(l_arr, 2)
    Else
        arrDim2 = 1
    End If
    Dim resp As Variant
    Dim respIdx As Long
    If allColumns And excludeRowIndex = False Then
        ''Second Dimension starts at 0 and includes row index
        ReDim resp(1 To arrDim1, 0 To arrDim2)
    Else
        ReDim resp(1 To arrDim1, 1 To arrDim2)
    End If
    For i = LBound(l_IndexResults) To UBound(l_IndexResults)
        If l_IndexResults(i, 2) = addIfMatched Then
            respIdx = respIdx + 1
            If allColumns = False Then
                resp(respIdx, 1) = i
            Else
                If excludeRowIndex = False Then
                    resp(respIdx, 0) = i
                End If
                For c = 1 To arrDim2
                    resp(respIdx, c) = l_arr(i, c)
                Next c
            End If
        End If
    Next i
    BuildResultsArray = resp
    
End Function

Public Function NewSearch( _
    listObj As ListObject, _
    listColumn, _
    matchLogic As SearchTypeOperator, _
    searchVal1, _
    Optional searchVal2, _
    Optional searchOperator As XlFormatConditionOperator = XlFormatConditionOperator.xlEqual) As pbUpdateLO
        
    Dim initialMatches As Collection
    
    If l_newSearch Then
        RaiseEvent OnError(1004, "pbUpdateLO.NewSearch has already been defined. To create another search, first create a new instance of this class")
        Err.Raise 1004, Description:="pbUpdateLO.NewSearch has already been defined. To create another search, first create a new instance of this class"
    End If
    If listObj.listRows.Count = 0 Then
        RaiseEvent OnError(1004, "pbUpdateLO.NewSearch Cannot Search Empty List Object")
        Err.Raise 1004, Description:="pbUpdateLO.NewSearch Cannot Search Empty List Object"
    End If
    If searchOperator = xlBetween Or searchOperator = xlNotBetween Then
        If IsMissing(searchVal2) Then
            RaiseEvent OnError(1004, "pbUpdateLO.NewSearch: [searchVal2] is missing - required when [searchOperator] = xlBetween or xlNotBetween ")
            Err.Raise 1004, Description:="pbUpdateLO.NewSearch: [searchVal2] is missing - required when [searchOperator] = xlBetween or xlNotBetween "
        End If
    End If
    
    l_newSearch = True
    Set l_searchListObject = listObj
    l_arr = listObj.DataBodyRange.value
    '' If initial search [matchLogic] is to find things to exclude, all row indexes start out as 'matched',
    '' otherwise all row indexes start out as 'unmatched'
    CreateResultsArr IIf(matchLogic = stIncludeMatched, False, True)
    
    '' whether performing inclusive or exclusive searching, get matches ([MatchedIndexes]),
    ''  and then use that list to turn on (inclusive) or off (exlusive) matched listrow indexes
    Set initialMatches = MatchedIndexes(listColumn, searchVal1, searchVal2, searchOperator)
    '' defines initial search results
    UpdateResults initialMatches, matchLogic
    Set NewSearch = Me
End Function


Private Function MatchedIndexes(listColumn, _
    searchVal1, _
    Optional searchVal2, _
    Optional searchOperator As XlFormatConditionOperator = XlFormatConditionOperator.xlEqual) As Collection
    
    Dim resp As New Collection
    Dim colIdx As Long
    colIdx = l_searchListObject.ListColumns(listColumn).index
    
    Dim i As Long, curValue As Variant, isMatched As Boolean
    For i = 1 To UBound(l_IndexResults)
        curValue = l_arr(i, colIdx)
        Select Case searchOperator
            Case XlFormatConditionOperator.xlEqual
                isMatched = StringsMatch(curValue, searchVal1) Or (curValue = searchVal1)
            Case XlFormatConditionOperator.xlNotEqual
                isMatched = Not StringsMatch(curValue, searchVal1) Or Not (curValue = searchVal1)
            Case XlFormatConditionOperator.xlGreater
                isMatched = curValue > searchVal1
            Case XlFormatConditionOperator.xlGreaterEqual
                isMatched = curValue >= searchVal1
            Case XlFormatConditionOperator.xlLess
                isMatched = curValue < searchVal1
            Case XlFormatConditionOperator.xlLessEqual
                isMatched = curValue <= searchVal1
            Case XlFormatConditionOperator.xlBetween
                isMatched = (curValue >= MinV(searchVal1, searchVal2)) And (curValue <= MaxV(searchVal1, searchVal2))
            Case XlFormatConditionOperator.xlNotBetween
                isMatched = (curValue < MinV(searchVal1, searchVal2)) Or (curValue > MaxV(searchVal1, searchVal2))
        End Select
        If isMatched Then
            resp.Add i
        End If
    Next i
    Set MatchedIndexes = resp
    Set resp = Nothing
End Function

Private Function UpdateResults(indexesCol As Collection, matchLogic As SearchTypeOperator)
    LogTRACE "pbUpdateLO.UpdateResults Start Match Count = " & l_matchCount
    If indexesCol.Count > 0 Then
        Dim tIdx As Variant, curFlag As Boolean, includeFlag As Boolean
        includeFlag = (matchLogic = stIncludeMatched)
        For Each tIdx In indexesCol
            curFlag = l_IndexResults(tIdx, 2)
            If Not curFlag = includeFlag Then
                If includeFlag Then
                    l_matchCount = l_matchCount + 1
                Else
                    l_matchCount = l_matchCount - 1
                End If
                l_IndexResults(tIdx, 2) = includeFlag
            End If
        Next tIdx
    End If
    LogTRACE "pbUpdateLO.UpdateResults End Match Count = " & l_matchCount
End Function

Private Function CreateResultsArr(defaultToMatched As Boolean)
    Dim i As Long
    ReDim l_IndexResults(1 To l_searchListObject.listRows.Count, 1 To 2)
    For i = 1 To l_searchListObject.listRows.Count
        l_IndexResults(i, 1) = i
        l_IndexResults(i, 2) = defaultToMatched
    Next i
    If defaultToMatched Then
        l_matchCount = UBound(l_IndexResults)
    Else
        l_matchCount = 0
    End If
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Returns item from collection by Key
''   If [key] does not exist in collection, error object with
''   error code 1004 is return
''   suggested use:
''
''   Dim colItem as Variant
''   colItem = CollectionItemByKey([collection], [expectedKey])
''
''   'If expecting object, use 'Set'
''    Set colItem = CollectionItemByKey([collection], [expectedKey])
''
''   If Not IsError(colItem) Then
''       'value was returned
''   Else
''       'error was returned
''   End if
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function CollectionItemByKey(ByRef col As Collection, ByVal key)
On Error Resume Next
    key = CStr(key)
    If IsObject(col(key)) Then
        If Err.number = 0 Then
            Set CollectionItemByKey = col(key)
        End If
    Else
        If Err.number = 0 Then
            CollectionItemByKey = col(key)
        End If
    End If
    If Err.number <> 0 Then
        Err.Clear
        CollectionItemByKey = CVErr(1004)
    End If
End Function
Private Function CollectionKeyExists(ByRef col As Collection, ByVal key)
On Error Resume Next
    key = CStr(key)
    If IsObject(col(key)) Then
        If Err.number = 0 Then
            CollectionKeyExists = True
        Else
            CollectionKeyExists = False
        End If
    Else
        If Err.number = 0 Then
            CollectionKeyExists = True
        Else
            CollectionKeyExists = False
        End If
    End If
    If Err.number <> 0 Then
        Err.Clear
    End If
End Function

    Private Function StringsMatch( _
        ByVal checkString As Variant, ByVal _
        validString As Variant, _
        Optional smEnum As strMatchEnum = strMatchEnum.smEqual, _
        Optional compMethod As VbCompareMethod = vbTextCompare) As Boolean
         
    '       IF NEEDED, PUT THIS ENUM AT TOP OF A STANDARD MODULE
            'Public Enum strMatchEnum
            '    smEqual = 0
            '    smNotEqualTo = 1
            '    smContains = 2
            '    smStartsWithStr = 3
            '    smEndWithStr = 4
            'End Enum
            
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

    Private Function MaxV(Val1, Val2)
        If Val1 > Val2 Then
            MaxV = Val1
        Else
            MaxV = Val2
        End If
    End Function
    
    Public Function MinV(Val1, Val2)
        If Val1 > Val2 Then
            MinV = Val2
        Else
            MinV = Val1
        End If
    End Function

Private Function TimeStamp() As String
    TimeStamp = Format(Now, "yyyymmdd hh:mm:ss") & Right(Format(Timer, "0.000"), 4)
End Function

Private Function OnErrorRaise(errNumber, errMessage As String)
    OnLog "ERROR", errMessage
    RaiseEvent OnError(errNumber, errMessage)
    Err.Raise errNumber, Source:=Me, Description:=errMessage
End Function

Private Function OnLog(logLvl As String, msg As String)
    l_logMessages.Add TimeStamp & " " & UCase(logLvl) & " " & msg
    RaiseEvent OnLogMessage(logLvl, msg)
End Function



Private Sub Class_Initialize()
    Set l_logMessages = New Collection
    Set l_options = New Collection
    l_newSearch = False
    l_matchCount = 0
End Sub
