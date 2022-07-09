Attribute VB_Name = "mdlFindLstObjRow"
Option Explicit
Option Compare Text
Option Base 1


'   ~~~ ~~~ Very Fast Function to find the first row of a ListObject where EXACT MATCH filters can be applied for up to ALL the columns in the list object
'   ~~~ ~~~ Recommend at a minimum the List Object be sorted Ascending by the First Column in the [Columns] Array
'   ~~~ ~~~
'   ~~~ ~~~ Arguments:
'   ~~~ ~~~ [lstObj] = Reference to list object being searched
'   ~~~ ~~~ [Columns] = An Array of ListColumn Names or ListColumn Indexes
'   ~~~ ~~~ [Crit] = An Array of Search Criteria - 1 criteria for each ListObject Column in the [Columns] Array
'   ~~~ ~~~ Example: lstObjRowIndex = FindFirstListObjectRow([myListObject], Array("LastName","DOB"), Array("Smith",CDate("12/28/80"))
Public Function FindFirstListObjectRow(lstObj As ListObject, Columns As Variant, crit As Variant) As Long
On Error GoTo E:
    Dim failed As Boolean
    
    '   If no rows, no play
    If lstObj.ListRows.Count = 0 Then Exit Function
    
    '   Get reference to worksheet. Cutting out the 'middle man (list object)' for the Range.Find calls, to save even a few clock ticks
    Dim ws As Worksheet
    Set ws = lstObj.Range.Worksheet
    
    Dim matchedListObjIdx As Long, matchedWSRow As Long
    Dim firstRow As Long, lastRow As Long, rowOffset As Long, colOffset As Long
    
    '   get worksheet first/last row for possible ListObject range that can be searched
    firstRow = lstObj.ListRows(1).Range.Row
    lastRow = lstObj.HeaderRowRange.Row + lstObj.ListRows.Count
    
    '   Since we're searching columns based on ListObject Column Index, and since were returning the ListObject RowIndex if found,
    '   get the offset of the ListObject to the Worksheet, so we can search the right worksheet columns, and return the right ListObject row
     colOffset = lstObj.ListColumns(1).Range.Column - 1
     rowOffset = lstObj.ListRows(1).Range.Row - 1
    
    '   this reformats the search criteria so it can find results based on the Range.NumberFormat of the list columns.
    '   you may want to tweak for your own purposes as this will allow you to find, for example "$100.50" even though the actual value
    '   might be something like 100.5012
    Dim critIdx As Long
    For critIdx = LBound(crit) To UBound(crit)
        If TypeName(crit(critIdx)) <> "Boolean" And lstObj.ListColumns(Columns(critIdx)).DataBodyRange(1, 1).NumberFormat <> "General" Then
            crit(critIdx) = Format(crit(critIdx), lstObj.ListColumns(Columns(critIdx)).DataBodyRange(1, 1).NumberFormat)
        End If
    Next critIdx
    
    Dim startLooking As Range
    Dim lastCheckedRow As Long, colsArrIDX As Long, matched As Boolean
    Dim evalCol As Long, evalRow As Long
    Dim arrLB As Long, arrUB As Long
    arrLB = LBound(Columns)
    arrUB = UBound(Columns)
    
    '   Search for the First matched filter for the first Column
    Set startLooking = lstObj.ListColumns(Columns(arrLB)).Range.Find(crit(arrLB), LookIn:=xlValues, Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If startLooking Is Nothing Then
        GoTo Finalize:
    End If
    '   If only 1 column is being search, we're done
    If arrUB - arrLB = 0 Then
        matchedWSRow = startLooking.Row
        GoTo Finalize:
    Else
        '   Matched first column, but need to match addtional columns
        matchedWSRow = startLooking.Row
        lastCheckedRow = startLooking.Row
    End If
    
    '   Go through the remaining columns (we already matched the first) to see if the row matches
    Do While lastCheckedRow <= lastRow
        For colsArrIDX = arrLB + 1 To arrUB
            evalCol = lstObj.ListColumns(Columns(colsArrIDX)).Index + colOffset
            evalRow = lastCheckedRow
            If TypeName(crit(colsArrIDX)) = "String" Then
                matched = StrComp(ws.Cells(evalRow, evalCol).Text, CStr(crit(colsArrIDX)), vbTextCompare) = 0
            Else
                If TypeName(crit(colsArrIDX)) = "Date" Then
                    matched = CLng(ws.Cells(evalRow, evalCol).Value) = CLng(crit(colsArrIDX))
                Else
                    If IsNumeric(crit(colsArrIDX)) Then
                        matched = CDbl(ws.Cells(evalRow, evalCol).Value) = CDbl(crit(colsArrIDX))
                    Else
                        matched = ws.Cells(evalRow, evalCol) = crit(colsArrIDX)
                    End If
                End If
            End If
            If colsArrIDX = arrUB And matched Then
                    matchedWSRow = lastCheckedRow
                '   positive row match
                Exit Do
            ElseIf matched = False Then
                Set startLooking = lstObj.ListColumns(Columns(arrLB)).Range.FindNext(startLooking)
                If startLooking Is Nothing Then
                    matchedWSRow = 0
                    Exit Do
                ElseIf startLooking.Row <= lastCheckedRow Then
                    matchedWSRow = 0
                    Exit Do
                Else
                    lastCheckedRow = startLooking.Row
                End If
            End If
        Next colsArrIDX
    Loop
    
        
Finalize:
    On Error Resume Next
    Set ws = Nothing
    Set startLooking = Nothing
    If failed Then
        matchedListObjIdx = 0
    ElseIf matchedWSRow > 0 Then
        '   Adjust result to reflect the ListObjectRowIndex
        matchedListObjIdx = matchedWSRow - rowOffset
    End If
    FindFirstListObjectRow = matchedListObjIdx

    Exit Function
E:
    failed = True
    MsgBox "(Implement your own error handling) An error occured in FindFirstListRowMultCriteria: " & Err.Number & ", " & Err.Description
    Resume Finalize:

End Function
