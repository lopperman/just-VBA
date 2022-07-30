VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbHASH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   pbHASH v1.0.0
'   (c) Paul Brower - https://github.com/lopperman/just-VBA
'   u/ITFuture (Reddit)
'
'   Custom Hashlist implementation
'   (Based off a hashlist found here: https://www.vbforums.com/showthread.php?834515-Simple-and-fast-lightweight-HashList-Class-(no-APIs))
'
'
' @CLASS pbHASH
' @author Paul Brower
'   ' Added additional functions for values that are objects;
'   ' Added options for aggregating values while creating indexes
'   ' Added 'not fail' option for duplicate indexes with matching values
'   ' Effort in progress for dyanmicall adjusting buckets for performance
' @license GNU General Public License v3.0
'
'   * Defaults to vbTextCompare
'   * Use 'SetExpectedCount' For Additional Optimizations
'       (Default = 5000)

    Option Explicit

    Private Type KVP
      key As String
      value As Variant
    End Type
    Private Type HashIndexes
      DataIndexes() As Long
    End Type
    Private DataTable() As KVP
    Private HashTable() As HashIndexes
    Private lCount As Long, lDtUpper As Long, lHashTblSize As Long
    Public CompareMode As VbCompareMethod, UniqueKeys As Boolean

    Private Sub Class_Initialize()
      UniqueKeys = True
      CompareMode = vbTextCompare
      SetExpectedCount
    End Sub

    Public Function CreateOrUpdate_SUM(ByVal fKey As String, ByVal dVal As Variant)
        If Exists(fKey) Then
           If Not dVal = 0 Then Item(fKey) = Item(fKey) + dVal
        Else
            'Add new Key, even if value is 0
            Add dVal, key:=fKey
        End If
    End Function

    Public Function CreateOrUpdate_COUNT(ByVal fKey As String, Optional ByVal addCount As Long = 1)
        If Exists(fKey) Then
            If addCount <> 0 Then Item(fKey) = Item(fKey) + addCount
        Else
            Add addCount, key:=fKey
        End If
    End Function


    Public Function SetExpectedCount(Optional ByVal expCount As Long = 5000)
        lHashTblSize = 8
        Do Until lHashTblSize * 5 > expCount
            lHashTblSize = lHashTblSize * 2
        Loop
        ReDim HashTable(0 To lHashTblSize - 1)
        Dim i As Long
        For i = 0 To UBound(HashTable)
            ReDim HashTable(i).DataIndexes(0 To 0)
        Next i
        lDtUpper = 16
        ReDim DataTable(0 To lDtUpper)
        lCount = 0
    End Function

    Public Property Get count() As Long
        count = lCount
    End Property
    
    Public Function Exists(key As String) As Boolean
        Exists = FindIndex(key, CalcHash(key)) > 0
    End Function
    
    Public Function IndexByKey(key As String) As Long
        IndexByKey = FindIndex(key, CalcHash(key))
    End Function

    Public Function Add(Item, Optional key As String)
        Dim HashValue As Long, UB As Long
        HashValue = CalcHash(key)
        If UniqueKeys Then
            If FindIndex(key, HashValue) Then
                Err.Raise 457, Description:="Key already exists in pbHASH (UniqueKeys: TRUE, key: " & key & ")"
            End If
        End If
        
    '   prolong and add to the new entries to the DataTable-Array
        lCount = lCount + 1
        If lDtUpper < lCount Then
            lDtUpper = lDtUpper * 1.5
            ReDim Preserve DataTable(0 To lDtUpper)
        End If
      
        DataTable(lCount).key = key
        If IsObject(Item) Then
            Set DataTable(lCount).value = Item
        Else
            DataTable(lCount).value = Item
        End If
      
    '   add the new DataIndex to the proper Hash-Buckets .DataIndexes-Array
        With HashTable(HashValue)
            UB = UBound(.DataIndexes)
            UB = UB + 1
            ReDim Preserve .DataIndexes(0 To UB)
            .DataIndexes(UB) = lCount
        End With
    End Function

    Public Property Get KeyByIndex(ByVal IndexOneBased As Long)
        If IndexOneBased < 1 Or IndexOneBased > lCount Then Err.Raise 9
        KeyByIndex = DataTable(IndexOneBased).key
    End Property

    Public Property Get Item(KeyOrOneBasedIndex)
        Dim Index As Long
        If VarType(KeyOrOneBasedIndex) = vbString Then
            Index = FindIndex(KeyOrOneBasedIndex, CalcHash(KeyOrOneBasedIndex))
            If Index = 0 Then Err.Raise 457
        Else
            Index = KeyOrOneBasedIndex
            If Index < 1 Or Index > lCount Then Err.Raise 9
        End If
        
        If IsObject(DataTable(Index).value) Then
            Set Item = DataTable(Index).value
        Else
            Item = DataTable(Index).value
        End If
    End Property

    Public Property Let Item(KeyOrOneBasedIndex, newItem)
        Dim Index As Long
        If VarType(KeyOrOneBasedIndex) = vbString Then
            Index = FindIndex(KeyOrOneBasedIndex, CalcHash(KeyOrOneBasedIndex))
            If Index = 0 Then Err.Raise 457
        Else
            Index = KeyOrOneBasedIndex
            If Index < 1 Or Index > lCount Then Err.Raise 9
        End If
        
        If IsObject(newItem) Then
            Set DataTable(Index).value = newItem
        Else
            DataTable(Index).value = newItem
        End If
    End Property

    Public Property Set Item(KeyOrOneBasedIndex, newItem)
        Item(KeyOrOneBasedIndex) = newItem
    End Property

    Private Function FindIndex(key, ByVal HashValue As Long) As Long
        Dim i As Long, CM As VbCompareMethod
        With HashTable(HashValue)
            CM = CompareMode
            For i = 1 To UBound(.DataIndexes)
                If StrComp(key, DataTable(.DataIndexes(i)).key, CM) = 0 Then
                    FindIndex = .DataIndexes(i): Exit Function
                End If
            Next i
        End With
    End Function

    Public Function CalcHash(key) As Long
        Dim i As Long, l As Long, b() As Byte
        If CompareMode Then
            b = LCase$(key)
        Else
            b = key
        End If
        l = 7919
        For i = UBound(b) To 0 Step -1
            l = (i + b(i) + l) * 37 And &H7FFFFF
        Next i
        CalcHash = l * b(0) Mod lHashTblSize
    End Function

Friend Function CheckHashDistribution() As String
    Dim i As Long, UB As Long, cc As Long, Min As Long, Max As Long
    Min = &H7FFFFFFF
    For i = 0 To UBound(HashTable)
        UB = UBound(HashTable(i).DataIndexes)
        If UB Then
            If Min > UB Then Min = UB
            If Max < UB Then Max = UB
            cc = cc + 1
        End If
    Next i
  
    Dim retV As String
    retV = Concat("Distribution over a HashTable with ", UBound(HashTable) + 1, "slots:")
    retV = Concat(retV, vbNewLine, "Used-HashSlots: ", cc)
    retV = Concat(retV, vbNewLine, "Min-Entries: ", Min)
    retV = Concat(retV, vbNewLine, "Max-Entries: ", Max)
    
    CheckHashDistribution = retV
    
    Debug.Print "Distribution over a HashTable with"; UBound(HashTable) + 1; "slots:"
    Debug.Print "Used-HashSlots:"; cc
    Debug.Print "Min-Entries:"; Min
    Debug.Print "Max-Entries:"; Max

End Function
