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
'   (Concept based off https://www.vbforums.com/showthread.php?834515-Simple-and-fast-lightweight-HashList-Class-(no-APIs))
'
' @CLASS pbHASH
' @author Paul Brower
'   ' Added additional functions for values that are objects;
'   ' Added options for aggregating values while creating indexes
'   ' Added 'not fail' option for duplicate indexes with matching values
'   '
'   ' Effort in progress for dyanmicall adjusting buckets for performance
' @license GNU General Public License v3.0
'
'   * Defaults to vbTextCompare
'   * Use 'SetExpectedCount' For Additional Optimizations
'       (Default = 5000)
    Option Explicit
    Option Compare Text

'   ~~~ ADD KVP TYPE TO PUBLIC MODULE ~~~
'    Public Type KVP
'      KEY As String
'      value As Variant
'    End Type
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

    '   Create or Increment Item Value by 1 or [addCount]
    Public Function CreateOrUpdate_COUNT(ByVal fKey As String, Optional ByVal addCount As Long = 1)
        If Exists(fKey) Then
            If addCount <> 0 Then Item(fKey) = Item(fKey) + addCount
        Else
            Add addCount, key:=fKey
        End If
    End Function

    '   Create if missing, otherwise update if [dVal] > current hashed key item value
    Public Function CreateOrUseMAX(ByVal fKey As String, ByVal dVal As Variant)
        If Exists(fKey) Then
            Item(fKey) = ChooseMax(dVal, Item(fKey))
        Else
            Add dVal, key:=fKey
        End If
    End Function
    
    '   Create if missing, otherwise update if [dVal] < current hashed key item value
    Public Function CreateOrUseMIN(ByVal fKey As String, ByVal dVal As Variant)
        If Exists(fKey) Then
            Item(fKey) = ChooseMin(dVal, Item(fKey))
        Else
            Add dVal, key:=fKey
        End If
    End Function
    
    '   return which is greater - supports any intrinsic data type
    Private Function ChooseMax(Val1, Val2)
        If Val1 > Val2 Then
            ChooseMax = Val1
        Else
            ChooseMax = Val2
        End If
    End Function
    '   return which is lesser - supports any intrinsic data type
    Public Function ChooseMin(ByVal Val1, ByVal Val2)
        If Val1 > Val2 Then
            ChooseMin = Val2
        Else
            ChooseMin = Val1
        End If
    End Function
    
    '   improves performance by configuring hash buckets early
    Public Function SetExpectedCount(Optional ByVal expCount As Long = 5000)
'        using next prime doesn't seem to help
'        lHashTblSize = NextPrime(CLng(expCount * 1.3))
'
'
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

    Public Property Get Count() As Long
        Count = lCount
    End Property
    
    Public Function Exists(ByVal key As String) As Boolean
        Exists = FindIndex(key, CalcHash(key)) > 0
    End Function
    
    Public Function IndexByKey(ByVal key As String) As Long
        IndexByKey = FindIndex(key, CalcHash(key))
    End Function

    Public Function AddKV(ByVal key As String, Item, Optional alsoAddReverse As Boolean = False)
                 Add Item, key
                If alsoAddReverse Then
                    If IsObject(Item) = False Then
                        AddKV CStr(Item), key
                    End If
                End If
    End Function

    Public Function Add(Item, Optional ByVal key As String)
        Dim HashValue As Long, ub As Long
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
            Set DataTable(lCount).Value = Item
        Else
            DataTable(lCount).Value = Item
        End If
      
    '   add the new DataIndex to the proper Hash-Buckets .DataIndexes-Array
        With HashTable(HashValue)
            ub = UBound(.DataIndexes)
            ub = ub + 1
            ReDim Preserve .DataIndexes(0 To ub)
            .DataIndexes(ub) = lCount
        End With
    End Function

    Public Property Get KeyByIndex(ByVal IndexOneBased As Long)
        If IndexOneBased < 1 Or IndexOneBased > lCount Then Err.Raise 9
        KeyByIndex = DataTable(IndexOneBased).key
    End Property

    Public Function IsItemObject(ByVal KeyOrOneBasedIndex) As Boolean
        Dim index As Long
        If VarType(KeyOrOneBasedIndex) = vbString Then
            index = FindIndex(KeyOrOneBasedIndex, CalcHash(KeyOrOneBasedIndex))
            If index = 0 Then Err.Raise 457
        Else
            index = KeyOrOneBasedIndex
            If index < 1 Or index > lCount Then Err.Raise 9
        End If
        
        If IsObject(DataTable(index).Value) Then
            IsItemObject = True
        Else
            IsItemObject = False
        End If
    End Function

    Public Property Get Item(ByVal KeyOrOneBasedIndex)
        Dim index As Long
        If VarType(KeyOrOneBasedIndex) = vbString Then
            index = FindIndex(KeyOrOneBasedIndex, CalcHash(KeyOrOneBasedIndex))
            If index = 0 Then Err.Raise 457
        Else
            index = KeyOrOneBasedIndex
            If index < 1 Or index > lCount Then Err.Raise 9
        End If
        
        If IsObject(DataTable(index).Value) Then
            Set Item = DataTable(index).Value
        Else
            Item = DataTable(index).Value
        End If
    End Property

    Public Property Get Items() As KVP()
        Items = DataTable
    End Property

    Public Property Let Item(ByVal KeyOrOneBasedIndex, newItem)
        Dim index As Long
        If VarType(KeyOrOneBasedIndex) = vbString Then
            index = FindIndex(KeyOrOneBasedIndex, CalcHash(KeyOrOneBasedIndex))
            If index = 0 Then Err.Raise 457
        Else
            index = KeyOrOneBasedIndex
            If index < 1 Or index > lCount Then Err.Raise 9
        End If
        
        If IsObject(newItem) Then
            Set DataTable(index).Value = newItem
        Else
            DataTable(index).Value = newItem
        End If
    End Property

    Public Property Set Item(ByVal KeyOrOneBasedIndex, newItem)
        Item(KeyOrOneBasedIndex) = newItem
    End Property

    Private Function FindIndex(ByVal key, ByVal HashValue As Long) As Long
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

    Public Function CalcHash(ByVal key) As Long
        Dim i As Long, L As Long, B() As Byte
        If CompareMode Then
            B = LCase$(key)
        Else
            B = key
        End If
        L = 7919
        For i = UBound(B) To 0 Step -1
            L = (i + B(i) + L) * 37 And &H7FFFFF
        Next i
        CalcHash = L * B(0) Mod lHashTblSize
    End Function


Public Function NextPrime(ByVal startingNumber As Long) As Long
    Dim number As Long
    Dim divisor As Long
    Dim maxdivisor As Long
    Dim isprime As Boolean

    number = startingNumber

    Do While isprime = False
        Select Case number
            Case Is < 2
                isprime = False
            Case Is = 2
                isprime = True
            Case Is > 2
                isprime = True
                If number Mod 2 = 0 Then
                    isprime = False
                Else
                    maxdivisor = CLng(Sqr(number))
                    divisor = 3
                    Do While (divisor <= maxdivisor)
                        If (number Mod divisor = 0) Then
                           isprime = False
                           Exit Do
                        End If
                        divisor = divisor + 1
                    Loop
                End If
        End Select
        If (isprime = True) Then
            NextPrime = number
            Exit Function
        Else
            number = number + 1
        End If
    Loop
End Function


Friend Function CheckHashDistribution() As String
    Dim i As Long, ub As Long, cc As Long, Min As Long, Max As Long
    Min = &H7FFFFFFF
    For i = 0 To UBound(HashTable)
        ub = UBound(HashTable(i).DataIndexes)
        If ub Then
            If Min > ub Then Min = ub
            If Max < ub Then Max = ub
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

Private Function Concat(ParamArray Items() As Variant) As String
    Concat = Join(Items, "")
End Function

