VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHashList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type DataTableEntry
  key As String
  value As Variant
End Type
Private Type HashTableEntry
  DataIndexes() As Long
End Type
 
Private DataTable() As DataTableEntry, HashTable() As HashTableEntry
Private mCount As Long, mDTUBound As Long, mHashTableSize As Long
 
Public CompareMode As VbCompareMethod, UniqueKeys As Boolean

Private Sub Class_Initialize()
  UniqueKeys = True
  CompareMode = vbTextCompare
  ReInit
End Sub

Public Sub ReInit(Optional ByVal ExpectedItemCount As Long = 5000)
  mHashTableSize = 8
  Do Until mHashTableSize * 5 > ExpectedItemCount: mHashTableSize = mHashTableSize * 2: Loop
  ReDim HashTable(0 To mHashTableSize - 1)
  
  Dim i As Long
  For i = 0 To UBound(HashTable): ReDim HashTable(i).DataIndexes(0 To 0): Next
  mDTUBound = 16: ReDim DataTable(0 To mDTUBound)
  mCount = 0
End Sub

Public Property Get Count() As Long
  Count = mCount
End Property

Public Function Exists(key As String) As Boolean
  Exists = FindIndex(key, CalcHash(key)) > 0
End Function
Public Function IndexByKey(key As String) As Long
  IndexByKey = FindIndex(key, CalcHash(key))
End Function

Public Sub add(Item, Optional key As String)
Dim HashValue As Long, UB As Long
  HashValue = CalcHash(key)
  If UniqueKeys Then If FindIndex(key, HashValue) Then Err.Raise 457
  
  'prolong and add to the new entries to the DataTable-Array
  mCount = mCount + 1
  If mDTUBound < mCount Then mDTUBound = mDTUBound * 1.5: ReDim Preserve DataTable(0 To mDTUBound)
  DataTable(mCount).key = key
  DataTable(mCount).value = Item
  
  'add the new DataIndex to the proper Hash-Buckets .DataIndexes-Array
  With HashTable(HashValue)
    UB = UBound(.DataIndexes): UB = UB + 1
    ReDim Preserve .DataIndexes(0 To UB)
    .DataIndexes(UB) = mCount
  End With
End Sub

Public Property Get KeyByIndex(ByVal IndexOneBased As Long)
  If IndexOneBased < 1 Or IndexOneBased > mCount Then Err.Raise 9
  KeyByIndex = DataTable(IndexOneBased).key
End Property

Public Property Get Item(KeyOrOneBasedIndex)
Dim Index As Long
  If VarType(KeyOrOneBasedIndex) = vbString Then
    Index = FindIndex(KeyOrOneBasedIndex, CalcHash(KeyOrOneBasedIndex))
    If Index = 0 Then Err.Raise 457
  Else
    Index = KeyOrOneBasedIndex
    If Index < 1 Or Index > mCount Then Err.Raise 9
  End If
  If IsObject(DataTable(Index).value) Then
    Set Item = DataTable(Index).value
  Else
    Item = DataTable(Index).value
  End If
End Property

Public Property Let Item(KeyOrOneBasedIndex, NewItem)
Dim Index As Long
  If VarType(KeyOrOneBasedIndex) = vbString Then
    Index = FindIndex(KeyOrOneBasedIndex, CalcHash(KeyOrOneBasedIndex))
    If Index = 0 Then Err.Raise 457
  Else
    Index = KeyOrOneBasedIndex
    If Index < 1 Or Index > mCount Then Err.Raise 9
  End If
  If IsObject(NewItem) Then
    Set DataTable(Index).value = NewItem
  Else
    DataTable(Index).value = NewItem
  End If
End Property
Public Property Set Item(KeyOrOneBasedIndex, NewItem)
  Item(KeyOrOneBasedIndex) = NewItem
End Property

Private Function FindIndex(key, ByVal HashValue As Long) As Long
Dim i As Long, CM As VbCompareMethod
  With HashTable(HashValue)
    CM = CompareMode
    For i = 1 To UBound(.DataIndexes)
      If StrComp(key, DataTable(.DataIndexes(i)).key, CM) = 0 Then
        FindIndex = .DataIndexes(i): Exit Function
      End If
    Next
  End With 'returns Zero, when no Key can be found
End Function

Private Function CalcHash(key) As Long
Dim i As Long, l As Long, B() As Byte
  If CompareMode Then B = LCase$(key) Else B = key
  l = 7919
    For i = UBound(B) To 0 Step -1: l = (i + B(i) + l) * 37 And &H7FFFFF: Next
  CalcHash = l * B(0) Mod mHashTableSize
End Function

Friend Sub CheckHashDistribution()
Dim i As Long, UB As Long, cc As Long, Min As Long, Max As Long
  Min = &H7FFFFFFF
  For i = 0 To UBound(HashTable)
    UB = UBound(HashTable(i).DataIndexes)
    If UB Then
      If Min > UB Then Min = UB
      If Max < UB Then Max = UB
      cc = cc + 1
    End If
  Next
  Debug.Print "Distribution over a HashTable with"; UBound(HashTable) + 1; "slots:"
  Debug.Print "Used-HashSlots:"; cc
  Debug.Print "Min-Entries:"; Min
  Debug.Print "Max-Entries:"; Max
End Sub
