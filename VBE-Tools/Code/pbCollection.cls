VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Replacement for VBA.Collection on Mac or PC\r\n\r\npbCollection v1.0.0\r\n(c) Paul Brower - https://github.com/lopperman/\r\nAuthor: Paul Brower\r\nLicense: GNU General Public License v3.0\r\n"
Option Explicit
Option Compare Text
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Replacement for the built in Collection object
''      Adds Events for Collection Activities
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''      author (c) Paul Brower https://github.com/lopperman/just-VBA
''      ** My VBA Code Is MAC and PC Compatible and always will be **
''
''      class module pbCollection.cls
''      license GNU General Public License v3.0
''      Created Aug-2024
''
''      Download pbCollection.cls with NO DEPENDENCIES at
''      https://github.com/lopperman/just-VBA/blob/main/Code_NoDependencies/pbCollection.cls
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

Private l_collection As New Collection

Public Event BeforeAdd(item, ByRef Cancel As Boolean)
Public Event ItemAdded(item)
Public Event BeforeRemove(item, ByRef Cancel As Boolean)
Public Event ItemRemoved(item)

Public Sub Add(item, Optional key, Optional Before, Optional After)
Attribute Add.VB_Description = "Add item to collection"
    Dim cancelled As Boolean
    RaiseEvent BeforeAdd(item, cancelled)
    If cancelled = False Then
        l_collection.Add item, key, Before, After
        RaiseEvent ItemAdded(item)
    End If
End Sub
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns number of items in collection"
    Count = l_collection.Count
End Property
Public Function item(index)
Attribute item.VB_Description = "Returns item from collection if exists\r\nIf [index] is string, will return by key"
    item = l_collection.item(index)
End Function
Public Function KeyExists(key As String) As Boolean
    KeyExists = CollectionKeyExists(key)
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enables using 'For Each' / enumerating collection"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = l_collection.[_NewEnum]
End Property

Public Sub Remove(index)
Attribute Remove.VB_Description = "Remove item from collection if exists\r\nIf [index] is string, will remove by key"
    Dim cancelled As Boolean
    Dim item
    If TypeName(index) = "String" Then
        If CollectionKeyExists(index) Then
            If IsObject(CollectionItemByKey(index)) Then
                Set item = CollectionItemByKey(index)
            Else
                item = CollectionItemByKey(index)
            End If
        End If
    Else
        If IsObject(l_collection(index)) Then
            Set item = l_collection(index)
        Else
            item = l_collection(index)
        End If
    End If
    
    RaiseEvent BeforeRemove(item, cancelled)
    If cancelled = False Then
        l_collection.Remove index
        RaiseEvent ItemRemoved(item)
    End If
End Sub


Private Function CollectionItemByKey(ByVal key)
Attribute CollectionItemByKey.VB_Description = "Return collection item by [key] if exists, otherwise returns CVErr"
On Error Resume Next
    key = CStr(key)
    If IsObject(l_collection(key)) Then
        If Err.number = 0 Then
            Set CollectionItemByKey = l_collection(key)
        End If
    Else
        If Err.number = 0 Then
            CollectionItemByKey = l_collection(key)
        End If
    End If
    If Err.number <> 0 Then
        Err.Clear
        CollectionItemByKey = CVErr(1004)
    End If
End Function
Private Function CollectionKeyExists(ByVal key)
On Error Resume Next
    key = CStr(key)
    If IsObject(l_collection(key)) Then
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

