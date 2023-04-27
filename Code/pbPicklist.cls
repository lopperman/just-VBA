VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbPicklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Utility that enables you to used a dynamically built worksheet
'   to enable users to pick one or more values from a list
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbPicklist.cls
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Option Explicit
Option Base 1
Option Compare Text

Private arrKeys() As Variant
Private arrValues() As Variant
Private multiSelectMode As Boolean
Private minimumSelectCount As Long
Private showKeyColumn As Boolean
Private pickSht As Worksheet

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'   Name of Worksheet this is used for users to select items
'   from a list.  (Automatically Built if Missing)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Private Const PICK_SHEET_NAME As String = "pb-Picklist"

Public Event CancelPick(ByRef picklist As pbPicklist)
Public Event CompletePick(ByRef picklist As pbPicklist, selectedKeys() As Variant, selectedValues() As Variant)


Public Property Get PickSheet() As Worksheet

End Property
Public Property Get Keys() As Variant()
    Keys = arrKeys
End Property
Public Property Get Values() As Variant()
    Values = arrValues
End Property
Private Function OnCancel()
    
    RaiseEvent CancelPick(Me)
End Function

Private Function OnComplete()

    RaiseEvent CompletePick(Me, Keys, Values)
End Function


