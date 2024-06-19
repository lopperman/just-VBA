VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbKeyVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  GENERIC KEV/VALUE STRUCTURE
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbSPCommon.bas
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
Option Explicit
Option Base 1
Option Compare Text

Public key
Public Value

Public Property Get IsKeyObject() As Boolean
     IsKeyObject = IsObject(key)
End Property
Public Property Get IsValueObject() As Boolean
    IsValueObject = IsObject(Value)
End Property
Public Property Get ToString() As String
    ToString = ConcatWithDelim(" ", "KeyType=", TypeName(key), IIf(IsKeyObject, "'Object'", "'") & key & "'") & ", " & _
        ConcatWithDelim(" ", "ValueType=", TypeName(Value), IIf(IsValueObject, "'Object'", "'") & Value & "'")
End Property
