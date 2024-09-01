VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsCodeUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1
Option Compare Text


Public Function OnFormat()
    ProtectSheet Me, psProtectDrawingObjects + psUserInterfaceOnly + psProtectContents
'    With Me
'        .EnableSelection = xlNoSelection
'        .Range("P1").Resize(ColumnSize:=(.Columns.Count - .Range("P1").column + 1)).EntireColumn.Hidden = True
'        .Range("A23").Resize(RowSize:=(.rows.Count - .Range("A23").Row + 1)).EntireRow.Hidden = True
'    End With
    CheckPBShapeButtons Me
End Function
