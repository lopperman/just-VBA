VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbWKBK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1
Option Compare Text





Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If AppMode = appStatusRunning Then
        Cancel = True
        QuitApp
    End If
End Sub
