VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbCollectionTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents pbCol As pbCollection
Attribute pbCol.VB_VarHelpID = -1

Public Function work()

    Debug.Print "Items in Collecction: " & pbCol.Count
    pbCol.Add 1, key:="A"
    Debug.Print "Items in Collecction: " & pbCol.Count
    pbCol.Add 2, key:="B"
    Debug.Print "Items in Collecction: " & pbCol.Count
        
    If pbCol.KeyExists("A") Then
        pbCol.Remove "A"
        Debug.Print "Items in Collecction: " & pbCol.Count
    End If
    If pbCol.KeyExists("B") Then
        pbCol.Remove "B"
        Debug.Print "Items in Collecction: " & pbCol.Count
    End If
    

End Function

Private Sub Class_Initialize()
    Set pbCol = New pbCollection
End Sub

Private Sub pbCol_BeforeAdd(item As Variant, Cancel As Boolean)
    If MsgBox("Cancel Adding", vbYesNo + vbDefaultButton2) = vbYes Then
        Cancel = True
        Debug.Print TypeName(item) & " was not added because user cancelled"
    End If
End Sub

Private Sub pbCol_BeforeRemove(item As Variant, Cancel As Boolean)
    If MsgBox("Cancel Removing", vbYesNo + vbDefaultButton2) = vbYes Then
        Cancel = True
        Debug.Print TypeName(item) & " was not removed because user cancelled"
    End If
End Sub

Private Sub pbCol_ItemAdded(item As Variant)
    Debug.Print TypeName(item) & " was added"
End Sub

Private Sub pbCol_ItemRemoved(item As Variant)
    Debug.Print TypeName(item) & " was removed"
End Sub
