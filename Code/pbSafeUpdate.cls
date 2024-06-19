VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbSafeUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' pbSafeUpdate v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' Ensure ActiveCell is not in ListObject When Updating from Array
'
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' This class is designed work as a 'Default Instance Variable' (see url below)
' https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/189fb41b-cc3a-4999-a6d2-ba89f72d2870
'
' If this code is copy/pasted, then you'll need to export this class, edit the top portion so
' that: Attribute VB_PredeclaredId = True
' Import the updated class into your workspace
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1

'movedFrom should be a Range Object in
'order to prevent accidentally moving ActiveCell/Selection
'to the wrong workbook

Private movedFrom As Range
Private isSingleton As Boolean

'' Use this function to perform update - BeforeEdit and AfterEdit are
'' called automatically
Public Function UpdateListObjRange(lstObjRng As Range, srcArray)
On Error GoTo E:
    Dim dimension1 As Long, dimension2 As Long
    Dim evts As Boolean, scrn As Boolean

    dimension1 = (UBound(srcArray, 1) - LBound(srcArray, 1)) + 1
    dimension2 = (UBound(srcArray, 2) - LBound(srcArray, 2)) + 1
    
    If Not dimension1 = lstObjRng.Rows.Count Or Not dimension2 = lstObjRng.Columns.Count Then
        Err.Raise 1004, "pbSafeUpdate.UpdateListObjRange", "'srcArray' dimentions must match 'lstObjRng' row and column size"
    End If
        
    evts = Application.EnableEvents
    scrn = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    BeforeEdit
    lstObjRng.Value = srcArray
    UpdateListObjRange = True
    
Finalize:
    On Error Resume Next
    AfterEdit
    Application.EnableEvents = evts
    Application.ScreenUpdating = scrn
    Exit Function
E:
    ''Implement your desired error handling
    ''ErrorCheck
    Err.Raise Err.number, Err.Source, Err.Description
    Resume Finalize:

End Function

' call this immediately before updating a listobject from an array
' no need to check anything before calling.  If the activecell is in a listobject,
' the activecell will be moved just outside the UsedRange of the current sheet
' The screen will NOT scroll to the new ActiveCell location
' syntax:  pbSafeupdate.BeforeEdit
Public Function BeforeEdit()
    If Not isSingleton Then Exit Function
    ClearValues
    If Not ActiveCell Is Nothing Then
        If Not ActiveCell.ListObject Is Nothing Then
            Set movedFrom = ActiveCell
            MoveAway
        End If
    End If
End Function

' call this immediately after updating a listobject from an array
' no need to check anything before calling.  If the activecell was in a listobject
' before the update, then the activecell will be moved back to that location
' syntax:  pbSafeupdate.AfterEdit
Public Function AfterEdit()
   If Not isSingleton Then Exit Function
     'We can only move back if the movedFrom.Worksheet is the Active Worksheet
    If movedFrom Is Nothing Then Exit Function
    If ActiveSheet Is Nothing Then Exit Function
    If Not ActiveSheet Is movedFrom.Worksheet Then
        ClearValues
        Exit Function
    End If
    MoveBack
End Function

Private Function MoveBack()
    
    Dim scrn As Boolean, evts As Boolean
    scrn = Application.ScreenUpdating
    evts = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If Not movedFrom Is Nothing Then
        movedFrom.Select
        ClearValues
    End If
    
    Application.EnableEvents = evts
    Application.ScreenUpdating = scrn

End Function

Private Function MoveAway()

    Dim scrn As Boolean, evts As Boolean
    scrn = Application.ScreenUpdating
    evts = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If movedFrom.Worksheet.usedRange.Columns.Count < movedFrom.Worksheet.Columns.Count Then
        movedFrom.Worksheet.Cells(1, movedFrom.Worksheet.usedRange.Columns.Count + 1).Select
    End If
    
    Application.EnableEvents = evts
    Application.ScreenUpdating = scrn
    
End Function

Private Function ClearValues()
    Set movedFrom = Nothing
End Function

Private Sub Class_Initialize()
    isSingleton = True
    If Not Me Is pbSafeUpdate Then
        isSingleton = False
        Err.Raise 1004, Source:="pbSafeUpdate", Description:="pbSafeUpdate may not be created as a new Class Instance outside it's PredeclaredId 'pbSafeUpdate'"
    End If
End Sub

