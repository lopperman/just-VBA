VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbKillAddin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
''' Manage (disable and delete) Previous FinToolAddin
''' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'Option Explicit
'Option Base 1
'Option Compare Text
'
'Private Const FT_ADDIN_NAME As String = "FinToolADDIN.xlam"
'Private Const FT_ADDIN_NAME_XLSM As String = "FinToolADDIN.xlsm"
'
'
''' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
''' Returns true if FinToolAddin is Non-Existent, or was Found and Removed
''' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'Public Function FindAndDestroy(ParamArray addinWkbkNames() As Variant) As Boolean
'    On Error Resume Next
'
'    Dim wkbkName
'    For Each wkbkName In addinWkbkNames
'        If Not AddinOpen Then
'            FindAndDestroy = True
'            LogDEBUG "pbKillAddin.FindAndDestroy - FinToolAddin was not found"
'        Else
'            LogDEBUG "ftAddin.FindAndDestroy - FinToolAddin was found running -- attempting to remove"
'            If RemoveAddin Then
'                LogDEBUG "ftAddin.FindAndDestroy - FinToolAddin was removed successfully!"
'                FindAndDestroy = True
'            Else
'                LogERROR "ftAddin.FindAndDestroy - FinToolAddin was not removed"
'                FindAndDestroy = False
'            End If
'        End If
'        If Err.number <> 0 Then
'            Err.Clear
'        End If
'
'    Next
'
'    If Not AddinOpen Then
'        FindAndDestroy = True
'        LogDEBUG "pbKillAddin.FindAndDestroy - FinToolAddin was not found"
'    Else
'        LogDEBUG "ftAddin.FindAndDestroy - FinToolAddin was found running -- attempting to remove"
'        If RemoveAddin Then
'            LogDEBUG "ftAddin.FindAndDestroy - FinToolAddin was removed successfully!"
'            FindAndDestroy = True
'        Else
'            LogERROR "ftAddin.FindAndDestroy - FinToolAddin was not removed"
'            FindAndDestroy = False
'        End If
'    End If
'    If Err.number <> 0 Then
'        Err.Clear
'    End If
'End Function
'
'Private Function RemoveAddin() As Boolean
'    On Error Resume Next
'
'    Dim delPath As String
'    Dim deleteXlam As Workbook
'    Set deleteXlam = FTHelper
'    delPath = deleteXlam.FullName
'
'    If deleteXlam.IsAddin = False Then
'        FTHelper.Close SaveChanges:=False
'        Set deleteXlam = Nothing
'    ElseIf deleteXlam.IsAddin = True Then
'        Dim tAddin As AddIn
'        For Each tAddin In Application.AddIns2
'            If StrComp(tAddin.Name, FT_ADDIN_NAME, vbTextCompare) = 0 Then
'                tAddin.Installed = False
'                Exit For
'            End If
'        Next
'        If AddinOpen Then
'            FTHelper.Close
'        End If
'        Set deleteXlam = Nothing
'        Kill delPath
'    End If
'
'    DoEvents
'    RemoveAddin = Not AddinOpen
'
'    If Err.number <> 0 Then
'        Err.Clear
'    End If
'
'End Function
'
'Private Function FTHelper() As Workbook
'    On Error Resume Next
'    Dim tmpWkbk As Workbook
'    Dim wb As Workbook
'    For Each wb In Application.Workbooks
'        If StrComp(wb.Name, FT_ADDIN_NAME_XLSM, vbTextCompare) = 0 Then
'            Set tmpWkbk = wb
'            Exit For
'        End If
'    Next wb
'    If tmpWkbk Is Nothing Then
'        If Application.AddIns2.Count > 0 Then
'            Dim iAddin As Long
'            For iAddin = 1 To Application.AddIns2.Count
'                If StrComp(Application.AddIns2(iAddin).Name, FT_ADDIN_NAME, vbTextCompare) = 0 Then
'                    If Application.AddIns2(iAddin).IsOpen Then
'                        Set tmpWkbk = Workbooks(FT_ADDIN_NAME)
'                        Exit For
'                    End If
'                End If
'            Next iAddin
'        End If
'    End If
'
'    If Err.number <> 0 Then Err.Clear
'    If Not tmpWkbk Is Nothing Then
'        Set FTHelper = tmpWkbk
'    End If
'    Set tmpWkbk = Nothing
'
'End Function
'
'Private Function AddinOpen(wkbkName) As Boolean
'    On Error Resume Next
'    AddinOpen = Not Workbooks(wkbkName) Is Nothing
'
'    If Workbooks(wkbkName) Is Nothing Then
'
'    AddinOpen = Not FTHelper Is Nothing
'
'    If Err.number <> 0 Then
'        Err.Clear
'    End If
'End Function
'
'
