VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbCommonEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  Common Events
'  This class cannot be instantiated; it is configured to be a
'   static class module
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module pbCommonEvents.cls
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

Option Explicit
Option Compare Text
Option Base 1

Public Event OnErrorHandlerBegin(ByVal errNumber, ByVal errDescription, ByVal errSource)
Public Event OnErrorHandlerEnd(ByVal errNumber, ByVal errDescription, ByVal errSource)
Public Event OnRaiseError(ByVal errNumber, ByVal errDescription, ByVal errSource, ByRef Cancel As Boolean)
Public Event OnOpenStart(ByRef Cancel As Boolean)
Public Event BeforeClose(ByRef Cancel As Boolean)
Public Event OnOpenComplete()
Public Event OnBeforeProcess()
Public Event ChangeFontRequest(ByVal wksht As Worksheet)
Public Event OnShapeButtonClick(ByVal shapeButton As Shape)
Public Event ListRowsChanged(lstObj As ListObject, rowIds() As Long)
Public Event ListRowsAdded(lstObj As ListObject, rowIds() As Long)
Public Event ListRowDeleted(lstObj As ListObject, deletedVals() As Variant)

Private WithEvents l_stg As pbSettings
Attribute l_stg.VB_VarHelpID = -1


Public Function RaiseOnClose(ByRef Cancel As Boolean)
    LogTRACE "START pbCommonEvents.RaiseOnClose"
    RaiseEvent BeforeClose(Cancel)
    LogTRACE "END pbCommonEvents.RaiseOnClose - cancelled = " & Cancel
End Function

Public Function OnListRowDeleted(lstObj As ListObject, deletedVals() As Variant)
    LogTRACE "START pbCommonEvents.OnListRowDeleted - " & lstObj.Name
    RaiseEvent ListRowDeleted(lstObj, deletedVals)
    LogTRACE "END pbCommonEvents.OnListRowDeleted - " & lstObj.Name
End Function
Public Function OnListRowsAdded(lstObj As ListObject, rowIds() As Long)
    LogTRACE "START pbCommonEvents.OnListRowsAdded - " & lstObj.Name
    RaiseEvent ListRowsAdded(lstObj, rowIds)
    LogTRACE "END pbCommonEvents.OnListRowsAdded - " & lstObj.Name
End Function
Public Function OnListRowsChanged(lstObj As ListObject, rowIds() As Long)
    LogTRACE "START pbCommonEvents.OnListRowsChanged - " & lstObj.Name
    RaiseEvent ListRowsAdded(lstObj, rowIds)
    LogTRACE "END pbCommonEvents.OnListRowsChanged - " & lstObj.Name
End Function

Public Function RaiseOnShapeButtonClick(ByVal shapeButton As Shape)
    ftBeep btButton
    LogTRACE "START pbCommonEvents.RaiseOnShapeButtonClick - " & shapeButton.Name
    RaiseEvent OnShapeButtonClick(shapeButton)
    LogTRACE "END pbCommonEvents.RaiseOnShapeButtonClick - " & shapeButton.Name
End Function
Public Function RaiseChangeFontRequest(ByVal wksht As Worksheet)
    LogTRACE "START pbCommonEvents.RaiseChamgeFontRequest - " & wksht.CodeName
    RaiseEvent ChangeFontRequest(wksht)
    LogTRACE "END pbCommonEvents.RaiseChangeFontRequest - " & wksht.CodeName
End Function
Public Function RaiseOnOpenStart(ByRef Cancel As Boolean)
    LogTRACE "START pbCommonEvents.RaiseOnOpenStart"
    RaiseEvent OnOpenStart(Cancel)
    LogTRACE "END pbCommonEvents.RaiseOnOpenStart - cancelled = " & Cancel
End Function
Public Function RaiseOnOpenComplete()
    LogTRACE "START pbCommonEvents.RaiseOnOpenComplete"
    RaiseEvent OnOpenComplete
    LogTRACE "END pbCommonEvents.RaiseOnOpenComplete"
End Function
Public Function RaiseOnErrorHandlerBegin(errNumber, errDescription, errSource)
    RaiseEvent OnErrorHandlerBegin(errNumber, errDescription, errSource)
End Function
Public Function RaiseOnErrorHandlerEnd(errNumber, errDescription, errSource)
    LogTRACE "START pbCommonEvents.RaiseOnErrorHandlerEnd " & errNumber & " - " & errDescription
    RaiseEvent OnErrorHandlerEnd(errNumber, errDescription, errSource)
    LogTRACE "END pbCommonEvents.RaiseOnErrorHandlerEnd " & errNumber & " - " & errDescription
End Function
Public Function RaiseOnRaiseError(ByVal errNumber, ByVal errDescription, ByVal errSource, ByRef Cancel As Boolean)
    LogTRACE "START pbCommonEvents.RaiseOnRaiseError " & errNumber & " - " & errDescription
    RaiseEvent OnRaiseError(errNumber, errDescription, errSource, Cancel)
    LogTRACE "END pbCommonEvents.RaiseOnRaiseError " & errNumber & " - " & errDescription
End Function


Private Sub Class_Initialize()
    'Only the default instance can run this
    If Not Me Is pbCommonEvents Then RaiseError ERR_CLASS_INSTANCE_PROHIBITED, errorDesc:="pbCommonEvents may not be created as a new Class Instance outside it's PredeclaredId 'pbCommonEvents'"
    Set l_stg = stg
End Sub

Private Sub l_stg_OnAlert(ByVal isError As Boolean, ByVal msg As Variant)
    If isError Then
        LogERROR msg
    Else
        LogWARN msg
    End If
    
End Sub

Private Sub l_stg_OnDeleteSetting(ByVal settingKey As Variant, Cancel As Boolean)
    LogWARN pbCommonUtil.UserName & " DELETING Setting Key: " & settingKey
End Sub
