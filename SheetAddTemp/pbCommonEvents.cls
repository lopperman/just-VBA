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

Public Event OnErrorHandlerBegin(ByVal errNumber, ByVal errDescription, ByVal errSource)
Public Event OnErrorHandlerEnd(ByVal errNumber, ByVal errDescription, ByVal errSource)
Public Event OnRaiseError(ByVal errNumber, ByVal errDescription, ByVal errSource, ByRef Cancel)


Public Function RaiseOnErrorHandlerBegin(errNumber, errDescription, errSource)
    RaiseEvent OnErrorHandlerBegin(errNumber, errDescription, errSource)
End Function
Public Function RaiseOnErrorHandlerEnd(errNumber, errDescription, errSource)
    RaiseEvent OnErrorHandlerEnd(errNumber, errDescription, errSource)
End Function
Public Function RaiseOnRaiseError(ByVal errNumber, ByVal errDescription, ByVal errSource, ByRef Cancel)
    RaiseEvent OnRaiseError(errNumber, errDescription, errSource, Cancel)
    'RaiseOnRaiseError = Cancel
End Function

Private Sub Class_Initialize()
    'Only the default instance can run this
    If Not Me Is pbCommonEvents Then RaiseError ERR_CLASS_INSTANCE_PROHIBITED, errorDesc:="pbCommonEvents may not be created as a new Class Instance outside it's PredeclaredId 'pbCommonEvents'"
End Sub

