VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbNudgeRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''      *** WORK IN PROGRESS ***
''
''      *** NOT READY FOR GENERAL USE !! ***
''
''      REMAINING TO BE IMPLEMENTED (feature requests welcomed)
''      * Public Function DeleteBlankSheetColumns(wksht As Worksheet)
''      * Public Function DeleteBlankSheetRows(wksht As Worksheet)
''      * Public Function AddBlankListRows(lstObj As ListObject, Optional rowCountAdd, Optional rowCountTotal, Optional truncateAllowed As Boolean = False)
''      * Public Function PutSomething(data, writeAt As Range, Optional canOverwrite As Boolean = False, Optional canInsertCells As Boolean = False, Optional pushDirection As XlDirection)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Misc Utilities for 'pushing' and 'pulling' ("nudging") data around
''      specified range(s)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''      author (c) Paul Brower https://github.com/lopperman/just-VBA
''      ** My VBA Code Is MAC and PC Compatible and always will be **
''
''      class module pbNudgeRange.cls
''      license GNU General Public License v3.0
''      Created Aug-2024
''
''      Download pbNudgeRange.cls with NO DEPENDENCIES at
''      https://github.com/lopperman/just-VBA/blob/main/Code_NoDependencies/pbNudgeRange.cls
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''      IMPLEMENTED FUNCTIONS
''
''




Option Explicit
Option Compare Text
Option Base 1

    Public Enum NudgeOptions
        ndOPTIONS_NOT_SET = 0
        ndRaiseLogEvents = &H1
        ndIgnoreTRACELogEvents = &H2
        ndFillBottomUp = &H4
        ndFillRightLeft = &H8
        ndOnCompleteDeleteBlankRows = &H10
        ndOnCompleteDeleteBlankCols = &H20
        ndFillByRowsOnly = &H40
        ndFillByColsOnly = &H80
        ndFillByColumnThenRow = &H100
'        nd = &H200
'        nd = &H400
'        nd = &H800
'        nd = &H1000
'        nd = &H2000
'        nd = &H4000
'        nd = &H8000&
    End Enum


    Private Enum strMatchEnum
        smEqual = 0
        smNotEqualTo = 1
        smContains = 2
        smStartsWithStr = 3
        smEndWithStr = 4
    End Enum
    Private Enum ecComparisonType
        ecOR = 0 'default
        ecAnd
    End Enum
    

Private l_logMessages As Collection
Attribute l_logMessages.VB_VarHelpID = -1
Private l_options As NudgeOptions


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      EVENTS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Raises Event that can be used for logging
''      [logLevel] will be one of the following:
''              "TRACE", "ERROR"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
        Public Event OnLogMessage(logLevel As String, msg As String)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Raised after error occurs
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Event OnError(ByVal errNumber, ByVal errDescription)


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Default Nudget Options
''      Nude Behavior with Default Options is:
''      1.  Defaults to 'Fill By Column Only'
''      2.  Pushes values left to empty cells
''      3.  Raises Log Events (TRACE and ERROR)

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get DefaultOptions() As NudgeOptions
        Dim resp As NudgeOptions
        resp = NudgeOptions.ndRaiseLogEvents + ndFillByColsOnly
        DefaultOptions = resp
    End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      If the value of [opt] exists in current default options, it will be removed
''      Note if any of the Flag Option Values included in [opt] is NOT
''          in the default nudge options, no change will be made
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function DefaultOptionRemove(opt As NudgeOptions)
            If EnumCompare(l_options, opt, ecAnd) Then
                EnumModify l_options, opt, feVerifyEnumRemoved
            End If
    End Function
    Public Function DefaultOptionInclude(opt As NudgeOptions)
            If EnumCompare(l_options, opt, ecAnd) Then
                EnumModify l_options, opt, feVerifyEnumExists
            End If
    End Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      returns true if all flag enum values in [opt] exist in current default options
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function CheckDefaultOption(opt As NudgeOptions) As Boolean
        CheckDefaultOption = EnumCompare(l_options, opt, ecAnd)
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''       Returns all log messages as collection of strings
''      (Does not require 'OnLogMessage' events to be handled)
''      All TRACE and ERROR Messages are always written to this collection
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Get LogMessages() As Collection
    Set LogMessages = l_logMessages
End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Get Option Settings that are in effect for current instance of class
''      Returns Default NudgeOptions, if user has not set explicitely
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get SessionOptions() As NudgeOptions
        SessionOptions = l_options
    End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Set Option Settings that are in effect for current instance of class
''      Note:  NudgeOptions passed in to any public method always override
''          'SessionOptions'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let SessionOptions(opt As NudgeOptions)
        l_options = opt
    End Property


    Public Function Nudge(rng As Range, Optional nudgeOpt As NudgeOptions = NudgeOptions.ndOPTIONS_NOT_SET)

    End Function



' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Returns item from collection by Key
''   If [key] does not exist in collection, error object with
''   error code 1004 is return
''   suggested use:
''
''   Dim colItem as Variant
''   colItem = CollectionItemByKey([collection], [expectedKey])
''
''   'If expecting object, use 'Set'
''    Set colItem = CollectionItemByKey([collection], [expectedKey])
''
''   If Not IsError(colItem) Then
''       'value was returned
''   Else
''       'error was returned
''   End if
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function CollectionItemByKey(ByRef col As Collection, ByVal key)
On Error Resume Next
    key = CStr(key)
    If IsObject(col(key)) Then
        If Err.number = 0 Then
            Set CollectionItemByKey = col(key)
        End If
    Else
        If Err.number = 0 Then
            CollectionItemByKey = col(key)
        End If
    End If
    If Err.number <> 0 Then
        Err.Clear
        CollectionItemByKey = CVErr(1004)
    End If
End Function
Private Function CollectionKeyExists(ByRef col As Collection, ByVal key)
On Error Resume Next
    key = CStr(key)
    If IsObject(col(key)) Then
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

    Private Function StringsMatch( _
        ByVal checkString As Variant, ByVal _
        validString As Variant, _
        Optional smEnum As strMatchEnum = strMatchEnum.smEqual, _
        Optional compMethod As VbCompareMethod = vbTextCompare) As Boolean
         
    '       IF NEEDED, PUT THIS ENUM AT TOP OF A STANDARD MODULE
            'Public Enum strMatchEnum
            '    smEqual = 0
            '    smNotEqualTo = 1
            '    smContains = 2
            '    smStartsWithStr = 3
            '    smEndWithStr = 4
            'End Enum
            
        Dim str1, str2
            
        str1 = CStr(checkString)
        str2 = CStr(validString)
        Select Case smEnum
            Case strMatchEnum.smEqual
                StringsMatch = StrComp(str1, str2, compMethod) = 0
            Case strMatchEnum.smNotEqualTo
                StringsMatch = StrComp(str1, str2, compMethod) <> 0
            Case strMatchEnum.smContains
                StringsMatch = InStr(1, str1, str2, compMethod) > 0
            Case strMatchEnum.smStartsWithStr
                StringsMatch = InStr(1, str1, str2, compMethod) = 1
            Case strMatchEnum.smEndWithStr
                If Len(str2) > Len(str1) Then
                    StringsMatch = False
                Else
                    StringsMatch = InStr(Len(str1) - Len(str2) + 1, str1, str2, compMethod) = Len(str1) - Len(str2) + 1
                End If
        End Select
    End Function
    
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''   FLAG ENUM COMPARE
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function EnumCompare(theEnum As Variant, enumMember As Variant, Optional ByVal iType As ecComparisonType = ecComparisonType.ecOR) As Boolean
        Dim c As Long
        c = theEnum And enumMember
        EnumCompare = IIf(iType = ecOR, c <> 0, c = enumMember)
    End Function
    
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''   FLAG ENUM - ADD/REMOVE SPECIFIC ENUM MEMBER
    ''   (Works with any flag enum)
    ''   e.g. If you have vbMsgBoxStyle enum and want to make sure
    ''   'DefaultButton1' is included
    ''   msgBtnOption = vbYesNo + vbQuestion
    ''   msgBtnOption = EnumModify(msgBtnOption,vbDefaultButton1,feVerifyEnumExists)
    ''   'now includes vbDefaultButton1, would not modify enum value if it already
    ''   contained vbDefaultButton1
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function EnumModify(theEnum, enumMember, modifyType As FlagEnumModify) As Long
        Dim Exists As Boolean
        Exists = EnumCompare(theEnum, enumMember)
        If Exists And modifyType = feVerifyEnumRemoved Then
            theEnum = theEnum - enumMember
        ElseIf Exists = False And modifyType = feVerifyEnumExists Then
            theEnum = theEnum + enumMember
        End If
        EnumModify = theEnum
    End Function
    

    Private Function MaxV(Val1, Val2)
        If Val1 > Val2 Then
            MaxV = Val1
        Else
            MaxV = Val2
        End If
    End Function
    
    Public Function MinV(Val1, Val2)
        If Val1 > Val2 Then
            MinV = Val2
        Else
            MinV = Val1
        End If
    End Function

Private Function TimeStamp() As String
    TimeStamp = Format(Now, "yyyymmdd hh:mm:ss") & Right(Format(Timer, "0.000"), 4)
End Function

Private Function OnErrorRaise(errNumber, errMessage As String)
    OnLog "ERROR", errMessage
    RaiseEvent OnError(errNumber, errMessage)
    Err.Raise errNumber, Source:=Me, Description:=errMessage
End Function

Private Function OnLog(logLvl As String, msg As String)
    l_logMessages.Add TimeStamp & " " & UCase(logLvl) & " " & msg
    RaiseEvent OnLogMessage(logLvl, msg)
End Function

Private Sub Class_Initialize()
    Set l_logMessages = New Collection
    l_options = DefaultOptions
End Sub
