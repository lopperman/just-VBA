Attribute VB_Name = "pbProtection"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' pbProtection
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' General  Helper Utilities for Working with Worksheet Protection
'
' @module pbProtection
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  GLOBAL CONST & ENUMS
''
'' As an FYI, you can create a Public Enum in a class module, and the type
''  will be available anywhere in your project that contains that class file
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    '' DEFAULT SHEET PROTECTION PASSWORD
    Public Const CFG_PROTECT_PASSWORD As String = "00000"
    
    '' CONTROLS WHETHER THE 'UserInterfaceOnly' OPTION IS ALWAYS
    '' INCLUDED WHEN PROTECTING A WORKSHEET
    ''  THIS OPTION ALLOWS VBA TO PERFORM ACTIONS THAT HAVE
    ''  BEEN RESTRICTED FOR USERS
    Public Const CFG_USER_INTERFACE_ONLY_FORCE As Boolean = False

    ''  SHEET PROTECTION OPTIONS ENUM
    Public Enum SheetProtection
        psContents = 2 ^ 0
        ''psUsePassword = 2 ^ 1
        psDrawingObjects = 2 ^ 2
        psScenarios = 2 ^ 3
        psUserInterfaceOnly = 2 ^ 4
        psAllowFormattingCells = 2 ^ 5
        psAllowFormattingColumns = 2 ^ 6
        psAllowFormattingRows = 2 ^ 7
        psAllowInsertingColumns = 2 ^ 8
        psAllowInsertingRows = 2 ^ 9
        psAllowInsertingHyperlinks = 2 ^ 10
        psAllowDeletingColumns = 2 ^ 11
        psAllowDeletingRows = 2 ^ 12
        psAllowSorting = 2 ^ 13
        psAllowFiltering = 2 ^ 14
        psAllowUsingPivotTables = 2 ^ 15
    End Enum
    
    
    ''  FLAG ENUM HELPER ENUM
    Public Enum FlagEnumModify
        feVerifyEnumExists
        feVerifyEnumRemoved
    End Enum
    '' FLAG ENUM COMPARISON ENUM
    Public Enum ecComparisonType
        ecOR = 0 'default
        ecAnd
    End Enum

    ''  ENUM FOR HELPER FUNCTION 'STRINGSMATCH'
    Public Enum strMatchEnum
         smEqual = 0
         smNotEqualTo = 1
         smContains = 2
         smStartsWithStr = 3
         smEndWithStr = 4
     End Enum

 
 
 
 ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  PROTECT SHEET
''
''  @wksht: worksheet to project (required)
''  @options: Flags Enum that contains protection options to apply (optional)
''      (See 'DefaultProtectOptions' Property for default options used if
''      this argument is excluded)
''  @password: password to protect sheet (optional)
''      If excluded, uses public Const: CFG_PROTECT_PASSWORD
 ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function ProtectSheet(ByRef wksht As Worksheet _
    , Optional options As SheetProtection _
    , Optional password = CFG_PROTECT_PASSWORD)
    
    Dim protectDrawingObjects As Boolean
    Dim protectContents As Boolean
    Dim protectScenarios As Boolean
    Dim userInterfaceOnly As Boolean
    Dim allowFormattingCells As Boolean
    Dim allowFormattingColumns As Boolean
    Dim allowFormattingRows As Boolean
    Dim allowInsertingColumns As Boolean
    Dim allowInsertingRows As Boolean
    Dim allowInsertingHyperlinks As Boolean
    Dim allowDeletingColumns As Boolean
    Dim allowDeletingRows As Boolean
    Dim allowSorting As Boolean
    Dim allowFiltering As Boolean
    Dim allowUsingPivotTables As Boolean
    
    '' Use Default Options if argument was not included
    If options = 0 Then
        options = DefaultProtectOptions
    End If
    
    '' Configure Protection Arguments
    protectDrawingObjects = EnumCompare(options, SheetProtection.psDrawingObjects)
    protectContents = EnumCompare(options, SheetProtection.psContents)
    protectScenarios = EnumCompare(options, SheetProtection.psScenarios)
    userInterfaceOnly = EnumCompare(options, SheetProtection.psUserInterfaceOnly)
    allowFormattingCells = EnumCompare(options, SheetProtection.psAllowFormattingCells)
    allowFormattingColumns = EnumCompare(options, SheetProtection.psAllowFormattingColumns)
    allowFormattingRows = EnumCompare(options, SheetProtection.psAllowFormattingRows)
    allowInsertingColumns = EnumCompare(options, SheetProtection.psAllowInsertingColumns)
    allowInsertingRows = EnumCompare(options, SheetProtection.psAllowInsertingRows)
    allowInsertingHyperlinks = EnumCompare(options, SheetProtection.psAllowInsertingHyperlinks)
    allowDeletingColumns = EnumCompare(options, SheetProtection.psAllowDeletingColumns)
    allowDeletingRows = EnumCompare(options, SheetProtection.psAllowDeletingRows)
    allowSorting = EnumCompare(options, SheetProtection.psAllowSorting)
    allowFiltering = EnumCompare(options, SheetProtection.psAllowFiltering)
    allowUsingPivotTables = EnumCompare(options, SheetProtection.psAllowUsingPivotTables)

    '' Call Private ProtectSheet2 to perform protection call
    ProtectSheet = ProtectSheet2(wksht, password, protectDrawingObjects, _
        protectContents, protectScenarios, userInterfaceOnly, allowFormattingCells, _
        allowFormattingColumns, allowFormattingRows, allowInsertingColumns, _
        allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, _
        allowDeletingRows, allowSorting, allowFiltering, allowUsingPivotTables)

End Function

Private Function ProtectSheet2(ByRef wksht As Worksheet _
    , Optional password = CFG_PROTECT_PASSWORD _
    , Optional protectDrawingObjects As Boolean = True _
    , Optional protectContents As Boolean = True _
    , Optional protectScenarios As Boolean = False _
    , Optional userInterfaceOnly As Boolean = True _
    , Optional allowFormattingCells As Boolean = True _
    , Optional allowFormattingColumns As Boolean = True _
    , Optional allowFormattingRows As Boolean = True _
    , Optional allowInsertingColumns As Boolean = False _
    , Optional allowInsertingRows As Boolean = False _
    , Optional allowInsertingHyperlinks As Boolean = False _
    , Optional allowDeletingColumns As Boolean = False _
    , Optional allowDeletingRows As Boolean = False _
    , Optional allowSorting As Boolean = True _
    , Optional allowFiltering As Boolean = True _
    , Optional allowUsingPivotTables As Boolean = False)

    ''  Enforce UserInterfaceOnly if constant below is True
    If CFG_USER_INTERFACE_ONLY_FORCE Then
        userInterfaceOnly = True
    End If

    With wksht
       .Protect password:=password _
        , DrawingObjects:=protectDrawingObjects _
        , Contents:=protectContents _
        , Scenarios:=protectScenarios _
        , userInterfaceOnly:=userInterfaceOnly _
        , allowFormattingCells:=allowFormattingCells _
        , allowFormattingColumns:=allowFormattingColumns _
        , allowFormattingRows:=allowFormattingRows _
        , allowInsertingColumns:=allowInsertingColumns _
        , allowInsertingRows:=allowInsertingRows _
        , allowInsertingHyperlinks:=allowInsertingHyperlinks _
        , allowDeletingColumns:=allowDeletingColumns _
        , allowDeletingRows:=allowDeletingRows _
        , allowSorting:=allowSorting _
        , allowFiltering:=allowFiltering _
        , allowUsingPivotTables:=allowUsingPivotTables
    End With
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  UNPROTECT SHEET
''
''  @password: password to UNprotect sheet (optional)
''      If excluded, uses public Const: CFG_PROTECT_PASSWORD
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function UnprotectSheet(ByRef wksht As Worksheet _
    , Optional pwd As String = CFG_PROTECT_PASSWORD)
    wksht.Unprotect password:=pwd
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  UNPROTECT ALL SHEETS
''
''  @wkbk:Workbook to perform action, if not same workbook that contains
''      this method (Optional)
''  @unhideAll: set to True to ensure worksheets are visible (Optional, Default False)
''  @password: password to UNprotect sheet (optional)
''      If excluded, uses public Const: CFG_PROTECT_PASSWORD
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function UnprotectAllSheets( _
    Optional wkbk As Workbook _
    , Optional unhideAll As Boolean = False _
    , Optional password As String = CFG_PROTECT_PASSWORD)
    If wkbk Is Nothing Then
        Set wkbk = ThisWorkbook
    End If
    Dim tWS As Worksheet
    For Each tWS In wkbk.Worksheets
        UnprotectSheet tWS, pwd:=password
        If Not tWS.Visible = xlSheetVisible And unhideAll Then
            tWS.Visible = xlSheetVisible
        End If
    Next tWS
End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Return Default 'SheetProtection' Flag Enum
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Property Get DefaultProtectOptions() As SheetProtection
    DefaultProtectOptions = _
        psAllowFiltering _
        + psAllowFormattingCells _
        + psAllowFormattingColumns _
        + psAllowFormattingRows _
        + psDrawingObjects _
        + psUserInterfaceOnly _
        + psContents _
        + psAllowSorting _
        + psScenarios
End Property


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  HELPER FUNCTIONS
''
''  MOST OF THESE CAN BE FOUND IN MY pbCommon.bas module at:
''  https://github.com/lopperman/just-VBA/blob/main/Code_NoDependencies/
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function StringsMatch( _
    ByVal str1 As Variant, ByVal _
    str2 As Variant, _
    Optional smEnum As strMatchEnum = strMatchEnum.smEqual, _
    Optional compMethod As VbCompareMethod = vbTextCompare) As Boolean
    
    '' Requires Enum Below
    '' Public Enum strMatchEnum
    ''     smEqual = 0
    ''     smNotEqualTo = 1
    ''     smContains = 2
    ''     smStartsWithStr = 3
    ''     smEndWithStr = 4
    '' End Enum
    
    str1 = CStr(str1)
    str2 = CStr(str2)
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   FLAG ENUM COMPARE
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function EnumCompare(theEnum As Variant, enumMember As Variant, Optional ByVal iType As ecComparisonType = ecComparisonType.ecOR) As Boolean
    Dim c As Long
    c = theEnum And enumMember
    EnumCompare = IIf(iType = ecOR, c <> 0, c = enumMember)
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   FLAG ENUM - ADD/REMOVE SPECIFIC ENUM MEMBER
'   (Works with any flag enum)
'   e.g. If you have vbMsgBoxStyle enum and want to make sure
'   'DefaultButton1' is included
'   msgBtnOption = vbYesNo + vbQuestion
'   msgBtnOption = EnumModify(msgBtnOption,vbDefaultButton1,feVerifyEnumExists)
'   'now includes vbDefaultButton1, would not modify enum value if it already
'   contained vbDefaultButton1
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function EnumModify(theEnum, enumMember, modifyType As FlagEnumModify) As Long
    Dim Exists As Boolean
    Exists = EnumCompare(theEnum, enumMember)
    If Exists And modifyType = feVerifyEnumRemoved Then
        theEnum = theEnum - enumMember
    ElseIf Exists = False And modifyType = feVerifyEnumExists Then
        theEnum = theEnum + enumMember
    End If
    EnumModify = theEnum
End Function
