VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBItemProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  Manage Collection of VB Component Info
'  Intended Use: Comparing VBA Code
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  author (c) Paul Brower https://github.com/lopperman/just-VBA
'  module VBItemProps.cls
'  license GNU General Public License v3.0
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1

Private lWkbk As Workbook
Private lcomp As VBComponent
Private lProps As Dictionary
'Private lMembers Dictionary

Public Enum ClassInstanceType
    VBEInstancingNA = 0
    VBEClassPrivate = 1
    VBEClassPublicNotCreatable = 2
End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Part of the private implentation of the 'privEnumCompare' function
    Private Enum ecComparisonTypePriv
        ecOR = 0 'default
        ecAnd
    End Enum
    Private Enum FlagEnumModifyPriv
        feVerifyEnumExists
        feVerifyEnumRemoved
    End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Part of the private implentation of the 'StringsMatch' function
''  See: https://www.reddit.com/r/vba/comments/w74j85/handy_stringsmatch_method_that_handles_equal_not/?utm_source=share&utm_medium=web2x&context=3
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Enum strMatchEnumPriv
        smEqual = 0
        smNotEqualTo = 1
        smContains = 2
        smStartsWithStr = 3
        smEndWithStr = 4
    End Enum


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'  INITIALIZE CLASS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function Populate(wkbk As Workbook, vbCmpt As VBComponent)
    Set lWkbk = wkbk
    Set lcomp = vbCmpt
    PopulateProps
End Function

Public Function GetLines(Optional includeBlank As Boolean = False, Optional includeComments As Boolean = False) As Collection
    Dim lCount As Long, i As Long
    Dim resp As New Collection
    Dim tLine, tInclude As Boolean
    lCount = LineCount
    For i = 1 To lCount
        tInclude = True
        tLine = ItemModule.lines(i, 1)
        If LineBlank(tLine) Then
            If includeBlank = False Then tInclude = False
        ElseIf LineComment(tLine) Then
            If includeComments = False Then tInclude = False
        End If
        If tInclude Then
            resp.add tLine, Key:="Line" & Format(i, "0000")
        End If
    Next i
    Set GetLines = resp
    Set resp = Nothing
End Function

Private Function LineBlank(tLine) As Boolean
    LineBlank = Len(Trim(tLine)) = 0
End Function
Private Function LineComment(tLine) As Boolean
    LineComment = StringsMatchPriv(Trim(tLine), "'", smStartsWithStr)
End Function

Public Property Get ItemInstancing() As ClassInstanceType
    Dim tIns As Variant
    If lProps.Count > 0 Then
        If lProps.Exists("Instancing") Then
            tIns = lProps("Instancing")
        End If
    End If
    Select Case tIns
        Case 1
            ItemInstancing = VBEClassPrivate
        Case 2
            ItemInstancing = VBEClassPublicNotCreatable
        Case Else
            ItemInstancing = VBEInstancingNA
    End Select
    
End Property

Public Property Get LineCount() As Long
    LineCount = lcomp.CodeModule.CountOfLines
End Property

Public Property Get DeclarationLineCount() As Long
    DeclarationLineCount = lcomp.CodeModule.CountOfDeclarationLines
End Property

Public Property Get ItemType() As vbext_ComponentType
    ItemType = lcomp.Type
End Property

Public Property Get ItemName() As String
    ItemName = lcomp.Name
End Property

Public Property Get ItemModule() As CodeModule
    Set ItemModule = lcomp.CodeModule
End Property

Public Function PropertyExists(propName) As Boolean
    PropertyExists = lProps.Exists(propName)
End Function
Public Function PropertyObj(propName) As Property
    If lProps.Exists(propName) Then
        Set PropertyObj = lcomp.Properties(propName)
    End If
End Function
Public Function PropertyVal(propName) As Variant
    On Error Resume Next
    Dim isObj As Boolean
    isObj = IsObject(PropertyObj(propName).value)
    If Err.number <> 0 Then
        PropertyVal = CVErr(Err.number)
        Err.Clear
    ElseIf isObj Then
        Set PropertyVal = PropertyObj(propName).value
    Else
        PropertyVal = PropertyObj(propName).value
    End If
End Function

Private Function PopulateProps()
        On Error Resume Next
        Dim tVarType As VbVarType
        Dim tProp As Property
        
        Set lProps = Nothing
        Set lProps = New Dictionary
        lProps.CompareMode = TextCompare
        
        For Each tProp In lcomp.Properties
            tVarType = VarType(tProp.value)
            If Err.number = 0 And tProp.NumIndices <= 1 Then
                lProps.add tProp.Name, tProp.value
            ElseIf Err.number = 0 Then
                lProps.add tProp.Name, tProp.value
            Else
                Beep
                Err.Clear
            End If
        Next tProp
End Function





' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''      PRIVATE IMPLEMENTATION OF EnumCompare
''      See: '  author (c) Paul Brower https://github.com/lopperman/just-VBA
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function privEnumCompare(theEnum As Variant, enumMember As Variant, Optional ByVal iType As ecComparisonTypePriv = ecComparisonTypePriv.ecOR) As Boolean
    Dim c As Long
    c = theEnum And enumMember
    privEnumCompare = IIf(iType = ecOR, c <> 0, c = enumMember)
End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Private IMplentation of EnumModify
''   FLAG ENUM - ADD/REMOVE SPECIFIC ENUM MEMBER
''   (Works with any flag enum)
''   e.g. If you have vbMsgBoxStyle enum and want to make sure
''   'DefaultButton1' is included
''   msgBtnOption = vbYesNo + vbQuestion
''   msgBtnOption = EnumModify(msgBtnOption,vbDefaultButton1,feVerifyEnumExists)
''   'now includes vbDefaultButton1, would not modify enum value if it already
''   contained vbDefaultButton1
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function privEnumModify(theEnum, enumMember, modifyType As FlagEnumModifyPriv) As Long
        Dim Exists As Boolean
        Exists = privEnumCompare(theEnum, enumMember)
        If Exists And modifyType = FlagEnumModifyPriv.feVerifyEnumRemoved Then
            theEnum = theEnum - enumMember
        ElseIf Exists = False And modifyType = FlagEnumModifyPriv.feVerifyEnumExists Then
            theEnum = theEnum + enumMember
        End If
        privEnumModify = theEnum
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Private (class) version of the StringsMatch Function available in
'   some of the common modules at:
'     https://github.com/lopperman/just-VBA
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function StringsMatchPriv( _
        ByVal checkString As Variant, ByVal _
        validString As Variant, _
        Optional smEnum As strMatchEnumPriv = strMatchEnumPriv.smEqual, _
        Optional compMethod As VbCompareMethod = vbTextCompare) As Boolean
        
        Dim str1, str2
        str1 = CStr(checkString)
        str2 = CStr(validString)
        Select Case smEnum
            Case strMatchEnumPriv.smEqual
                StringsMatchPriv = StrComp(str1, str2, compMethod) = 0
            Case strMatchEnumPriv.smNotEqualTo
                StringsMatchPriv = StrComp(str1, str2, compMethod) <> 0
            Case strMatchEnumPriv.smContains
                StringsMatchPriv = InStr(1, str1, str2, compMethod) > 0
            Case strMatchEnumPriv.smStartsWithStr
                StringsMatchPriv = InStr(1, str1, str2, compMethod) = 1
            Case strMatchEnumPriv.smEndWithStr
                If Len(str2) > Len(str1) Then
                    StringsMatchPriv = False
                Else
                    StringsMatchPriv = InStr(Len(str1) - Len(str2) + 1, str1, str2, compMethod) = Len(str1) - Len(str2) + 1
                End If
        End Select
    End Function
    
    Private Function CollectionItemExists(ByRef col As Collection, ByVal item) As Boolean
    On Error Resume Next
        Dim colItem As Variant
        For Each colItem In col
            If StringsMatchPriv(colItem, item) Then
                CollectionItemExists = True
                Exit For
            End If
        Next
    End Function

