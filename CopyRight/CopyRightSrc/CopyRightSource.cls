VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CopyRightSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' NOTE: SET THESE ITEMS BEFORE FIRST USE:
'       CLOSE_WAIT_SCREEN_METHOD (Private Constant)
'       PROTECT_ACTIVE_SHEET_METHOD (Private Constant)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' CopyRight v1.0.0
' (c) Paul Brower - https://github.com/lopperman/just-VBA
'
' Object Used to Create Data/Formats from a SOURCE to a TARGET
'
' @classmodule CopyRight
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1

Private Const CLASS_NAME As String = "CopyRightSource"

Private l_lstObj As ListObject
Private l_sourceType As crSourceType
Private l_sources As Collection

Public Property Get SourceType() As crSourceType
    SourceType = l_sourceType
End Property

' ~~~ ~~~ COPYRIGHT SOURCE - RANGE ~~~ ~~~
Public Function AddRange(allowMultipleAreas As Boolean, ParamArray ranges() As Variant)
'   Add Source/sTo Be 'Copied' (Must be same Type)
    If Not CanAddSource(IIf(allowMultipleAreas, crSourceType.crRangeMultipleAreas, crSourceType.crRangeSingleArea)) Then
        Err.Raise ERR_COPY_RIGHT_SOURCE_TYPE_CONFLICT, Description:=CLASS_NAME & ".AddRange - Source Type Conflict - All Sources must have same crSourceType"
    End If

    Dim i As Long
    For i = LBound(ranges) To UBound(ranges)
        If Not TypeName(ranges(i)) = "Range" Then
            Err.Raise ERR_COPY_RIGHT_SOURCE_NOT_VALID, Description:=CLASS_NAME & ".AddRange - ranges(" & i & ") object must be type 'Range'"
        Else
            Dim tstRng As Range
            Set tstRng = ranges(i)
            If allowMultipleAreas = False And tstRng.Areas.Count > 1 Then
                Err.Raise ERR_COPY_RIGHT_SOURCE_TYPE_CONFLICT, Description:=CLASS_NAME & ".AddRange - Cannot Add Range with Multiple Areas when 'allowMultipleAreas' set to False"
            End If
            If l_sourceType = [_Unknown] Then
                l_sourceType = IIf(allowMultipleAreas, crSourceType.crRangeMultipleAreas, crSourceType.crRangeSingleArea)
            End If
            Dim areaIDX As Long
            For areaIDX = 1 To tstRng.Areas.Count
                l_sources.Add tstRng.Areas(areaIDX)
            Next areaIDX
        End If
    Next i
End Function

' ~~~ ~~~ COPYRIGHT SOURCE - LIST OBJECT ~~~ ~~~
Public Function AddListObject(lstObj As ListObject)
'   Add Source ListObject
'   Must be exclusive Source
    If l_sources.Count > 0 Then
        Err.Raise ERR_COPY_RIGHT_SOURCE_TYPE_CONFLICT, Description:=CLASS_NAME & ".AddListObject - Source Type Conflict - Previous Source exists."
    End If
    l_sourceType = crListObject
    l_sources.Add lstObj
End Function

' ~~~ ~~~ COPYRIGHT SOURCE - LIST OBJECT COLUMN(S) ~~~ ~~~
Public Function AddListObjectCols(lstObj As ListObject, ParamArray columns() As Variant)
'   'columns can be column number,column name, or actual ListColumn object
'   Add Source ListObject Columns
'   Set l_LstObj for verifying other list obj colums
'   Must all be from same list object
    If Not CanAddSource(crSourceType.crListObjectColumns) Then
        Err.Raise ERR_COPY_RIGHT_SOURCE_TYPE_CONFLICT, Description:=CLASS_NAME & ".AddListObjectCols - Source Type Conflict - All Sources must have same crSourceType"
    End If
        
    If l_lstObj Is Nothing Then
        Set l_lstObj = lstObj
    End If
    If l_sourceType = [_Unknown] Then
        l_sourceType = crListObjectColumns
    End If
         
    If Not lstObj Is l_lstObj Then
        Err.Raise ERR_COPY_RIGHT_SOURCE_TYPE_CONFLICT, Description:=CLASS_NAME & ".AddListObjectCols - Cannot add columns from separate ListObjects "
    End If
         
    Dim colIdx As Long, tmpLstCol As listColumn, lcAdded As Boolean
    For colIdx = LBound(columns) To UBound(columns)
        Set tmpLstCol = Nothing
        lcAdded = False
        If IsObject(columns(colIdx)) Then
            If TypeName(columns(colIdx)) = "ListColumn" Then
                Set tmpLstCol = columns(colIdx)
            End If
        ElseIf TypeName(columns(colIdx)) = "String" Then
            If pbListObj.ListColumnExists(lstObj, CStr(columns(colIdx))) Then
                Set tmpLstCol = lstObj.ListColumns(CStr(columns(colIdx)))
            End If
        ElseIf IsNumeric(columns(colIdx)) Then
            If lstObj.ListColumns.Count >= CLng(columns(colIdx)) Then
                Set tmpLstCol = lstObj.ListColumns(CLng(columns(colIdx)))
            End If
        End If
        If tmpLstCol Is Nothing Then
            Err.Raise ERR_COPY_RIGHT_SOURCE_NOT_VALID, Description:=CLASS_NAME & ".AddListObjectCols - Unable to validate ListColumn"
        End If
        If Not tmpLstCol.Parent Is l_lstObj Then
            Err.Raise ERR_COPY_RIGHT_SOURCE_TYPE_CONFLICT, Description:=CLASS_NAME & ".AddListObjectCols - Cannot add columns from separate ListObjects "
        End If
        If l_sources.Count > 0 Then
            Dim i As Long
            For i = 1 To l_sources.Count
                If tmpLstCol Is l_sources(i) Then
                    lcAdded = True
                    Exit For
                End If
            Next i
        End If
        If Not lcAdded Then
            l_sources.Add tmpLstCol
        End If
    Next colIdx
    
    l_sourceType = crListObject
    l_sources.Add lstObj
    
End Function

Public Property Get Sources() As Collection
    Set Sources = l_sources
End Property

Public Property Get ListObjectForColumns() As ListObject
    'Only Valid if Source Type is ListObjCols
    If l_sourceType = crListObjectColumns Then
        Set ListObjectForColumns = l_lstObj
    End If
End Property

Public Property Get SourceCount() As Long
    SourceCount = l_sources.Count
End Property

Private Function CanAddSource(SourceType As crSourceType) As Boolean
    If l_sourceType = [_Unknown] Then
        CanAddSource = True
    Else
        CanAddSource = SourceType = l_sourceType
    End If
End Function

Private Sub Class_Initialize()
    Set l_sources = New Collection
    
End Sub
