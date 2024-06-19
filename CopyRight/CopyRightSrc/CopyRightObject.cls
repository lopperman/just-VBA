VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CopyRightObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' CopyRightObject v1.0.0
' (c) Paul Brower - https://github.com/lopperman/just-VBA
'
' Class for storing information about Source or Target Object about to be copied (Range, List Object, etc)
'
' @classmodule CopyRightObject
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1

Private Const CLASS_NAME As String = "CopyRightObj"

Private l_lstObj As ListObject
Private l_ObjType As crObjectType
Private l_Objects As Collection

Public Property Get ObjectType() As crObjectType
    ObjectType = l_ObjType
End Property

' ~~~ ~~~ COPYRIGHT OBJECT - RANGE ~~~ ~~~
Public Function AddRange(allowMultipleAreas As Boolean, ParamArray ranges() As Variant)
'   Add Object/sTo Be 'Copied' (Must be same Type)
    If Not CanAdd(IIf(allowMultipleAreas, crObjectType.crRangeMultipleAreas, crObjectType.crRangeSingleArea)) Then
        Err.Raise ERR_COPY_RIGHT_OBJECT_TYPE_CONFLICT, Description:=CLASS_NAME & ".AddRange - Source Type Conflict - All Sources must have same crObjectType"
    End If

    Dim i As Long
    For i = LBound(ranges) To UBound(ranges)
        If Not TypeName(ranges(i)) = "Range" Then
            Err.Raise ERR_COPY_RIGHT_OBJECT_NOT_VALID, Description:=CLASS_NAME & ".AddRange - ranges(" & i & ") object must be type 'Range'"
        Else
            Dim tstRng As Range
            Set tstRng = ranges(i)
            If allowMultipleAreas = False And tstRng.Areas.Count > 1 Then
                Err.Raise ERR_COPY_RIGHT_OBJECT_TYPE_CONFLICT, Description:=CLASS_NAME & ".AddRange - Cannot Add Range with Multiple Areas when 'allowMultipleAreas' set to False"
            End If
            If l_ObjType = [_Unknown] Then
                l_ObjType = IIf(allowMultipleAreas, crObjectType.crRangeMultipleAreas, crObjectType.crRangeSingleArea)
            End If
            Dim areaIDX As Long
            For areaIDX = 1 To tstRng.Areas.Count
                l_Objects.Add tstRng.Areas(areaIDX)
            Next areaIDX
        End If
    Next i
End Function

' ~~~ ~~~ COPYRIGHT SOURCE - LIST OBJECT ~~~ ~~~
Public Function AddListObject(lstObj As ListObject)
'   Add Source ListObject
'   Must be exclusive Source
    If l_Objects.Count > 0 Then
        Err.Raise ERR_COPY_RIGHT_OBJECT_TYPE_CONFLICT, Description:=CLASS_NAME & ".AddListObject - Object Type Conflict - Previous Object exists."
    End If
    l_ObjType = crListObject
    l_Objects.Add lstObj
End Function

' ~~~ ~~~ COPYRIGHT SOURCE - LIST OBJECT COLUMN(S) ~~~ ~~~
Public Function AddListObjectCols(lstObj As ListObject, ParamArray columns() As Variant)
'   'columns can be column number,column name, or actual ListColumn object
'   Add ListObject Columns
'   Set l_LstObj for verifying other list obj colums
'   Must all be from same list object
    If Not CanAdd(crObjectType.crListObjectColumns) Then
        Err.Raise ERR_COPY_RIGHT_OBJECT_TYPE_CONFLICT, Description:=CLASS_NAME & ".AddListObjectCols - Object Type Conflict - All Objects must have same crObjectType"
    End If
        
    If l_lstObj Is Nothing Then
        Set l_lstObj = lstObj
    End If
    If l_ObjType = [_Unknown] Then
        l_ObjType = crListObjectColumns
    End If
         
    If Not lstObj Is l_lstObj Then
        Err.Raise ERR_COPY_RIGHT_OBJECT_TYPE_CONFLICT, Description:=CLASS_NAME & ".AddListObjectCols - Cannot add columns from separate ListObjects "
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
            Err.Raise ERR_COPY_RIGHT_OBJECT_NOT_VALID, Description:=CLASS_NAME & ".AddListObjectCols - Unable to validate ListColumn"
        End If
        If Not tmpLstCol.Parent Is l_lstObj Then
            Err.Raise ERR_COPY_RIGHT_OBJECT_TYPE_CONFLICT, Description:=CLASS_NAME & ".AddListObjectCols - Cannot add columns from separate ListObjects "
        End If
        If l_Objects.Count > 0 Then
            Dim i As Long
            For i = 1 To l_Objects.Count
                If tmpLstCol Is l_Objects(i) Then
                    lcAdded = True
                    Exit For
                End If
            Next i
        End If
        If Not lcAdded Then
            l_Objects.Add tmpLstCol
        End If
    Next colIdx
    
    l_ObjType = crListObject
    l_Objects.Add lstObj
    
End Function

Public Property Get Objects() As Collection
    Set Objects = l_Objects
End Property

Public Property Get ListObjectForColumns() As ListObject
    'Only Valid if Object Type is ListObjCols
    If l_ObjType = crListObjectColumns Then
        Set ListObjectForColumns = l_lstObj
    End If
End Property

Public Property Get ObjectCount() As Long
    ObjectCount = l_Objects.Count
End Property

Private Function CanAdd(objType As crObjectType) As Boolean
    If l_ObjType = [_Unknown] Then
        CanAdd = True
    Else
        CanAdd = objType = l_ObjType
    End If
End Function

Private Sub Class_Initialize()
    Set l_Objects = New Collection
    
End Sub
