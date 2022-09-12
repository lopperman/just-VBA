VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbAutoFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit
    Option Compare Text
    Option Base 1
    
    Private colParams As Collection
    Private searchCol As Variant
    Private crit1 As Variant
    Private operator As XlFormatConditionOperator
    Private crit2 As Variant
    
    Public Function AddParam(ByVal searchCol As Variant, _
        ByVal crit1 As Variant, _
        Optional ByVal operator As XlFormatConditionOperator = XlFormatConditionOperator.xlEqual, _
        Optional crit2 As Variant)
        colParams.Add Array(searchCol, crit1, operator, crit2)
    End Function
    Private Function GetSearchCol(idx As Long) As Variant
        GetSearchCol = colParams(idx)(1)
    End Function
    Private Function GetCrit1(idx As Long) As Variant
        GetCrit1 = colParams(idx)(2)
    End Function
    Private Function GetOperator(idx As Long) As XlFormatConditionOperator
        GetOperator = colParams(idx)(3)
    End Function
    Private Function GetCrit2(idx As Long) As Variant
        GetCrit2 = colParams(idx)(4)
    End Function
    Public Property Get Count() As Long
        Count = colParams.Count
    End Property
    
    Public Function Execute(dataRng As Range)
        'Remove Any Existing Filters
        If dataRng.Worksheet.FilterMode Then dataRng.Worksheet.ShowAllData
        Dim srchV As Variant, pIdx As Long, updCrit1 As String, updCrit2 As String, isMult As Boolean
        Dim evts As Boolean: evts = Application.EnableEvents: Application.EnableEvents = False
        Dim scrn As Boolean: scrn = Application.ScreenUpdating: Application.ScreenUpdating = False
        
        With dataRng
            For pIdx = 1 To colParams.Count
                Select Case GetOperator(pIdx)
                    Case XlFormatConditionOperator.xlEqual
                        updCrit1 = "=" & GetCrit1(pIdx)
                    Case XlFormatConditionOperator.xlGreater
                        updCrit1 = ">" & GetCrit1(pIdx)
                    Case XlFormatConditionOperator.xlGreaterEqual
                        updCrit1 = ">=" & GetCrit1(pIdx)
                    Case XlFormatConditionOperator.xlLess
                        updCrit1 = "<" & GetCrit1(pIdx)
                    Case XlFormatConditionOperator.xlLessEqual
                        updCrit1 = "<=" & GetCrit1(pIdx)
                    Case XlFormatConditionOperator.xlNotEqual
                        updCrit1 = "<>" & GetCrit1(pIdx)
                    Case XlFormatConditionOperator.xlBetween
                        isMult = True
                        updCrit1 = ">=" & GetCrit1(pIdx)
                        updCrit2 = "<=" & GetCrit2(pIdx)
                    Case XlFormatConditionOperator.xlNotBetween
                        isMult = True
                        updCrit1 = "<" & GetCrit1(pIdx)
                        updCrit2 = ">" & GetCrit2(pIdx)
                End Select
                If Not isMult Then
                    .AutoFilter field:=GetSearchCol(pIdx), Criteria1:=updCrit1
                Else
                    .AutoFilter field:=GetSearchCol(pIdx), Criteria1:=updCrit1, operator:=xlAnd, Criteria2:=updCrit2
                End If
            Next pIdx
        End With
        
        Application.ScreenUpdating = scrn
        Application.EnableEvents = evts
    
    End Function

    Private Sub Class_Initialize()
        Set colParams = New Collection
    End Sub

