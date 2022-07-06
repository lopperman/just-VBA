Attribute VB_Name = "CopyRightCommon"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' NOTE: THIS CLASS REQUIRES VB_PredeclaredId = True
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' pbCopyRight v1.0.0
' (c) Paul Brower - https://github.com/lopperman/just-VBA
'
'   Copy & "Paste" Without the .Paste.   EVER.
'   ** This is the Way **
'
' @module pbCopyRight
' @author Paul Brower (lopperman)
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1

Public Const ERR_CANNOT_ACCESS_PROTECTED_OBJECT As Long = vbObjectError + 1028
Public Const ERR_COPY_RIGHT_OBJECT_TYPE_CONFLICT As Long = vbObjectError + 1029
Public Const ERR_COPY_RIGHT_OBJECT_NOT_VALID As Long = vbObjectError + 1030

Public Enum crObjectType
    [_Unknown] = 0
    crRangeSingleArea
    crRangeMultipleAreas
    crListObject
    crListObjectColumns
End Enum

Public Enum crRangeAreas
    crRetainGaps
    crStackVerticalAll
    crStackHorizontalAll
    crStackVerticalOnSheetChange
    crStackHorizontalOnSheetChange
End Enum

Public Enum crCopyOptions
    coFormulasToValues = 2 ^ 0
    coKeepFormulas = 2 ^ 1
    coVisibleOnly = 2 ^ 2
    coIncludeListObjectHeaders = 2 ^ 3
End Enum

'Public Enum CopyOptions
'    [_coError] = 0
'    'Modifies What's Being Copied
'    coFormulas = 2 ^ 0
'    coVisibleCellsOnly = 2 ^ 1
'    coUniqueRows = 2 ^ 2
'    coUniqueCols = 2 ^ 3
'
'    'Modifies Target Structure
'    coIncludeListObjHeaders = 2 ^ 4 'Valid LstObj, and LstObjCols only
'    coCreateListObj = 2 ^ 5
'    coPullRowsTogether = 2 ^ 6 'Only Valid Range w/multiple disparate areas
'    coPullColsTogether = 2 ^ 7 'Only ValidRange w/multiple disparate areas, OR LstCols with disparate cols
'
'    'Modifies Format
'    coMatchFontStyle = 2 ^ 8
'    coMatchInterior = 2 ^ 9
'    coMatchRowColSize = 2 ^ 10
'    coMatchMergeAreas = 2 ^ 11
'    coMatchLockedCells = 2 ^ 12
'
'    coDROPUnmatchedLstObjCols = 2 ^ 13
'    coClearTargetLstObj = 2 ^ 14
'    coManualLstObjMap = 2 ^ 15
'
'    'Create Destination
'    coNewWorkbook = 2 ^ 16
'End Enum
'Public Enum CopyTo
'    ftRange
'    ftListObj
'    ftListObjCols
'    toNewWorksheet
'    toNewWorkbook
'End Enum

' ~~~ ~~~ ~~~   ALL POSSIBLE COMBINATIONS ~~~ ~~~ ~~~

'   CONTIGUOUS RANGE -
'   01 - Contiguous Range to Contiguous Range - Exist Sheet
'   02 - Contiguous Range to Contiguous Range - New Sheet (Optionally New Workbook)
'   03 - Contiguous Range to Auto Matched List Obj Cols (Clear, Don't Clear)
'   04 - Contiguous Range to Manual Matched List Obj Cols (Clear, Don't Clear)

'   05  - Multiples Range Areas with Matching Rows or Cols to Continuous Area - Exist or New Sheet /Workbook
'   06  - Multiples Range Areas with Matching Rows or Cols to Continuous Area - Exist or New Sheet /Workbook


Public Function pbCopyListObj(sourceObj As ListObject, options As CopyOptions)

End Function
Public Function pbCopyListObjCols(sourceObj As ListObject, options As CopyOptions)

End Function
Public Function pbCopyRange(sourceRng As Range, options As CopyOptions)

End Function







' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
'                    Helpers
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~



Private Function RestoreUICR()
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.Interactive = True
        Application.Cursor = xlDefault
        Application.Calculation = xlCalculationAutomatic
        Application.EnableAnimations = True
        Application.EnableMacroAnimations = True
End Function
Private Function PauseUICR(Optional calcMode As XlCalculation = XlCalculation.xlCalculationManual)
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        Application.Interactive = False
        Application.Cursor = xlWait
        Application.Calculation = calcMode
        Application.EnableAnimations = False
        Application.EnableMacroAnimations = False
End Function

