Attribute VB_Name = "pbError"
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' NOTE: SET THESE ITEMS BEFORE FIRST USE:
'       CLOSE_WAIT_SCREEN_METHOD (Private Constant)
'       PROTECT_ACTIVE_SHEET_METHOD (Private Constant)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' pbError v1.0.0
' (c) Paul Brower - https://github.com/lopperman/VBA-pbUtil
'
' General  Global Error Handler and Error Handling Utilities
'
' @module pbError
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Compare Text
Option Base 1

'   *** GENERIC ERRORS ONLY
'   *** STARTS AT 1001 - (513-1000 RESERVED FOR APPLICATION SPECIFIC ERRORS)

Public Const ERR_OBSOLETE As Long = vbObjectError + 1001
Public Const ERR_MISSING_EXPECTED_WORKSHEET As Long = vbObjectError + 1002
Public Const ERR_HOME_SHEET_NOT_SET As Long = vbObjectError + 1003
Public Const ERR_HOME_SHEET_ALREADY_SET As Long = vbObjectError + 1004
Public Const ERR_INVALID_CALLER_SOURCE As Long = vbObjectError + 1005
Public Const ERR_ACTION_PATH_NOT_DEFINED As Long = vbObjectError + 1006
Public Const ERR_UNABLE_TO_GET_WORKSHEET_BUTTON As Long = vbObjectError + 1007
Public Const ERR_RANGE_SRC_TARGET_MISMATCH = vbObjectError + 1008
Public Const ERR_INVALID_FT_ARGUMENTS = vbObjectError + 1009
Public Const ERR_NO_ERROR As String = "(INFO)"
Public Const ERR_ERROR As String = "(ERROR)"
Public Const ERR_INVALID_SETTING_OPERATION As Long = vbObjectError + 1010
Public Const ERR_TARGET_OBJECT_MUST_BE_EMPTY As Long = vbObjectError + 1011
Public Const ERR_EXPECTED_SHEET_NOT_ACTIVE As Long = vbObjectError + 1012
Public Const ERR_INVALID_RANGE_SIZE = vbObjectError + 1013
Public Const ERR_IMPORT_COLUMN_ALREADY_MATCHED = vbObjectError + 1014
Public Const ERR_RANGE_AREA_COUNT = vbObjectError + 1015
Public Const ERR_LIST_OBJECT_RESIZE_CANNOT_DELETE = vbObjectError + 1016
Public Const ERR_LIST_OBJECT_RESIZE_INVALID_ARGUMENTS = vbObjectError + 1017
Public Const ERR_INVALID_ARRAY_SIZE = vbObjectError + 1018
Public Const ERR_EXPECTED_MANUAL_CALC_NOT_FOUND = vbObjectError + 1019
Public Const ERR_WORKSHEET_OBJECT_NOT_FOUND = vbObjectError + 1020
Public Const ERR_NOT_IMPLEMENTED_YET = vbObjectError + 1021
'CLASS INSTANCE PROHIBITED, FOR CLASS MODULE WITH Attribute VB_PredeclaredId = True
    Public Const ERR_CLASS_INSTANCE_PROHIBITED = vbObjectError + 1022
Public Const ERR_CONTROL_STATE = vbObjectError + 1023
Public Const ERR_PREVIOUS_PerfState_EXISTS = vbObjectError + 1024
Public Const ERR_FORCED_WORKSHEET_NOT_FOUND = vbObjectError + 1025
Public Const ERR_CANNOT_CHANGE_PACKAGE_PROPERTY = vbObjectError + 1026
Public Const ERR_INVALID_PACKAGE_OPERATION = vbObjectError + 1027

'If you have a method to close a 'wait' screen, put the name here, otherwise should be ""
'Private Const CLOSE_WAIT_SCREEN_METHOD As String = "CloseBusy"
'If you have a method to protect active sheet, put the name here, otherwise should be ""
'Private Const PROTECT_ACTIVE_SHEET_METHOD As String = "ProtectActiveSheet"

Private l_lastError As Date
Private l_currentErrCount As Long
Private l_totalErrrorCount As Long
Private l_IsDeveloper As Boolean

Public Enum ErrorOptionsEnum
    ftDefaults = 2 ^ 0
    ftERR_ProtectSheet = 2 ^ 1
    ftERR_MessageIgnore = 2 ^ 2
    ftERR_NoBeeper = 2 ^ 3
    ftERR_DoNotCloseBusy = 2 ^ 4
    ftERR_ResponseAllowBreak = 2 ^ 5
End Enum

Public Property Get IsDeveloper() As Boolean
    IsDeveloper = l_IsDeveloper
End Property
Public Property Let IsDeveloper(devMode As Boolean)
    l_IsDeveloper = devMode
End Property


Public Function ErrorCheck(Optional Source As String, Optional options As ErrorOptionsEnum, Optional customErrorMsg As String) As Long
    Dim errNumber As Long, errDESC As Variant, errorInfo As String, ignoreError As Boolean, errERL As Long
    errNumber = Err.number
    errDESC = Err.Description
    errERL = Erl
    errorInfo = ErrString(customSrc:=Source, errNUM:=errNumber, errDESC:=errDESC, errERL:=errERL)
    If errNumber = 0 Then Exit Function
    LogError Concat("pbError.ErrorCheck: ", errorInfo)
    
    If ThisWorkbook.ReadOnly Then
        LogError "Workbook is READ-ONLY - Closing NOW"
        pbPerf.RestoreDefaultAppSettingsOnly
        ThisWorkbook.Close SaveChanges:=False
        Exit Function
    End If
    
    ftBeep btError
    
    If options = 0 Then options = ftDefaults
    On Error GoTo -1
    On Error Resume Next
    
    
    Dim cancelMsg As String
    Dim okONLY As Boolean
    okONLY = True
    If (IsDEV And DebugMode) Or EnumCompare(options, ftERR_ResponseAllowBreak) Then
        okONLY = False
    End If
    If okONLY Then
        cancelMsg = vbCrLf & "** PRESS OK TO CONTINUE"
    Else
        cancelMsg = vbCrLf & "** PRESS OK TO CONTINUE CODE EXECUTION -- WHICH IS NORMALLY WHAT YOU WANT TO DO.  'CANCEL'  WILL STOP CODE FROM RUNNING, AND SHOULD ONLY BE USED IF YOU'RE CONTINUALLY SEEING ERROR MESSAGES"
    End If
    
    Dim EventsOn As Boolean: EventsOn = Application.EnableEvents
    Application.EnableEvents = False
    
    '***** ftErr_NoBeeper
    If ignoreError = False And EnumCompare(options, ftERR_NoBeeper) = False Then Beeper
    
    If (customErrorMsg & "") <> vbNullString Then
        errorInfo = errorInfo & vbCrLf & customErrorMsg
    End If
    
    '*****ftERR_MessageIgnore
    If ignoreError = False And EnumCompare(options, ftERR_MessageIgnore) = False Then
        If okONLY = False Then
            If MsgBox_FT(CStr(errorInfo) & cancelMsg, vbOKCancel + vbCritical + vbSystemModal + vbDefaultButton1, "ERROR LOGGED TO LOG FILE") = vbCancel Then
                If Err.number <> 0 Then Err.Clear
                FatalEnd
                Exit Function
            End If
        Else
            If IsDEV Then
                If MsgBox_FT(CStr(errorInfo) & cancelMsg, vbOKCancel + vbSystemModal, "DEV MSG: CANCEL TO STOP CODE") = vbCancel Then
                    Beep
                    DoEvents
                    Stop
                End If
            Else
                MsgBox_FT CStr(errorInfo) & cancelMsg, vbOKOnly + vbError, "ERROR LOGGED TO LOG FILE"
            End If
            If Err.number <> 0 Then Err.Clear
            Exit Function
        End If
    End If
    
    If errNumber <> 0 Then Err.Clear
    If errNumber = 18 Then
        ResetUI
        MsgBox_FT "User Cancelled Current Process", title:="Cancelled"
        End
    End If
    
    '*****ftERR_ProtectSheet
    If ignoreError = False And EnumCompare(options, ftERR_ProtectSheet) Then
        ProtectSht ThisWorkbook.ActiveSheet
    End If
    If Err.number <> 0 Then
        ErrorCheck
    End If
    If ignoreError = False Then
        ErrorCheck = errNumber
    Else
        ErrorCheck = 0
    End If
    If Err.number <> 0 Then Err.Clear
End Function

Public Property Let LastErrorDt(dtTm As Variant)
    l_lastError = dtTm
    l_currentErrCount = l_currentErrCount + 1
    l_totalErrrorCount = l_totalErrrorCount + 1
End Property


Public Property Get LastErrorDt()
    LastErrorDt = l_lastError
End Property

Public Property Get TotalErrorCount() As Long
    TotalErrorCount = l_totalErrrorCount
End Property

'   ~~~ ~~~ COUNT OF ERRORS RAISED  ~~~ ~~~
Public Property Get CurrentErrorCount() As Long
    CurrentErrorCount = l_currentErrCount
End Property

'   ~~~ ~~~ FORMAT TIMESTAMPT FOR PRE-PRENDING TO LOG MESSAGES ~~~ ~~~
Private Function TimeStamp() As String
    TimeStamp = Format(Now, "mm/dd/yyyy hh:mm:ss")
End Function

'   ~~~ ~~~ RAISE ERROR ~~~ ~~~
Public Sub RaiseError(errNumber As Long, Optional errorDesc As String = vbNullString, Optional resetUserInterface As Boolean = True)
    LogError Concat("RaiseError Called: ", errNumber, " - errorDesc: ", errorDesc)
    If resetUserInterface Then
        ResetUI
    End If
    If Len(errorDesc) > 0 Then
        Err.Raise errNumber, Description:=errorDesc
    Else
        Err.Raise errNumber
    End If
End Sub

Public Property Get ErrString(Optional customSrc As String, Optional errNUM As Variant, Optional errDESC As Variant, Optional errERL As Variant) As String
'   Format Known Error Information

    If IsMissing(errNUM) Then errNUM = Err.number
    If IsMissing(errDESC) Then errDESC = Err.Description
    If IsMissing(errERL) Then errERL = Erl
    If errNUM <> 0 Then
        Dim msg As String
        msg = "ERROR: " & errNUM & ", Desc: " & errDESC & ", Src: " & Err.Source
        If errERL <> 0 Then msg = msg & " (ERL: " & errERL & ") "
        If Len(customSrc & "") > 0 Then
            msg = Concat(customSrc, vbCrLf, msg)
        End If
        ErrString = msg
    End If

End Property

Private Function ResetUI()
        
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.Interactive = True
        Application.Cursor = xlDefault
        Application.Calculation = xlCalculationAutomatic
        Application.EnableAnimations = True
        Application.EnableMacroAnimations = True
        
End Function


Public Function Beeper()
'   Enable control over beeps -- put that here
    Beep

End Function


Public Function FatalEnd()
On Error Resume Next
         
        If IsDEV Then
            Beep
            Stop
            Exit Function
        End If
         
        ResetUI
                
        'ADD ANY ADDITIONAL CUSTOM CODE BEFORE KILLING THINGS
        'LIKE GO TO SPECIFIC WORKSHEET
    
        Beep
        MsgBox_FT "All running code has been terminated.  Please completely quit excel and then re-open."
        
        ' *** END VBE ENGINE ***
        End
        ' ***
End Function
