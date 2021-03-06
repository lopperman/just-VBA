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
Private Const CLOSE_WAIT_SCREEN_METHOD As String = "CloseBusy"
'If you have a method to protect active sheet, put the name here, otherwise should be ""
Private Const PROTECT_ACTIVE_SHEET_METHOD As String = "ProtectActiveSheet"

Private l_lastError As Date
Private l_currentErrCount As Long
Private l_totalErrrorCount As Long

Public Enum ErrorOptionsEnum
    ftDefaults = 2 ^ 0
    ftERR_ControlStateClear = 2 ^ 1
    ftERR_ProtectSheet = 2 ^ 2
    ftERR_MessageIgnore = 2 ^ 3
    ftERR_NoBeeper = 2 ^ 4
    ftERR_IGNORE_TREAT_AS_NO_ERR = 2 ^ 5
    ftERR_DoNotCloseBusy = 2 ^ 6
End Enum

Public Function ErrorCheck(Optional Source As String, Optional options As ErrorOptionsEnum, Optional customErrorMsg As String) As Long
    Dim errNumber As Long, errDESC As Variant, errorInfo As String, ignoreError As Boolean, errERL As Long
    errNumber = Err.Number
    errDESC = Err.Description
    errERL = Erl
    errorInfo = ErrString(customSrc:=Source, errNUM:=errNumber, errDESC:=errDESC, errERL:=errERL)
    
    If errNumber = 0 Then Exit Function
    
    ' ~~~ ~~~ FOR DEVELOPER ONLY ~~~ ~~~
    If IsDEV Then
        If DebugMode Then
            DebugPrint errorInfo
            Beep
            'Comment / Uncomment The 'STOP and  'EXIT FUNCTION lines as needed
            'Stop
            'Exit Function
        End If
    End If
    
    ' ~~~ ~~~ ~~~ ~~~ WHAT TO DO IF ERROR 51 (INTERNAL ERROR) ~~~ ~~~ ~~~ ~~~
    '   If you don't want to go to 'FatalEnd', d
    If errNumber = 51 Then
        Dim msgFO As String
        msgFO = "An INTERNAL ERROR (Error 51) has occured.  This error is not specific to the Financial Tool, is is a general EXCEL Error." & vbNewLine
        msgFO = msgFO & "The suggested course of action is to close ALL excel files and then re-open them. Please do that now to avoid any further issues." & vbCrLf
        msgFO = msgFO & "An attempt will be made now to save the Fin Tool, and a BYPASS will be created that should allow you to close the file without further problems."
        MsgBox_FT msgFO, vbOKOnly + vbCritical, "AN INTERNAL MS EXCEL ERROR HAS OCCURED!"
        
    End If
    
    If ThisWorkbook.readOnly Then
        'JUST IN CASE
        Exit Function
    End If
    
    If options = 0 Then options = ftDefaults
    On Error GoTo -1
    On Error Resume Next
    
    ignoreError = EnumCompare(options, ftERR_IGNORE_TREAT_AS_NO_ERR)
    If EnumCompare(options, ftERR_DoNotCloseBusy) = False Then
        If Not ignoreError And ThisWorkbook.ActiveSheet Is wsBusy Then
            If Len(CLOSE_WAIT_SCREEN_METHOD) > 0 Then Application.Run CLOSE_WAIT_SCREEN_METHOD
        End If
    End If
    
    Dim cancelMsg As String:
    cancelMsg = vbCrLf & "** PRESS OK TO CONTINUE CODE EXECUTION -- WHICH IS NORMALLY WHAT YOU WANT TO DO.  'CANCEL'  WILL STOP CODE FROM RUNNING, AND SHOULD ONLY BE USED IF YOU'RE CONTINUALLY SEEING ERROR MESSAGES"

    Dim EventsOn As Boolean: EventsOn = Application.EnableEvents
    Application.EnableEvents = False
    
    '***** ftErr_NoBeeper
    If ignoreError = False And EnumCompare(options, ftERR_NoBeeper) = False Then Beeper
    
    If (customErrorMsg & "") <> vbNullString Then
        errorInfo = errorInfo & vbCrLf & customErrorMsg
    End If
    
    '***** TRACE
    If DebugMode Or debugReleaseOverride Then
        DebugPrint errorInfo
    End If
    Trace errorInfo, True
        
    '*****ftERR_MessageIgnore
    If ignoreError = False And EnumCompare(options, ftERR_MessageIgnore) = False Then
        If MsgBox_FT(CStr(errorInfo) & cancelMsg, vbOKCancel + vbCritical + vbSystemModal + vbDefaultButton1, "ERROR LOGGED TO DEBUG SCREEN") = vbCancel Then
            If Err.Number <> 0 Then Err.Clear
            FatalEnd
            Exit Function
        End If
    End If

    If errNumber <> 0 Then Err.Clear
    
    If errNumber = 18 Then
        ResetUI
        MsgBox_FT "User Cancelled Current Process", title:="Cancelled"
        End
    End If
    
    '*****ftERR_ControlStateClear
    If ignoreError = False And EnumCompare(options, ftERR_ControlStateClear) Then
        'will also protect active sheet
        ResetUI
        If Len(PROTECT_ACTIVE_SHEET_METHOD) > 0 Then
            Application.Run PROTECT_ACTIVE_SHEET_METHOD
        End If
        If Err.Number <> 0 Then
            If MsgBox_FT("An error occured trying to restore the screen back to interactive mode." & cancelMsg, vbOKCancel + vbDefaultButton1, "ERROR") = vbCancel Then
                 
                FatalEnd
                Exit Function
            End If
            Err.Clear
        End If
    Else
        '*****ftERR_ProtectSheet
        If ignoreError = False And EnumCompare(options, ftERR_ProtectSheet) Then
            If Len(PROTECT_ACTIVE_SHEET_METHOD) > 0 Then
                Application.Run PROTECT_ACTIVE_SHEET_METHOD
            End If
        End If
        If Err.Number <> 0 Then
            ErrorCheck PROTECT_ACTIVE_SHEET_METHOD, ftERR_MessageIgnore
        End If
        'ResetUI forces events on -- otherwise only force them on if ForceEventsON
        If ignoreError = False And Events = False And EnumCompare(options, ftERR_ControlStateClear) Then
            Application.EnableEvents = True
        End If
    End If
    If ignoreError = False Then
        ErrorCheck = errNumber
    Else
        ErrorCheck = 0
    End If
    If Err.Number <> 0 Then Err.Clear
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

    If IsMissing(errNUM) Then errNUM = Err.Number
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
