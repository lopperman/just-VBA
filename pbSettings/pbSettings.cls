VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Manage Settings For Current Workbook and/or General Settings
''   for certain types of workbooks
''
''  * Supports configuration of settings for:
''      * Current Workbook, regardless of who is using it
''      * Unique Workbook Settings based on current user, regardless of
''         which platform is being used (Mac/PC)
''      * Unique Workbook Settings based on Platform (Mac/PC)
''      * Unique Workbook Settings based on current user and platform
''        (an example of a use case for this would be default zoom, or
''        workbook window size)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Online Documentation
''  https://github.com/lopperman/just-VBA/blob/main/pbSettings/pbSettings.md
''  Download pbSettings.cls Demo App
''  https://github.com/lopperman/just-VBA/raw/main/pbSettings/pbSettings_Example.xlsm
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  author (c) Paul Brower https://github.com/lopperman/just-VBA
''  class module pbSettings.cls
''  license GNU General Public License v3.0
''  Created Dec-2023
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  To Use:
''      01 - import pbSettings.cls into your Workbook project
''      02 - Copy the functions below into a new or existing a basic module:
''
''        Public Property Get pbSettingsReady() As Boolean
''            On Error Resume Next
''            If Not pbStg Is Nothing Then
''                pbSettingsReady = pbStg.ValidConfig
''            End If
''        End Property
''
''        Public Function pbpbStg.Setting() As pbSettings
''            On Error Resume Next
''            Static stgObj As pbSettings
''            If stgObj Is Nothing Then
''                Set stgObj = New pbSettings
''            End If
''            If Err.number = 0 Then
''                If Not stgObj Is Nothing Then
''                    If stgObj.ValidConfig Then
''                        Set pbStg = stgObj
''                    End If
''                End If
''            Else
''                Err.Clear
''            End If
''        End Function
''
''      03 - The line of code below will trigger the 'First Run' and set everything
''              up for pbSettings:
''
''           If pbSettingsReady = True Then
''               ' all action can now be used by calling: pbStg.[Method]
''           End if
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  THESE CONSTANTS DEFINE SETTINGS KEYS USED TO MANAGE THE
''  BEHAVIOR OF THE pbSettings Worksheet and ListObject
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Const PBSTG_HIDE As String = "PBSTG_AUTOHIDE"
    Private Const PBSTG_WS_CODENAME As String = "PBSTG_WS_CODE_NAME"
#If Mac Then
    Private Const PBSTG_OS As String = "MAC"
#Else
    Private Const PBSTG_OS As String = "PC"
#End If
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  allow workbook after save event to be captured
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private WithEvents stgwb As Workbook
Attribute stgwb.VB_VarHelpID = -1

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private Constants
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Const PBSETTINGS_TEMPLATE_NAME As String _
        = "pbSettingsTemplate.xlsx"
        
    Private Const PBSETTINGS_TEMPLATE_PATH As String _
        = "https://github.com/lopperman/just-VBA/raw/main/Templates/pbSettingsTemplate.xlsx"
        
    Private Const SETTINGS_SHEET_WEAKNAME As String _
        = "pbSettings"
        
    Private Const SETTINGS_SHEET_CODENAME As String _
        = "wsSTGWS"
        
    Private Const SETTINGS_LO As String _
        = "tbl_pbSettings"
        
    Private Const SETTINGS_COL_NAMES As String _
        = "SettingKey|SettingValue|SettingType|Updated"

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private variable used to store whether or not the pbSettings is
''      Valid and usable
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private l_validCfg As Boolean

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private variable used to store whether or not pbSettings has unsaved
''  changes
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private l_isDirty As Boolean

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  'l_ws' stores the reference to the pbSettings worksheet
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private l_ws As Worksheet

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  'l_lo' stores the reference to the pbSettings ListObject
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private l_lo As ListObject

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Part of the private implentation of the 'StringsMatch' function
''  See: https://www.reddit.com/r/vba/comments/w74j85/handy_stringsmatch_method_that_handles_equal_not/?utm_source=share&utm_medium=web2x&context=3
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Enum strMatchEnumPriv
        smEqual = 0
        smNotEqualTo = 1
        smContains = 2
        smStartsWithStr = 3
        smEndWithStr = 4
    End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Enum that can (and should) be  used when referring to a ListColumn
''  in the Settings List Object
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Enum SettingCols
        scSettingKey = 1
        scSettingValue = 2
        scSettingType = 3
        scUpdated = 4
    End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Value of 'TypeEnum' is stored in SettingCols.scSettingType (3) Column
''  Unless Specified When SETTING a setting value, will be 'teGeneral' (0)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Enum TypeEnum
        teGeneral = 0
        teNumeric = 1
        teBoolean = 2
        teDateTime = 3
    End Enum
    
    Private Enum XSearch
    '   DO NOT EDIT ENUM VALUES
    '   USED FOR SEARCH MODE ARGS ON 'MATCH' FUNCTION
        xsFirstToLast = 1
        xsLastToFirst = -1
        xsBinaryAsc = 2
        xsBinaryDesc = -2
    End Enum
    Private Enum XMatch
    '   DO NOT EDIT ENUM VALUES
    '   USED FOR MATCH MODE ARGS ON 'MATCH' FUN CTION
        exactMatch = 0
        ExactMatchOrNextSmaller = -1
        ExactMatchOrNextLarger = 1
        WildcardCharacterMatch = 2
    End Enum
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  pbSettings Events
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

Public Event OnDeleteSetting(ByVal settingKey, ByRef Cancel As Boolean)

    

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  MAIN EXPOSED SETTINGS METHODS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns True if [settingKey] Exists
''  Note:  A setting will be created automatically, if needed, when
''      it is set using: Setting([stgKey]) = [stgVal]
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function Exists(ByVal settingKey, Optional isOSSpecific As Boolean = False, Optional isUSERSpecific As Boolean = False) As Boolean
        If Not ValidConfig Then Exit Function
        Exists = Index(CheckKey(settingKey, isOSSpecific, isUSERSpecific)) > 0
    End Function
    
    Public Property Get IsDirty() As Boolean
        IsDirty = l_isDirty
    End Property
    
    Public Property Get AutoHide() As Boolean
        If Not ValidConfig Then Exit Property
        AutoHide = Setting(PBSTG_HIDE)
    End Property
    Public Property Let AutoHide(hideSheetDefault As Boolean)
        If Not ValidConfig Then Exit Property
        Setting(PBSTG_HIDE) = hideSheetDefault
        If hideSheetDefault = False And l_ws.visible <> xlSheetVisible Then
            ''  Disable App Events
            Dim evts As Boolean: evts = Application.EnableEvents: Application.EnableEvents = False
            l_ws.visible = xlSheetVisible
            l_lo.Range.EntireColumn.AutoFit
            ''  Restore App Events to prior setting
            Application.EnableEvents = evts
        Else
            CheckSettingsHealth
        End If
    End Property
    
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Gets [stgKey] value, if missing setting, create new setting with [stgDefaultVal]
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function SettingWithDefault(ByVal stgKey, Optional ByVal stgDefaultVal)
        If Not ValidConfig Then Exit Function
        If Not Exists(stgKey) Then
            If Not IsMissing(stgDefaultVal) Then
                Setting(stgKey) = stgDefaultVal
            End If
        End If
        SettingWithDefault = Setting(stgKey)
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Setting Value Property Getter
''  If [stgKey] does not exist, returns an Empty Type
''  e.g. If IsEmpty(Setting("invalidkey")) Then ....
''  If Setting SettingType is '3' (TypeEnum.teDateTime = 3), the value will
''    be converted using 'CDate' before it is returned
''  If Setting SettingType is '1' (TypeEnum.teNumeric = 1), the value will
''    be returned using:  [settingValue] = Val([settingValue])
''  If Setting SettingType is '2' (TypeEnum.teBoolean = 2), the value will
''      be return using: [settingValue] = CBool([settingValue])
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get Setting(ByVal stgKey)
        If Not ValidConfig Then Exit Property
        Dim tmpIdx As Long: tmpIdx = Index(stgKey)
        Dim resp, respType
        If tmpIdx > 0 Then
            Dim stgArr() As Variant
            stgArr = l_lo.listRows(tmpIdx).Range.Value
            resp = stgArr(1, SettingCols.scSettingValue)
            respType = stgArr(1, SettingCols.scSettingType)
            resp = FormatValue(resp, respType)
        End If
        Setting = resp
    End Property
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Setting Value Property Setter
''  If [stgKey] does not exist, it will be created, and the SettingType
''      will be '0' (TypeEnum.teGeneral = 0)
''  Typically you should not have to modify the SettingType unless for
''      some reason you'd be create the setting with a 'Value2' value, and
''      would need to return it with a different format
''      e.g. if you created a setting with the value: 45263.8997685185,
''          but would need to return it as a date: 12/3/23 9:35:40 PM
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let Setting(ByVal stgKey, ByVal stgVal)
        If Not ValidConfig Then Exit Property
        On Error Resume Next
        Dim tmpIdx As Long: tmpIdx = Index(stgKey)
        Dim resp, respType
        ''  Disable App Events
        Dim evts As Boolean: evts = Application.EnableEvents: Application.EnableEvents = False
        If tmpIdx > 0 Then
            'Update Setting
            If Not StringsMatch(TypeName(stgVal), TypeName(l_lo.listRows(tmpIdx).Range(1, SettingCols.scSettingValue).Value)) Then
                l_lo.listRows(tmpIdx).Range(1, SettingCols.scSettingValue).numberFormat = "General"
            End If
            
            l_lo.listRows(tmpIdx).Range(1, SettingCols.scSettingValue).Value = stgVal
            l_lo.listRows(tmpIdx).Range(1, SettingCols.scUpdated).Value = Now()
        Else
            Dim stgArr() As Variant
            ReDim stgArr(1 To 1, 1 To l_lo.ListColumns.Count)
            stgArr(1, SettingCols.scSettingKey) = stgKey
            stgArr(1, SettingCols.scSettingValue) = stgVal
            stgArr(1, SettingCols.scUpdated) = Now()
            stgArr(1, SettingCols.scSettingType) = TypeEnum.teGeneral
            l_lo.listRows.Add().Range.Value = stgArr
            CheckSort
        End If
        l_isDirty = True
        If Err.number <> 0 Then
            Err.Clear
        End If
        Application.EnableEvents = evts
    End Property
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Setting Property Getter for OS-specific setting
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get SettingForOS(ByVal stgKey)
        If Not ValidConfig Then Exit Property
        SettingForOS = Setting(BuildOSSpecificKey(stgKey))
    End Property
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Setting Property Setter for OS-specific setting
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let SettingForOS(ByVal stgKey, ByVal stgVal)
        If Not ValidConfig Then Exit Property
        Setting(BuildOSSpecificKey(stgKey)) = stgVal
    End Property
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Setting Property Getter for USER-specific setting
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get SettingForUser(ByVal stgKey)
        If Not ValidConfig Then Exit Property
        SettingForUser = Setting(BuildUSERSpecificKey(stgKey))
    End Property
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Setting Property Setter for USER-specific setting
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let SettingForUser(ByVal stgKey, ByVal stgVal)
        If Not ValidConfig Then Exit Property
        Setting(BuildUSERSpecificKey(stgKey)) = stgVal
    End Property
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Exposed for convenience -- return name of setting key
''      e.g. CheckKey("TEST",true,false) = "TEST_OS_MAC" OR "TEST_OS_PC"
''      e.g. CheckKey("TEST",false,true) =
''          "TEST_USER_[LoginName]"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function CheckKey(ByVal stgKey, isOSSpecific As Boolean, isUSERSpecific As Boolean)
        If Not ValidConfig Then Exit Function
        If isOSSpecific And isUSERSpecific Then Err.Raise 1004, Description:="A pbSetting Key cannot be both OS and User specific"
        If isOSSpecific Then
            CheckKey = BuildOSSpecificKey(stgKey)
        ElseIf isUSERSpecific Then
            CheckKey = BuildUSERSpecificKey(stgKey)
        Else
            CheckKey = stgKey
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns timestamp when Setting [stgKey] was last modified
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function SettingUpdated(ByVal stgKey, Optional isOSSpecific As Boolean = False, Optional isUSERSpecific As Boolean = False)
        If Not ValidConfig Then Exit Function
        stgKey = CheckKey(stgKey, isOSSpecific, isUSERSpecific)
        If Exists(stgKey) Then
            SettingUpdated = CDate(l_lo.listRows(Index(stgKey)).Range(1, SettingCols.scUpdated))
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Check [stgKey] and append OS-specific string to key, if missing
''      e.g. on a PC,BuildOSSpecificKey("TEST") = "TEST_OS_PC"
''      e.g. on a PC,BuildOSSpecificKey("TEST_OS_PC") = "TEST_OS_PC"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function BuildOSSpecificKey(ByVal stgKey)
        Dim osKey
        If Not StringsMatchPriv(stgKey, "_OS_" & PBSTG_OS, smEndWithStr) Then
            osKey = stgKey & "_OS_" & PBSTG_OS
        Else
            osKey = stgKey
        End If
        BuildOSSpecificKey = osKey
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Check [stgKey] and append USER-specific string to key, if missing
''      e.g. if user = 'browerp',BuildUSERSpecificKey("TEST") =
''          'TEST_USER_browerp"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function BuildUSERSpecificKey(ByVal stgKey)
        Dim userKey
        If Not StringsMatchPriv(stgKey, "_USER_" & StgUserName, smEndWithStr) Then
            userKey = stgKey & "_USER_" & StgUserName
        Else
            userKey = stgKey
        End If
        BuildUSERSpecificKey = userKey
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Get current user name, based on OS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function StgUserName() As String
        #If Mac Then
            StgUserName = VBA.Interaction.Environ("LOGNAME")
        #Else
            StgUserName = VBA.Interaction.Environ("USERNAME")
        #End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Force Setting Value to Go Through 'Val()' Function
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ForceNumericFormat(ByVal stgKey, Optional isOSSpecific As Boolean = False, Optional isUSERSpecific As Boolean = False)
        If Not ValidConfig Then Exit Function
        stgKey = CheckKey(stgKey, isOSSpecific, isUSERSpecific)
        If Exists(stgKey) Then
            l_isDirty = True
            l_lo.listRows(Index(stgKey)).Range(1, SettingCols.scSettingType) = TypeEnum.teNumeric
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Force Setting Value to Go Through 'CDate()' Function
''  Note: A setting value that cannot be converted to a date will cause
''      an error
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ForceDateFormat(ByVal stgKey, Optional isOSSpecific As Boolean = False, Optional isUSERSpecific As Boolean = False)
        If Not ValidConfig Then Exit Function
        stgKey = CheckKey(stgKey, isOSSpecific, isUSERSpecific)
        If Exists(stgKey) Then
            l_isDirty = True
            l_lo.listRows(Index(stgKey)).Range(1, SettingCols.scSettingType) = TypeEnum.teDateTime
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Force Setting Value to Go Through 'CBool()' Function
''  Note: A setting value that cannot be converted to a Boolean will cause
''      an error
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ForceBooleanFormat(ByVal stgKey, Optional isOSSpecific As Boolean = False, Optional isUSERSpecific As Boolean = False)
        If Not ValidConfig Then Exit Function
        stgKey = CheckKey(stgKey, isOSSpecific, isUSERSpecific)
        If Exists(stgKey) Then
            l_isDirty = True
            l_lo.listRows(Index(stgKey)).Range(1, SettingCols.scSettingType) = TypeEnum.teBoolean
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Force Setting Value to be returned as it was created (This is the Default)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ForceGeneralFormat(ByVal stgKey, Optional isOSSpecific As Boolean = False, Optional isUSERSpecific As Boolean = False)
        If Not ValidConfig Then Exit Function
        stgKey = CheckKey(stgKey, isOSSpecific, isUSERSpecific)
        If Exists(stgKey) Then
            l_isDirty = True
            l_lo.listRows(Index(stgKey)).Range(1, SettingCols.scSettingType) = TypeEnum.teGeneral
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Delete Setting with key [stgKey], if it exists
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function Delete(ByVal stgKey, Optional isOSSpecific As Boolean = False, Optional isUSERSpecific As Boolean = False)
        If Not ValidConfig Then Exit Function
        stgKey = CheckKey(stgKey, isOSSpecific, isUSERSpecific)
        If Index(stgKey) > 0 Then
            Dim cancelOperation As Boolean
            RaiseEvent OnDeleteSetting(stgKey, cancelOperation)
            Beep
            Stop
            If cancelOperation = True Then
                'cancel
            Else
                Dim evts As Boolean: evts = Application.EnableEvents: Application.EnableEvents = False
                ''  Delete Setting
                l_lo.listRows(Index(stgKey)).Delete
                ''  Restore App Events to prior setting
                Application.EnableEvents = evts
                l_isDirty = True
            End If
        End If
    End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns Index of List for [settingKey]
''  Returns 0 if Missing
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function Index(ByVal stgKey, Optional isOSSpecific As Boolean = False, Optional isUSERSpecific As Boolean = False) As Long
        If Not ValidConfig Then Exit Function
        stgKey = CheckKey(stgKey, isOSSpecific, isUSERSpecific)
        If ValidConfig And Count > 0 Then
            Dim tmpIdx As Variant
            On Error Resume Next
            tmpIdx = WorksheetFunction.XMatch(stgKey, l_lo.ListColumns(SettingCols.scSettingKey).DataBodyRange, XMatch.exactMatch, XSearch.xsBinaryAsc)
            If Err.number = 0 Then
                Index = CLng(tmpIdx)
            Else
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Exposed for Convenience - not needed for using pbSettings
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''  SETTINGS WORKSHEET REFERENCE
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get pbSettingsSheet() As Worksheet
        If Not ValidConfig Then Exit Property
        Set pbSettingsSheet = l_ws
    End Property
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''  SETTINGS LIST OBJECT REFERENCE
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get pbSettingsListObj() As ListObject
        If Not ValidConfig Then Exit Property
        Set pbSettingsListObj = l_lo
    End Property
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''  Earliest Modified Date for any Setting
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get ModifiedEarliestDate() As Variant
        If Not ValidConfig Then Exit Property
        If ValidConfig Then
            Dim modified As Variant
            modified = CDate(WorksheetFunction.Min(l_lo.ListColumns(SettingCols.scUpdated).DataBodyRange))
            ModifiedEarliestDate = modified
        End If
    End Property
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''  Earliest Modified Date for any Setting
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get ModifiedLatestDate() As Variant
        If Not ValidConfig Then Exit Property
        If ValidConfig Then
            Dim modified As Variant
            modified = CDate(WorksheetFunction.Max(l_lo.ListColumns(SettingCols.scUpdated).DataBodyRange))
            ModifiedLatestDate = modified
        End If
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns the Count of Unique Settings
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get Count() As Long
        If Not ValidConfig Then Exit Property
        Count = l_lo.listRows.Count
    End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Exports Settings that match [wildcardsearch] if provided, otherwise all
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ExportSettings(Optional wildcardSearch As Variant)
        If Not ValidConfig Then Exit Function
        If Count > 0 Then
            Dim exportArr() As Variant, iRow, iCol, newWB As Workbook, newWS As Worksheet
            Dim tmpIndexes As New Collection
            exportArr = l_lo.Range.Value
            For iRow = LBound(exportArr) To UBound(exportArr)
                If iRow = 1 Then
                    'Add Column Headers Index
                    tmpIndexes.Add iRow
                Else
                    If IsMissing(wildcardSearch) Then
                        tmpIndexes.Add iRow
                    Else
                        For iCol = LBound(exportArr, 2) To UBound(exportArr, 2)
                            If StringsMatchPriv(exportArr(iRow, iCol), wildcardSearch, smContains) Then
                                tmpIndexes.Add iRow
                                Exit For
                            End If
                        Next iCol
                    End If
                End If
            Next
            Dim newArr() As Variant
            ReDim newArr(1 To tmpIndexes.Count, 1 To l_lo.ListColumns.Count)
            Dim goodIndex, populatingIndex As Long
            For Each goodIndex In tmpIndexes
                populatingIndex = populatingIndex + 1
                For iCol = LBound(exportArr, 2) To UBound(exportArr, 2)
                    newArr(populatingIndex, iCol) = exportArr(goodIndex, iCol)
                Next iCol
            Next
            Set newWB = Workbooks.Add
            Set newWS = newWB.Worksheets(1)
            With newWS
                .Name = "SettingsExport"
                .Range("A5").Resize(RowSize:=tmpIndexes.Count, ColumnSize:=l_lo.ListColumns.Count).Value = newArr
                .UsedRange.EntireColumn.AutoFit
                
                .Range("A1").Value = "Settings Exported From: " & ThisWorkbook.FullName
                .Range("A2").Value = "Exported on: " & CStr(Now())
                If IsMissing(wildcardSearch) Then
                    wildcardSearch = "[None]"
                Else
                    wildcardSearch = """" & wildcardSearch & """"
                End If
                .Range("A3").Value = "Search Expression: " & wildcardSearch
            End With
            newWB.Activate
            Set newWS = Nothing
            Set newWB = Nothing
            Beep
        End If
    
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns True If SETTINGS_SHEET_CODENAME worksheet exists
''  and the worksheet contains a valid listobject for managing settings
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get ValidConfig() As Boolean
        ValidConfig = l_validCfg
    End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  If found, returns the worksheet identified by the constant:
''      SETTINGS_SHEET_CODENAME
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function VerifySettingsSheet() As Boolean
        Dim tmpWS As Worksheet
        Dim tmpLO As ListObject
        Dim lstCol As ListColumn
        '' First Try to Find by 'CodeName'
        Dim findCodeName: findCodeName = SETTINGS_SHEET_CODENAME
        If Exists(PBSTG_WS_CODENAME) Then
            findCodeName = Setting(PBSTG_WS_CODENAME)
        End If
        For Each tmpWS In ThisWorkbook.Worksheets
            If StringsMatchPriv(tmpWS.CodeName, findCodeName) Then
                Set l_ws = tmpWS
                Exit For
            End If
        Next
        If l_ws Is Nothing Then
            For Each tmpWS In ThisWorkbook.Worksheets
                If StringsMatchPriv(tmpWS.Name, SETTINGS_SHEET_WEAKNAME) Then
                    Set l_ws = tmpWS
                    Exit For
                End If
            Next
        End If

        If Not l_ws Is Nothing Then
            Set l_lo = GetSettingsLO(l_ws)
        End If
        VerifySettingsSheet = Not l_ws Is Nothing And Not l_lo Is Nothing
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  If found, returns the validated pbSettings list object from [wksht]
''  The list object name must match the constant SETTINGS_LO exactly
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Note: This method is used to get the reference to the settings list object
''      in the current workbook, and it is also used to validate the pbSettings
''      ListObject structure in the template that is used for FIRST RUN
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function GetSettingsLO(wksht As Worksheet) As ListObject
            Dim tmpLO As ListObject
            Dim validLO As ListObject, lstCol As ListColumn
            For Each tmpLO In wksht.ListObjects
                If StringsMatchPriv(tmpLO.Name, SETTINGS_LO) Then
                    Dim expectCols As Variant
                    expectCols = ExpectedColsArr
                    If ArrSize(expectCols) = tmpLO.ListColumns.Count Then
                        For Each lstCol In tmpLO.ListColumns
                            If StringsMatchPriv(lstCol.Name, expectCols(lstCol.Index)) = False Then
                                Exit For
                            End If
                            If lstCol.Index = tmpLO.ListColumns.Count Then
                                Set validLO = tmpLO
                                Exit For
                            End If
                        Next
                    End If
                End If
                If Not validLO Is Nothing Then
                    Exit For
                End If
            Next
            Set GetSettingsLO = validLO
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns the explicit name of ListColumns as a Base-1 1D Array,
''       that are expected to be in the settings ListObject
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Property Get ExpectedColsArr() As Variant
        Dim tmpArr As Variant: tmpArr = Split(SETTINGS_COL_NAMES, delimiter:="|", compare:=vbTextCompare)
        Dim i As Long, colName As Variant, tmpArrSize As Long
        tmpArrSize = ArrSize(tmpArr)
        Dim respArr() As Variant
        ReDim respArr(1 To ArrSize(tmpArr))
        For Each colName In tmpArr
            i = i + 1
            If i <= tmpArrSize Then
                respArr(i) = colName
            End If
        Next
        ExpectedColsArr = respArr
    End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Called when class is initialized, and any time a new settings Key
'   has been added
'   Will check the following:
'       1 - Ensure there are no duplicate settings keys
'             (Will remove earliest modified row if duplicate keys are found)
'       2 - Ensure ListObject is sorted by the Key Column
'             (Use: SettingCols.scSettingKey to refer to Key Column Index)
'       3 - Remove Any Invalid 'SettingType'
'             (Valid values are: String, Boolean, Number, DateTime)
'             (This helps ensure retrieved settings formats will populate
'              into Variant or explicitly declared types correctly --
'              e.g. if you added '45262.9420949074' as the setting value
'              for a setting declared as 'DateTime', the setting would return
'              '12/2/23 10:36:37 PM' as a Variant of type Date)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function CheckSettingsHealth()
        Dim evts As Boolean: evts = Application.EnableEvents: Application.EnableEvents = False
        If l_lo.ShowTotals Then l_lo.ShowTotals = False
        CheckSort
        CheckRequiredSettings
        If AutoHide And l_ws.visible = xlSheetVisible And OtherSheetsVisible Then
            l_ws.visible = xlSheetVeryHidden
        End If
        CheckDuplicates
        l_lo.Range.EntireColumn.AutoFit
        Application.EnableEvents = evts
    End Function



' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Used to verify there is another worksheet that's visible, so that
''   we can autohide the pbSettings sheet if it is visible and the
''  setting "PBSTG_AUTOHIDE" is set to TRUE
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function CheckDuplicates()
        On Error Resume Next
        Dim uniqueKeys As Variant
        Dim loName, keyColName, expr
        loName = l_lo.Name
        keyColName = l_lo.ListColumns(SettingCols.scSettingKey).Name
        expr = "COUNTA(UNIQUE(" & loName & "[" & keyColName & "]))"
        uniqueKeys = Evaluate(expr)
        If val(uniqueKeys) < Count Then
            l_isDirty = True
            RemoveDuplicates
        End If
        If Err.number <> 0 Then
            Err.Clear
            '' If there was an error checking, then look through each row
            RemoveDuplicates
        End If
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Removes duplicates
''  Will keep the setting with the latest updated date if duplicates are found
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function RemoveDuplicates()
        Dim deleteItems As New Collection
        Dim srcArr() As Variant
        srcArr = l_lo.DataBodyRange.Value
        Dim idx As Long, startIdx As Long
        Dim keepIdx, keepKey, keepDt
        For idx = LBound(srcArr) To UBound(srcArr)
            If CollectionItemExists(deleteItems, idx) = False Then
                keepIdx = idx
                keepKey = srcArr(idx, SettingCols.scSettingKey)
                keepDt = CDate(srcArr(idx, SettingCols.scUpdated))
                For startIdx = idx + 1 To UBound(srcArr)
                    If CollectionItemExists(deleteItems, startIdx) = False Then
                        If StringsMatchPriv(keepKey, srcArr(startIdx, SettingCols.scSettingKey)) Then
                            If CDate(srcArr(startIdx, SettingCols.scUpdated)) >= keepDt Then
                                deleteItems.Add keepIdx
                                keepIdx = startIdx
                                keepDt = CDate(srcArr(startIdx, SettingCols.scUpdated))
                            Else
                                deleteItems.Add startIdx
                            End If
                        End If
                    End If
                Next startIdx
            End If
        Next
        Dim deleteItem
        ''  Disable App Events
        Dim evts As Boolean: evts = Application.EnableEvents: Application.EnableEvents = False
        For Each deleteItem In deleteItems
            l_lo.listRows(deleteItem).Range(1, SettingCols.scSettingKey) = "PB_SETTING_DUPLICATE_DELETE"
        Next
        Sort
        Dim delRng As Range
        Set delRng = l_lo.ListColumns(SettingCols.scSettingKey).Range.Find("PB_SETTING_DUPLICATE_DELETE")
        If Not delRng Is Nothing Then
            Do While Not delRng Is Nothing
                delRng.Resize(ColumnSize:=l_lo.ListColumns.Count).Delete xlShiftUp
                Set delRng = l_lo.ListColumns(SettingCols.scSettingKey).Range.Find("PB_SETTING_DUPLICATE_DELETE")
            Loop
        End If
        ''  Restore App Events to prior setting
        Application.EnableEvents = evts
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Used to verify there is another worksheet that's visible, so that
''   we can autohide the pbSettings sheet if it is visible and the
''  setting "PBSTG_AUTOHIDE" is set to TRUE
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function OtherSheetsVisible() As Boolean
        Dim vWS As Worksheet
        For Each vWS In ThisWorkbook.Worksheets
            If vWS.visible = xlSheetVisible And Not vWS Is l_ws Then
                OtherSheetsVisible = True
                Exit For
            End If
        Next
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Check Required Settings is for settings that need to exist that
''   are used to manage pbSettings.
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function CheckRequiredSettings()
        '' If 'PBSTG_AUTOHIDE' key does not exist, create it with default: True
        If Exists(PBSTG_HIDE) = False Then
            AutoHide = True
        End If
        If Exists(PBSTG_WS_CODENAME) = False Then
            Setting(PBSTG_WS_CODENAME) = l_ws.CodeName
        End If
        
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Check is settings ListObject is ordered -- needed to support the
'       binary searching of keys
'   This method is called when class is initialzed and whever a settings
'       KEY is added
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function CheckSort()
        If Not Sorted Then Sort
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Ensure Settings are sorted by Keys, so that binary search can be
'   used to most quickly find the setting
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function Sort()
        With l_lo.Sort
            ''  Disable App Events
            Dim evts As Boolean: evts = Application.EnableEvents: Application.EnableEvents = False
            .SortFields.Clear
            .SortFields.Add l_lo.ListColumns(SettingCols.scSettingKey).DataBodyRange, SortOn:=xlSortOnValues, Order:=XlSortOrder.xlAscending
            .Apply
            ''  Restore App Events to prior setting
            Application.EnableEvents = evts
        End With
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Build and evaluate Array Formulat to determine if sorted by Key
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function Sorted() As Boolean
        If Not ValidConfig Then Exit Function
        With l_lo
            If .listRows.Count <= 1 Then
                Sorted = True
            Else
                Dim expr As String, rng1 As Range, rng2 As Range
                Set rng1 = .ListColumns(SettingCols.scSettingKey).DataBodyRange.Resize(.listRows.Count - 1)
                Set rng2 = rng1.offset(rowOffset:=1)
                expr = "AND(" & "'[" & ThisWorkbook.Name & "]" & rng1.Worksheet.Name & "'!" & rng1.Address & "<='[" & ThisWorkbook.Name & "]" & rng2.Worksheet.Name & "'!" & rng2.Address & ")"
                Sorted = Evaluate(expr)
            End If
        End With
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Returns the count of  elements from the first dimension of an array
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function ArrSize(inputArr As Variant) As Long
        ArrSize = UBound(inputArr) - LBound(inputArr) + 1
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Private (class) version of the StringsMatch Function available in
'   some of the common modules at:
'     https://github.com/lopperman/just-VBA
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
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
    
    Private Function FormatValue(ByVal item, ByVal formatType)
        If Len(formatType & "") = 0 Then
            FormatValue = item
        ElseIf StringsMatchPriv(formatType, TypeEnum.teGeneral) Then
            FormatValue = item
        ElseIf StringsMatchPriv(formatType, TypeEnum.teBoolean) Then
            FormatValue = CBool(item)
        ElseIf StringsMatchPriv(formatType, TypeEnum.teNumeric) Then
            FormatValue = val(item)
        ElseIf StringsMatchPriv(formatType, TypeEnum.teDateTime) Then
            FormatValue = CDate(item)
        Else
            FormatValue = item
        End If
    End Function
    

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   'FIRST RUN' METHODS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Once the correctly structured pbSettings Worksheet and ListObject
'   Have been added, these methods are not longer used
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   ONLY USED IN 'FIRST RUN'
'   Create settings sheet
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function CreateSettingsSheet() As Boolean
        On Error Resume Next
        Dim tmplSheet As Worksheet
        Dim evts As Boolean, intr As Boolean, curs As Variant
        evts = Application.EnableEvents
        intr = Application.Interactive
        curs = Application.Cursor
        Application.StatusBar = "Configuring new Settings Area ...": DoEvents
        Application.EnableEvents = False
        Application.Interactive = False
        Application.Cursor = XlMousePointer.xlWait
        
        Set tmplSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        tmplSheet.Name = SETTINGS_SHEET_WEAKNAME
        tmplSheet.Activate
        ThisWorkbook.Windows(1).DisplayGridlines = False
                
        Dim tRng As Range
        Set tRng = tmplSheet.Range("A6").Resize(ColumnSize:=ArrSize(ExpectedColsArr))
        tRng.Value = ExpectedColsArr
        Set tRng = tRng.Resize(RowSize:=2)
        Dim lstObj As ListObject
        Set lstObj = tmplSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tRng, XlListObjectHasHeaders:=xlYes, TableStyleName:="TableStyleMedium2")
        lstObj.Name = SETTINGS_LO
        With lstObj.listRows(1)
            .Range(1, SettingCols.scSettingKey) = PBSTG_HIDE
            .Range(1, SettingCols.scSettingValue) = True
            .Range(1, SettingCols.scSettingType) = TypeEnum.teBoolean
            .Range(1, SettingCols.scUpdated) = Now()
        End With
        lstObj.HeaderRowRange.HorizontalAlignment = xlHAlignCenter
        lstObj.HeaderRowRange.VerticalAlignment = xlVAlignCenter
        lstObj.HeaderRowRange.Font.Bold = True
        lstObj.ListColumns(SettingCols.scSettingKey).DataBodyRange.Font.Bold = True
        lstObj.ListColumns(SettingCols.scSettingKey).DataBodyRange.Font.color = 16724484
        Set tmplSheet = Nothing
        If Err.number = 0 Then
            Set l_ws = lstObj.Range.Worksheet
            Set l_lo = lstObj
            CreateSettingsSheet = VerifySettingsSheet
            AutoHide = True
        End If
        Application.EnableEvents = evts
        Application.Interactive = intr
        Application.Cursor = curs
        Application.StatusBar = False
        If Err.number <> 0 Then Err.Clear
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   END OF 'FIRST RUN' METHODS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   CLASS INITIALIZE
'   Verify Settings Worksheet
'   Start FIRST RUN process (CreateSettingsSheet) If Necessary
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Sub Class_Initialize()

    l_validCfg = VerifySettingsSheet
    If Not ValidConfig Then
        l_validCfg = CreateSettingsSheet
    Else
        CheckSettingsHealth
    End If
    Set stgwb = ThisWorkbook

End Sub

Private Sub Class_Terminate()
    Set stgwb = Nothing
    Set l_lo = Nothing
    Set l_ws = Nothing
End Sub


Private Sub stgwb_AfterSave(ByVal Success As Boolean)
    l_isDirty = False
End Sub

Private Sub stgwb_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    Me.Setting("WKBK_SAVED") = Now()
End Sub
