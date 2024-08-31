VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pbSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Settings Management
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      author (c) Paul Brower https://github.com/lopperman/just-VBA
''      class module pbSettings.cls
''      license GNU General Public License v3.0
''      Created Dec-2023
''      Updated 03-FEB-2024
''
''      Download pbSettings.cls with NO DEPENDENCIES at
''      https://github.com/lopperman/just-VBA/blob/main/Code_NoDependencies/pbSettings.cls
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Online Documentation
''      https://github.com/lopperman/just-VBA/blob/main/pbSettings/pbSettings.md
''      Download pbSettings Demo App
''      https://github.com/lopperman/just-VBA/raw/main/pbSettings/pbSettings_Example.xlsm
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  * Supports configuration of settings for:
''
''      * General Settings which can be used for any user / any OS
''
''      * Automatically make a setting for any user by prefixing
''          SettingKey with 'USER_' (e.g. USER_ZOOM becomes USER_[LogonName]_ZOOM)
''          This enables customization by users which will not interfere with other users
''
''      * Automatically make a setting specific to PC or MAC by including '_OS' as
''          the SettingKey suffix (e.g. SHEET1_FONTSIZE_OS becomes
''          'SHEET1_FONTSIZE_OS_MAC' or 'SHEET1_FONTSIZE_OS_PC'
''          This enables customizations per platform, which can improve general
''          'Cleanliness' for font/fontsize defaults
''
''      * 'USER_' prefix and '_OS' suffix can be combined if needed.  e.g.
''          'USER_DEFAULT_FONT_OS' becomes
''          'USER_[LogonName]_DEFAULT_FONT_OS)['MAC' OR 'PC']
''
''      * 'USER_' and/or '_OS' can be 'forced' into a SettingKey by including one or
''          both of these enums anywhere a 'pbSettingOption' parameter is available
''          e.g. SetValue(stgKey, stgVal, optional options as pbSettingOption) could be used like:
''          SetValue "GENERIC_SETTING", Now() _
''              , pbSettingOption.optOSSpecific + pbSettingOption.optUserSpecific _
''              , pbSettingOption.optValueAsDate
''          This would save a value of the current date (without time -- because of 'optValueAsDate'
''          and the stgKey would be 'USER_[LogonName]_GENERIC_SETTING_['PC' OR 'MAC']
''          NOTE:  ONCE A DATA TYPE (e.g. optValueAsDate) has been set -- usually when
''              A Setting ** KEY ** is CREATED, subsequent call to set or get the setting value
''              will automatically convert the value to the defined format (Date, DateTime, Boolean, etc)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  NOTE REGARDING 'StandardSettingEnum' Enum, and Setting Key Value Type
''  Any public method that has a [stgKey] argument will always attempt to convert
''      numeric values to a StandardSettingEnum enum
''      e.g.  StandardSettingEnum.stgArraySettingDelimiter  = 2, which is represented
''          by the key: "PBSTG_ARRAY_DELIM"
''  If a numeric arguement cannot be converted to StandardSettingEnum enum value
''      then the arguement will be converted to a string
''      e.g.  stgKey=200 will be converted to "200"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Option Explicit
Option Compare Text
Option Base 1
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  THESE CONSTANTS DEFINE SETTINGS KEYS USED TO MANAGE THE
''  BEHAVIOR OF THE pbSettings Worksheet and ListObject, and some
''  COMMON UTILITY ("PBSTG_") AND DEVELOPER UTILITY SETTINGS ('PBSTG_DEV_***)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
                
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  SETTING KEY CONSTANTS FOR CONTROLLING pbSettings FUNCTIONS, INCLUDING
''  UI FOR WORKSHEET WHERE SETTINGS ARE STORED
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Const PBSTG_HIDE As String = "PBSTG_AUTOHIDE"
    Private Const PBSTG_WS_CODENAME As String = "PBSTG_WS_CODE_NAME"
    Private Const PBSTG_ARRAY_DELIMITER As String = "PBSTG_ARRAY_DELIM"
    Private Const PBSTG_FONT_NAME As String = "PBSTG_SETTINGS_FONT_NAME_OS"
    Private Const PBSTG_FONT_SIZE As String = "PBSTG_SETTINGS_FONT_SIZE_OS"
    Private Const PBSTG_MAX_SETTINGS_COL_WIDTH As String = "PBSTG_SETTINGS_COLWIDTH_MAX"
    Private Const PBSTG_RETURN_1BASED_ARRAYS As String = "PBSTG_1BASED_ARRAYS"
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  'DEVELOPER HELPER' SETTING KEY CONSTANTS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Const PBSTG_DEV_LOG_TO_IMMEDIATE As String = "DEV_OUTPUT_TRACE"
    Private Const PBSTG_DEV_MODE As String = "DEV_MODE"
    Private Const PBSTG_DEV_SHOW_TOOLS As String = "SHOW_DEV_TOOLS"
    Private Const PBSTG_DEV_NAMES = "DEV_NAMES"
    Private Const PBSTG_DEV_NO_NETWORK = "DEV_NO_NETWORK"
    Private Const PBSTG_ADMIN_NAMES = "ADMIN_NAMES"
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  SETTING KEY FOR 'VERSION'
''  'VERSION' CAN BE USED TO STORE A VERSION NUMBER FOR YOUR WORKBOOK
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Const PBSTG_VERSION As String = "VERSION"
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  SETTING KEYS WHICH pbSettings USES TO MODIFY GENERAL WORKBOOK BEHAVIOR
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Const PBSTG_DISABLE_AUTO_SAVE As String = "PBSTG_AUTOSAVE_DISABLE"
    Private Const PBSTG_ENABLE_CONNECTIONS As String = "PBSTG_ENABLE_CONNECTION"
    Private Const PBSTG_SAVE_LINK_VALUES As String = "PBSTG_SAVE_LINK_VALS"
    Private Const PBSTG_ZOOM_MAC_DEF As String = "PBSTG_MAC_ZOOM_DEF"
    Private Const PBSTG_ZOOM_PC_DEF As String = "PBSTG_PC_ZOOM_DEF"
    Private Const PBSTG_ZOOM_ENABLE As String = "PBSTG_ENABLE_AUTO_ZOOM"
    Private Const PBSTG_USER_ZOOM As String = "USER_ZOOM_OS"
    Private Const PBSTG_ARRAY_ZOOM_LIMITS As String = "PBSTG_ARRAY_ZOOM_LIMITS"
    Private Const PBSTG_DO_NOT_HIDE As String = "DO_NOT_HIDE_SHEETCODENAMES"
    Private Const PBSTG_DO_NOT_PROTECT As String = "DO_NOT_PROTECT"
    
    Private Const BUTTON_ACTION_PREFIX As String = "BTN_ACT"


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  PUBLIC ENUM (Get's Registered in TypeLibrary and can be referenced outside '
''      of pbSettings
''  THIS IS A FLAG ENUM: ENUM MEMBERS CAN BE ADDED TOGETHER
''  NOTE: CERTAIN COMBINATIONS ARE NOT ALLOWED OR WILL BE IGNORED, e.g.
''      If more than one 'optValueAs...' is included, an error will be raised
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Enum pbSettingOption
        optDefault = 0
        optUserSpecific = &H1
        optOSSpecific = &H2
        optUserAndOS = &H4
        '' -- --
        [_optValueMin] = &H8
        optValueAsGeneral = &H8
        optValueAsText = &H10
        optValueAsLong = &H20
        optValueAsDouble = &H40
        optValueAsDate = &H80
        optValueAsDateTime = &H100
        optValueAsBoolean = &H200
        [_optValueMax] = &H200
        '' -- --
        optValueAsArray = &H400
        optMigratableKey = &H800
        optNonMigratableKey = &H1000
        optCreateKeyIfMissing = &H2000
        '' optSaveDefaultIfMissingValue = &H4000
        optRaiseValidationErrors = &H8000&
        optKeyRequired = &H10000
        '' -- --
        '' include optLogChange
        optLogChanges = &H20000
        optDoNotLogChanges = &H40000
        optAllowDuplicateArrayValues = &H80000
        
    End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  StandardSettingEnum Enum
''  Can be used in place of 'stgKey'
''      e.g. If stg.GetValue(StandardSettingEnum.stgAutoHideSettingsSheet) = False Then ...
''  Can also use 'GetStandardSettingValue' where StandardSettingEnum is an explicit
''      arguement, so 'intellisense' will present the enum choices when writing code
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Enum StandardSettingEnum
        ''  SHOULD ALWAYS BE '1'
        [_StandardSettingKeyFirst] = 1
        stgAutoHideSettingsSheet = 1
        stgArraySettingDelimiter
        stgSettingsSheetMaxColWidth
        stgDisableAutoSave
        stgEnableConnections
        stgSaveLinkedValues
'        stgArrayDoNotHideSheets
        stgArrayDoNotProtectSheets
        stgDEVShowLogTrace
        stgDEVIsDevMode
        stgShowDevTools
        stgDEVNames
        stgDEVRunAsNoNetwork
        ''  IF A LATER ENUM ITEM IS ADDED, DON'T FORGET TO CHANGE
        ''  [_StandardSettingKeyLast] TO EQUAL THE LAST ENUM IN LIST
        [_StandardSettingKeyLast] = stgDEVRunAsNoNetwork
    End Enum

#If Mac Then
    Private Const PBSTG_OS As String = "MAC"
#Else
    Private Const PBSTG_OS As String = "PC"
#End If
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  PRIVATE TYPE TO PASS AROUND INSIDE OF CLASS TO REMOVE REDUNDANT
''  LOOKUPS FOR KEY SETTINGS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Type pbKey
        UncheckedKey As Variant
        Key As Variant
        OldVal As Variant
        OldValFormatted As Variant
        NewVal As Variant
        NewValFormatted As Variant
        defaultValue As Variant
        '' ROW INDEX FOR KEY CURRENTLY BEING UPDATED
        IDX As Long
        Updated As Variant
        '' value (Long) currently written to settings table
        OldOptions As pbSettingOption
        NewOptions As pbSettingOption
        '' options that were passed in to a function call, or created within pbSettings
        ''  for temporary use
        procOptions As pbSettingOption
    End Type
    

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  allow workbook after save event to be captured to mark pbSettings 'not dirty'
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private WithEvents stgwb As Workbook
Attribute stgwb.VB_VarHelpID = -1

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private Constants
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
        
    Private Const SETTINGS_SHEET_WEAKNAME As String _
        = "pbSettings"
        
    Private Const SETTINGS_SHEET_CODENAME As String _
        = "wsSTGWS"
        
    Private Const SETTINGS_LO As String _
        = "tbl_pbSettings"
        
    Private Const SETTINGS_COL_NAMES As String _
        = "SettingKey|SettingValue|SettingOptions|Updated"
        
    Private Const SETTINGS_MAX_COLUMN_WIDTH As Long _
        = 70
        
    Private Const COLUMN_NAME_KEY As String = "SettingKey"
    Private Const COLUMN_NAME_VALUE As String = "SettingValue"
    Private Const COLUMN_NAME_OPTIONS As String = "SettingOptions"
    Private Const COLUMN_NAME_UPDATED As String = "Updated"
    

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private variable used to store whether or not the pbSettings is
''      Valid and usable
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private l_validCfg As Boolean
    Private l_InitAttempted As Boolean
    Private l_invalidMessageShown As Boolean

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Private variable used to store whether or not pbSettings has unsaved
''  changes
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private l_isDirty As Boolean

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  'l_ws' stores the reference to the pbSettings worksheet
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private l_ws As Worksheet

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  'l_lo' stores the reference to the pbSettings ListObject
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private l_lo As ListObject


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

    Private Enum DateDiffType_priv
        dtpSecond
        dtpMinute
        dtpHour
        dtpDay
        dtpWeek
        dtpMonth
        dtpYear
        dtpQuarter
        dtpDayOfYear
        dtpWeekday
        dtpDate_NoTime
        dtpTimeString
    End Enum


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Enum that can (and should) be  used when referring to a ListColumn
''  in the Settings List Object
''  (Public Enums declared in class modules create common types that are
''      accessible anywhere in project where class module is included)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Enum pbSTGColumns
        scSettingKey = 1
        scSettingValue = 2
        scSettingOptions = 3
        scUpdated = 4
        [_TOTALCOLS] = 4
    End Enum

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Value of 'SettingTypeEnum' is stored in pbSTGColumns.scSettingType (3) Column
''  Unless Specified When SETTING a setting value, will be 'teGeneral' (0)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Enum SettingTypeEnum
        teUNKNOWN = -1
        teGeneral = 0
        ''for compatibility (teNumeric)
        teNumeric = 1
        teDouble
        teBoolean
        teDatetime
        teLong
        teDate
        teText
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
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  pbSettings Events
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  If event is handled, changing 'Cancel' to True will prevent
''  delete operation
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Event OnDeleteSetting(ByVal settingKey, ByRef Cancel As Boolean)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  If event is handled, changing 'Cancel' to True will prevent
''  delete operation
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Event OnModifySetting(ByVal settingKey, ByVal settingValue, ByVal isNew As Boolean)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Warning or Error Occured -- handle this event to get data to show users or send to
''      Log
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Event OnAlert(ByVal isError As Boolean, ByVal msg)
    Public Event OnRequestWorksheetUnprotect(wks As Worksheet)

    Public Function TryUnprotect(wksht As Worksheet) As Boolean
        RaiseEvent OnRequestWorksheetUnprotect(wksht)
        TryUnprotect = Not (wksht.protectContents Or wksht.protectDrawingObjects)
    End Function

    Private Function NewKey(ByVal stgKey, options As pbSettingOption, Optional defaultVal) As pbKey
        Dim resp As pbKey
        If Not StringsMatchPriv(TypeName(stgKey), "String") Then
            On Error Resume Next
            Dim stdKey: stdKey = GetStandardSettingKey(CLng(stgKey))
            If stdKey = vbNullString Then
                stgKey = CStr(stgKey)
            Else
                stgKey = stdKey
            End If
        End If
        resp.UncheckedKey = stgKey
        resp.Key = CheckKey(stgKey, options)
        resp.IDX = keyIndex(resp.Key)
        If Not IsMissing(defaultVal) Then
            resp.defaultValue = defaultVal
        End If
        If resp.IDX > 0 Then
            Dim tARR() As Variant
            tARR = l_lo.listRows(resp.IDX).Range.value
            resp.OldOptions = tARR(1, pbSTGColumns.scSettingOptions)
            resp.OldVal = tARR(1, pbSTGColumns.scSettingValue)
            resp.OldValFormatted = FormatValue(resp.OldVal, resp.OldOptions)
            resp.Updated = tARR(1, pbSTGColumns.scUpdated)
        ElseIf privEnumCompare(options, pbSettingOption.optCreateKeyIfMissing) Then
            If Not IsMissing(defaultVal) Then
            
            End If
            resp.NewOptions = options
        End If
        If Not IsMissing(defaultVal) And resp.OldVal = vbNullString Then
            If Not isArray(defaultVal) Then
                resp.OldValFormatted = FormatValue(resp.defaultValue, resp.OldOptions)
            End If
        End If
        NewKey = resp
    End Function
''
''    Private Function SavePBKEY(pbk As pbKey)
''        On Error Resume Next
''        Dim tmpIdx As Long: tmpIdx = pbk.IDX
''        Dim validKey: validKey = pbk.Key
''        Dim resp, respType
''        Dim options As pbSettingOption
''        If pbk.IDX > 0 Then
''            options = pbk.OldOptions
''        Else
''            options = pbk.NewOptions
''        End If
''        ''  Disable App Events
''        Dim evts As Boolean: evts = Application.EnableEvents: Application.EnableEvents = False
''        If tmpIdx > 0 Then
''            'Update Setting
''            stgVal = pbk.NewValFormatted
''
''            If privEnumCompare(VarType(stgVal), VbVarType.vbArray) And Not privEnumCompare(options, pbSettingOption.optValueAsArray) Then
''                RaiseEvent OnAlert(False, "pbSettings.Setting (LET): [stgVal] was found to be of type 'Array', but the setting (" & stgKey & ") was originally set to another data type. SettingOption was forcibly changed to an array")
''                options = privEnumModify(options, pbSettingOption.optValueAsArray, FlagEnumModifyPriv.feVerifyEnumExists)
''                 l_lo.ListRows(tmpIdx).Range(1, pbSTGColumns.scSettingOptions).value = options
''                ArraySetting(validKey) = stgVal
''                Exit Function
''            End If
''            If privEnumCompare(options, pbSettingOption.optValueAsArray) Then
''                ArraySetting(validKey) = stgVal
''                Exit Function
''            End If
''            RaiseEvent OnModifySetting(validKey, stgVal, False)
''            If Not StringsMatchPriv(TypeName(stgVal), TypeName(l_lo.ListRows(tmpIdx).Range(1, pbSTGColumns.scSettingValue).value)) Then
''                l_lo.ListRows(tmpIdx).Range(1, pbSTGColumns.scSettingValue).NumberFormat = "General"
''            End If
''            l_lo.ListRows(tmpIdx).Range(1, pbSTGColumns.scSettingValue).value = ConvertItemFormat(stgVal, options)
''            l_lo.ListRows(tmpIdx).Range(1, pbSTGColumns.scUpdated).value = Now()
''        Else
''            If privEnumCompare(VarType(stgVal), VbVarType.vbArray) Then
''                options = optValueAsArray + optValueAsGeneral
''            Else
''                options = optValueAsGeneral
''            End If
''            RaiseEvent OnModifySetting(validKey, stgVal, True)
''            Dim stgArr() As Variant
''            Dim newRow As ListRow
''            ReDim stgArr(1 To 1, 1 To l_lo.ListColumns.Count)
''            stgArr(1, pbSTGColumns.scSettingKey) = "'" & validKey
''            stgArr(1, pbSTGColumns.scSettingValue) = ConvertItemFormat(stgVal, options)
''            stgArr(1, pbSTGColumns.scUpdated) = Now()
''            stgArr(1, pbSTGColumns.scSettingOptions) = options
''            Set newRow = l_lo.ListRows.Add
''            newRow.Range.value = stgArr
''            newRow.Range(1, pbSTGColumns.scSettingValue).NumberFormat = "General"
''            CheckSort
''        End If
''        l_isDirty = True
''        If Err.number <> 0 Then
''            Err.Clear
''        End If
''        Application.EnableEvents = evts
''    End Function

''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  MAIN EXPOSED SETTINGS METHODS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Return True If [checkUser] Is an entry in the array-value setting 'DEV_NAMES'
''  If [checkUser] is missing,will use computer username of current user
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function IsDeveloper(Optional checkUser, Optional doNotCheckPause As Boolean = False) As Boolean
        On Error Resume Next
        If IgnoreDev Then
            IsDeveloper = False
            Exit Function
        End If
        Dim curUserName: curUserName = UserNameOrLogin
        Dim devNames As Variant
        Dim findKey: findKey = PBSTG_DEV_NAMES
        Dim foundIdx As Long
        Dim resp As Boolean
        If ValidConfig Then
            devNames = Setting(findKey)
        Else
            Dim rng As Range, foundRng As Range
            Set rng = l_lo.ListColumns(pbSTGColumns.scSettingKey).Range
            Set foundRng = rng.Find(findKey, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
            If Not foundRng Is Nothing Then
                Set foundRng = foundRng.Offset(ColumnOffset:=(scSettingValue - scSettingKey))
                devNames = Split(CStr(foundRng.value), "|", compare:=vbTextCompare)
            End If
        End If
        If privEnumCompare(VarType(devNames), VbVarType.vbArray) Then
            Dim checkName
            For Each checkName In devNames
                If StringsMatchPriv(checkName, curUserName) Then
                    resp = True
                    Exit For
                End If
            Next checkName
        End If
        If resp = True Then
            If Exists("DEV_PAUSE_POPUPS_UNTIL") And doNotCheckPause = False Then
                Dim pUntil
                pUntil = Setting("DEV_PAUSE_POPUPS_UNTIL")
                If StringsMatchPriv(TypeName(pUntil), "Date") Then
                    If pUntil > Now() Then
                        resp = False
                    End If
                End If
            End If
        End If
        IsDeveloper = resp
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Return True If [checkUser] Is an entry in the array-value setting 'ADMIN_NAMES'
''  If [checkUser] is missing,will use computer username of current user
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function IsAdmin(Optional checkUser) As Boolean
        On Error Resume Next
        If IsMissing(checkUser) Then checkUser = UserNameOrLogin
        Dim adminNames As Variant
        Dim findKey: findKey = PBSTG_ADMIN_NAMES
        Dim foundIdx As Long
        If ValidConfig Then
            adminNames = Setting(findKey)
        Else
            Dim rng As Range, foundRng As Range
            Set rng = l_lo.ListColumns(pbSTGColumns.scSettingKey).Range
            Set foundRng = rng.Find(findKey, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
            If Not foundRng Is Nothing Then
                Set foundRng = foundRng.Offset(ColumnOffset:=(scSettingValue - scSettingKey))
                adminNames = Split(CStr(foundRng.value), "|", compare:=vbTextCompare)
            End If
        End If
        If privEnumCompare(VarType(adminNames), VbVarType.vbArray) Then
            Dim checkName
            For Each checkName In adminNames
                If StringsMatchPriv(checkName, checkUser) Then
                    IsAdmin = True
                    Exit Function
                End If
            Next checkName
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Record timestamp to indicate when a table / listobject was edited
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function ListObjectEditKey(ByVal lstObj) As Variant
        Dim tmpKey
        If StringsMatchPriv(TypeName(lstObj), "ListObject") Then
            tmpKey = lstObj.Name
        Else
            tmpKey = CStr(lstObj)
        End If
        ListObjectEditKey = "EDIT_" & tmpKey
    End Function
    
    Public Function ListObjectOnEdit(ByVal lstObj)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        Setting(ListObjectEditKey(lstObj)) = Now()
    End Function
    Public Property Get ListObjectEdited(ByVal lstObj) As Variant
        If Not ValidConfig Then Exit Property
        Dim tmpKey: tmpKey = ListObjectEditKey(lstObj)
        If Exists(tmpKey) = False Then
            Setting(tmpKey) = Now()
        End If
        ListObjectEdited = Setting(tmpKey)
    End Property
    
    
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns True if [settingKey] Exists
''  Note:  A setting will be created automatically, if needed, when
''      it is set using: Setting([stgKey]) = [stgVal]
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function Exists(ByVal settingKey, Optional options As pbSettingOption = pbSettingOption.optDefault) As Boolean
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        Exists = Index(CheckKey(settingKey, options)) > 0
    End Function
    
    Public Property Get IsDirty() As Boolean
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        IsDirty = l_isDirty
    End Property
        

    
    Private Function SizeColumns()
        On Error Resume Next
        l_lo.Range.EntireColumn.AutoFit
        Dim lcCol As listColumn
        For Each lcCol In l_lo.ListColumns
            If lcCol.Range.EntireColumn.ColumnWidth > SETTINGS_MAX_COLUMN_WIDTH Then
                lcCol.Range.EntireColumn.ColumnWidth = SETTINGS_MAX_COLUMN_WIDTH
                lcCol.Range.WrapText = True
            End If
        Next lcCol
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''      SETTINGS KEYS STRUCTURE
''
''      *   All Keys are created as strings.  If you create a setting key: 1234, it will be
''          created as "1234".

''      *  If a Key value starts with "USER_", CURRENT user name automatically get's
''          appended with the current user network/login name, if it doesn't already exist
''          e.g. if you login name is 'smithj', the key "USER_ZOOM" is automatically
''              converted to 'USER_smithj_ZOOM'
''
''      *   If a Key value ends with "_OS", then that is automatically replaced with
''              "_OS_MAC", or "_OS_PC", depending on the current user's computer.
''
''      *   "USER" and "OS" are supported in the same Key, for example if 'JonSmith'
''          sometimes used a Mac and also used a PC, he would eventually have
''          2 settings keys for "USER_MESSAGE_BEEP_OS", WHICH WOULD BE:
''              "USER_JOHNSMITH_MESSAGE_BEEP_OS_MAC", and
''              "USER_JOHNSMITH_MESSAGE_BEEP_OS_PC"
''
''      *   If a method contains and optional parameter type 'pbSettingOption',
''          you can use optOSSpecific or optUserSpecific or optUserAndOS to force the
''          key to be convert to USER_[key], or [key]_OS, or [USER_[key]_OS
''           - A key: "USER_SOMESTRING" WILL ALWAYS CONVERT TO "USER_[user network name]_SOMESTRING"
''           - A KEY: "SOMESTRING" will be converted to"USER_[user network name]_SOMESTRING"
''              if the setting key is created using a method that accepts a pbSettingOption,
''              and that value is provided as optUserSpecific
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''      CREATING, UPDATING, AND RETRIEVING pbSETTINGS VALUES
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''

    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Gets [stgKey] value, if missing setting, create new setting with [stgDefaultVal]
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function SettingWithDefault(ByVal stgKey, Optional ByVal stgDefaultVal)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        
        If Not Exists(stgKey) Then
            If Not IsMissing(stgDefaultVal) Then
                Setting(stgKey) = stgDefaultVal
            End If
        End If
        SettingWithDefault = Setting(stgKey)
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Get ButtonAction - pbSettings Standard Setting
''      @worksheetCodeName = 'CodeName' Property of Worksheet containing Shape Button
''          whose Button Action will be determined by the name of this public sub or
''          function that does not require any parameters
''      @btnName = name of button on worksheet with CodeName = [@worksheetCodeName]
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get stdButtonAction(wksht As Worksheet, ByVal btnName As String) As String
        Dim vKey: vKey = BUTTON_ACTION_PREFIX & "_" & wksht.CodeName & "_" & btnName
        Dim baArr As Variant
        baArr = GetValue(vKey)
        If privEnumCompare(VarType(baArr), VbVarType.vbArray) Then
            stdButtonAction = baArr(UBound(baArr))
        Else
            stdButtonAction = baArr
        End If
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Let ButtonAction - pbSettings StandardSetting
''      @worksheetCodeName = 'CodeName' Property of Worksheet containing Shape Button
''          whose Button Action will be determined by the name of this public sub or
''          function that does not require any parameters
''      @btnName = name of button on worksheet with CodeName = [@worksheetCodeName]
''      @methodName = name of public sub or function that does not require any parameters
''          valid expression = [standard module name].[sub or function name], or [sub or function name]
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let stdButtonAction(wksht As Worksheet, ByVal btnName As String, ByVal methodName As String)
        Dim vKey: vKey = ButtonActionKey(wksht.CodeName, btnName)
        Dim pbk As pbKey
        pbk = NewKey(vKey, optNonMigratableKey + optValueAsText + optValueAsArray)
        pbk.NewVal = Array(wksht.CodeName, btnName, methodName)
        ''
'        Setting(pbk.Key) = pbk.NewVal
        ArraySetting(pbk.Key) = pbk.NewVal
        ''
        If pbk.IDX = 0 Then
            pbk.NewOptions = optNonMigratableKey + optValueAsText + optValueAsArray
            SettingOptions(pbk.Key) = pbk.NewOptions
        End If
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Get AutoHide - pbSettings Standard Setting
''      Returns Boolean value that indicates if pbSettings worksheet is automatically
''      Hidden (default: True)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get stdAutoHide() As Boolean
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        stdAutoHide = Setting(PBSTG_HIDE)
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Let AutoHide - pbSettings Standard Setting
''      @hideSheetDefault =  Set Boolean value that indicates if pbSettings worksheet is automatically
''      Hidden (default: True)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let stdAutoHide(ByVal hideSheetDefault As Boolean)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        On Error Resume Next
        Setting(PBSTG_HIDE) = hideSheetDefault
        
        '' Add/Remove the settings worksheet from the DO_NOT_HIDE_SHEETCODENAMES key
        Dim pbk As pbKey
        pbk = NewKey(PBSTG_DO_NOT_HIDE, optDefault)
        If hideSheetDefault Then
            ArraySettingRemoveFrom pbk.Key, pbk.OldOptions, pbSettingsSheet.CodeName
        Else
            ArraySettingAppendTo pbk.Key, pbk.OldOptions, pbSettingsSheet.CodeName
        End If
        '' Ensure PBSTG_HIDE Is Boolean
        If Not SettingType(PBSTG_HIDE) = teBoolean Then ForceBooleanFormat (PBSTG_HIDE)
        '' Ensure Settings Sheet Is Visible or VeryHidden
        If hideSheetDefault = False And l_ws.visible <> xlSheetVisible Then
            ''  Disable App Events
            Dim evts As Boolean: evts = Me.Events: Me.Events = False
            l_ws.visible = xlSheetVisible
            SizeColumns
            ''  Restore App Events to prior setting
            Me.Events = evts
        Else
            CheckSettingsHealth
        End If
    End Property
    Public Property Get stdArrayDoNotHideSheets() As Variant()
        stdArrayDoNotHideSheets = GetValue(PBSTG_DO_NOT_HIDE)
    End Property
    Public Property Let stdArrayDoNotHideSheets(doNotHide() As Variant)
        Setting(PBSTG_DO_NOT_HIDE) = doNotHide
    End Property
    
    
'            If hideSheetDefault Then
'            ArraySettingRemoveFrom pbk.Key, pbk.OldOptions, pbSettingsSheet.CodeName
'        Else
'            ArraySettingAppendTo pbk.Key, pbk.OldOptions, pbSettingsSheet.CodeName
'        End If

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Get Delimiter for Settings Values that return as arrays - pbSettings Standard Setting
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get stdArrayDelimiter() As String
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Dim resp: resp = Setting(PBSTG_ARRAY_DELIMITER)
        If Len(resp) = 0 Then
            Setting(PBSTG_ARRAY_DELIMITER) = "|"
            resp = "|"
        End If
        stdArrayDelimiter = resp
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Let Delimiter for Settings Values that return as arrays - pbSettings Standard Setting
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let stdArrayDelimiter(ByVal arrDelimiter As String)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        If Len(CStr(arrDelimiter)) > 0 Then
            Setting(PBSTG_ARRAY_DELIMITER) = arrDelimiter
        End If
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Get Font Name Used on pbSettings Worksheet - pbSettings Standard Setting
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get stdSettingsFontName() As String
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Dim resp: resp = GetValue(PBSTG_FONT_NAME)
        If Len(resp) = 0 Then
            On Error Resume Next
            resp = stgwb.Styles("Normal").Font.Name
            If Err.number <> 0 Then
                Err.Clear
                resp = pbSettingsSheet.Range("A1").Font.Name
            End If
            On Error GoTo 0
            Setting(PBSTG_FONT_NAME) = resp
        End If
        stdSettingsFontName = resp
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Set Font Name Used on pbSettings Worksheet - pbSettings Standard Setting
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let stdSettingsFontName(ByVal fontName As String)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        If Len(CStr(fontName)) > 0 Then
            On Error Resume Next
            pbSettingsSheet.Cells.Font.Name = fontName
            If Err.number = 0 Then
                Setting(PBSTG_FONT_NAME) = fontName
            Else
                Err.Clear
                RaiseEvent OnAlert(True, "Error Setting pbSettingsSheet Font Name to: " & fontName)
            End If
        End If
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Get Font Size Used on pbSettings Worksheet - pbSettings Standard Setting
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get stdSettingsFontSize() As Double
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Dim resp: resp = GetValue(PBSTG_FONT_SIZE)
        If Len(resp) = 0 Then
            On Error Resume Next
            resp = stgwb.Styles("Normal").Font.Size
            If Err.number <> 0 Then
                Err.Clear
                resp = pbSettingsSheet.Range("A1").Font.Size
            End If
            On Error GoTo 0
            Setting(PBSTG_FONT_SIZE) = resp
        End If
        stdSettingsFontSize = resp
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Set Font Size Used on pbSettings Worksheet - pbSettings Standard Setting
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let stdSettingsFontSize(ByVal fontSize As Double)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        If Len(CStr(fontSize)) > 0 Then
            On Error Resume Next
            pbSettingsSheet.Cells.Font.Size = fontSize
            If Err.number = 0 Then
                Setting(PBSTG_FONT_SIZE) = fontSize
            Else
                Err.Clear
                RaiseEvent OnAlert(True, "Error Setting pbSettingsSheet Font Size to: " & fontSize)
            End If
        End If
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Get Max Column Width for Settings Sheet - pbSettings Standard Setting
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get stdMaxSettingsSheetColWidth() As Long
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Dim resp: resp = Setting(PBSTG_MAX_SETTINGS_COL_WIDTH)
        If Len(resp) = 0 Then
            Setting(PBSTG_MAX_SETTINGS_COL_WIDTH) = 70
            resp = 70
        End If
        stdMaxSettingsSheetColWidth = resp
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Let Max Column Width for Settings Sheet - pbSettings Standard Setting
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let stdMaxSettingsSheetColWidth(ByVal maxColWidth As Long)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        If maxColWidth > 0 Then
            Setting(PBSTG_MAX_SETTINGS_COL_WIDTH) = maxColWidth
        End If
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Get Workbook Disable AutoSave - pbSettings Standard Setting (Default True)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get stdDisableAutoSave() As Boolean
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Dim resp: resp = Setting(PBSTG_DISABLE_AUTO_SAVE)
        If Len(resp) = 0 Then
            Setting(PBSTG_DISABLE_AUTO_SAVE) = True
            resp = True
        End If
        stdDisableAutoSave = resp
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Let Workbook Disable AutoSave - pbSettings Standard Setting (Default True)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let stdDisableAutoSave(ByVal disableWkbkAutoSave As Boolean)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Setting(PBSTG_DISABLE_AUTO_SAVE) = disableWkbkAutoSave
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Get Workbook EnableConnections - pbSettings Standard Setting (Default True)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get stdEnableWkbkConnections() As Boolean
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Dim resp: resp = Setting(PBSTG_ENABLE_CONNECTIONS)
        If Len(resp) = 0 Then
            Setting(PBSTG_ENABLE_CONNECTIONS) = True
            resp = True
        End If
        stdEnableWkbkConnections = resp
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Let Workbook EnableConnections - pbSettings Standard Setting (Default True)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let stdEnableWkbkConnections(ByVal enableWkbkConnections As Boolean)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Setting(PBSTG_ENABLE_CONNECTIONS) = enableWkbkConnections
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Get Workbook SaveLinkedValues - pbSettings Standard Setting (Default True)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get stdSaveLinkedValues() As Boolean
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Dim resp: resp = Setting(PBSTG_SAVE_LINK_VALUES)
        If Len(resp) = 0 Then
            Setting(PBSTG_SAVE_LINK_VALUES) = True
            resp = True
        End If
        stdSaveLinkedValues = resp
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Let Workbook Workbook SaveLinkedValues - pbSettings Standard Setting (Default True)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let stdSaveLinkedValues(ByVal saveLinkedWkbkValues As Boolean)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Setting(PBSTG_SAVE_LINK_VALUES) = saveLinkedWkbkValues
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  ** DEVELOPER MENU/TOOLBAR ITEMS WILL NEVER BE HIDDEN BY pbSettings **
''  Get Value (Boolean) That Determines if pbSettings will automatically ENABLE/SHOW
''  the built-in Developer Excel Menu/Toolbar items
''  (If setting is False, no action is taken by pbSettings)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get stdShowDevTools() As Boolean
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Dim resp: resp = GetValue(PBSTG_DEV_SHOW_TOOLS, False, optCreateKeyIfMissing + optNonMigratableKey + optValueAsBoolean)
        If Len(resp) = 0 Then
            Setting(PBSTG_DEV_SHOW_TOOLS) = False
            resp = False
        End If
        stdShowDevTools = resp
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  ** DEVELOPER MENU/TOOLBAR ITEMS WILL NEVER BE HIDDEN BY pbSettings **
''  Set Value (Boolean) That Determines if pbSettings will automatically ENABLE/SHOW
''  the built-in Developer Excel Menu/Toolbar items
''  (If setting is False, no action is taken by pbSettings)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let stdShowDevTools(ByVal showDevTools As Boolean)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Setting(PBSTG_DEV_SHOW_TOOLS) = showDevTools
    End Property


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Get 'ShowLogTrace' - pbSettings DEV Setting (Default True)
''  Use, add the following line of code to the method used to log to log file
''  (pseudo-code -- adjust as necessary)
''      If stg.stdDEVShowLogTrace = True Then
''          Debug.Print [log message]
''      Endif
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get stdDEVShowLogTrace() As Boolean
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Dim resp: resp = Setting(PBSTG_DEV_LOG_TO_IMMEDIATE)
        If Len(resp) = 0 Then
            Setting(PBSTG_DEV_LOG_TO_IMMEDIATE) = False
            resp = False
        End If
        stdDEVShowLogTrace = resp
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Let 'ShowLogTrace' - pbSettings DEV Setting (Default False)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let stdDEVShowLogTrace(ByVal showLogTraceOutput As Boolean)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Setting(PBSTG_DEV_LOG_TO_IMMEDIATE) = showLogTraceOutput
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Get 'IsDevMode' - pbSettings DEV Setting (Default True)
''  Use to check if your preferred "developer mode" configuration should be called
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get stdDEVIsDevMode() As Boolean
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Dim resp: resp = Setting(PBSTG_DEV_LOG_TO_IMMEDIATE)
        If Len(resp) = 0 Then
            Setting(PBSTG_DEV_LOG_TO_IMMEDIATE) = False
            resp = False
        End If
        stdDEVIsDevMode = resp
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Let 'IsDevMode' - pbSettings DEV Setting (Default True)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let stdDEVIsDevMode(ByVal devMode As Boolean)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Setting(PBSTG_DEV_LOG_TO_IMMEDIATE) = devMode
    End Property
    
    Public Function stdDevNamesAppend(ParamArray addNames() As Variant)
        '' make sure duplicates are not allowed
        If privEnumCompare(SettingOptions(PBSTG_DEV_NAMES), pbSettingOption.optAllowDuplicateArrayValues) Then
            Dim updOptions As pbSettingOption
            updOptions = privEnumModify(SettingOptions(PBSTG_DEV_NAMES), pbSettingOption.optAllowDuplicateArrayValues, feVerifyEnumRemoved)
            SettingOptions(PBSTG_DEV_NAMES) = updOptions
        End If
        Dim addName As Variant
        For Each addName In addNames
            ArraySettingAppendTo PBSTG_DEV_NAMES, optDefault, addName
        Next addName
    End Function
    Public Function stdDevNamesRemove(ParamArray removeNames() As Variant)
        '' make sure duplicates are not allowed
        Dim removeName As Variant
        For Each removeName In removeNames
            ArraySettingRemoveFrom PBSTG_DEV_NAMES, optDefault, removeName
        Next removeName
    End Function

    Public Function stdAdminNamesAppend(ParamArray addNames() As Variant)
        '' make sure duplicates are not allowed
        If privEnumCompare(SettingOptions(PBSTG_ADMIN_NAMES), pbSettingOption.optAllowDuplicateArrayValues) Then
            Dim updOptions As pbSettingOption
            updOptions = privEnumModify(SettingOptions(PBSTG_ADMIN_NAMES), pbSettingOption.optAllowDuplicateArrayValues, feVerifyEnumRemoved)
            SettingOptions(PBSTG_ADMIN_NAMES) = updOptions
        End If
        Dim addName As Variant
        For Each addName In addNames
            ArraySettingAppendTo PBSTG_ADMIN_NAMES, optDefault, addName
        Next addName
    End Function
    Public Function stdAdminNamesRemove(ParamArray removeNames() As Variant)
        '' make sure duplicates are not allowed
        Dim removeName As Variant
        For Each removeName In removeNames
            ArraySettingRemoveFrom PBSTG_ADMIN_NAMES, optDefault, removeName
        Next removeName
    End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  EXECUTE BUTTON ACTION
''  IF A SETTING HAS BEEN CREATED FOR THE [btnName] and .CodeName of [wksht]
''  Then the public sub or function will be called in the workbook that contains
''  [wksht] reference
''
''  EXAMPLE USAGE
''      A 'ButtonAction' Setting is created with:
''      > stg.stdButtonAction(wsDashboard, "btnDoSomething") = "AddInvoice"
''          (Would create or update Setting with key: "BTN_ACT_wsDashboard_btnDoSomething"
''          (Setting Value would be set to: "AddInvoice")
''      When a button is clicked that you have specified a custom 'OnAction' method that
''          will intentionally use pbSettings to execute the call, from your 'OnAction'
''          method,you will use the following to execute the methodName ( stored in settings.
''      > Dim clickedButtonName as string
''      > If TypeName(Application.Caller) = "String" Then
''      >     clickedButtonName = Application.Caller
''      >     stg.ExecuteButtonAction clickedButtonName, ThisWorkbook.ActiveSheet
''      > Endif
''
''      For the above example, if worksheet.codeName 'wsDashboard' was in a workbook
''          named "My Workbook.xlsm", the code that would be executed would be:
''      > Application.Run " 'My Workbook.xlsm'!AddInvoice "
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ExecuteButtonAction(ByVal btnName, wksht As Worksheet)
        On Error Resume Next
        Dim callMethod As String
        callMethod = stdButtonAction(wksht, btnName)
        Dim wkbk As Workbook
        Set wkbk = wksht.Parent
        ''
        If Len(callMethod) = 0 Then
            LogERROR "pbSettingButtonAction not found for: " & btnName & ", from: " & wksht.CodeName & " (" & wksht.Name & ")"
        Else
            LogTRACE "pbSettingButtonAction: " & callMethod & ", running for: " & btnName & ", from: " & wksht.CodeName & " (" & wksht.Name & ")"
            Application.Run "'" & wkbk.Name & "'!" & callMethod
        End If
        ''
        If Err.number <> 0 Then
            RaiseEvent OnAlert(True, "pbSettings.ExecuteButtonAction called from '" & wksht.CodeName & ":" & btnName & "' to execute method: " & callMethod)
        Else
            If Len(callMethod) > 0 Then
                RaiseEvent OnAlert(False, "pbSettings.ExecuteButtonAction called from '" & wksht.CodeName & ":" & btnName & "' to execute method: " & callMethod)
            End If
        End If
        Set wkbk = Nothing
    End Function
    Private Function ButtonActionKey(ByVal wkshtCodeName, ByVal btnName)
        ButtonActionKey = BUTTON_ACTION_PREFIX & "_" & wkshtCodeName & "_" & btnName
    End Function
    
    Public Function GetValue( _
        ByVal stgKey _
        , Optional defaultValue _
        , Optional options As pbSettingOption = pbSettingOption.optCreateKeyIfMissing)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        Dim pbk As pbKey
        pbk = NewKey(stgKey, options)
        GetValue = pbk.OldValFormatted
    End Function
    
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Setting Value Property Getter
''  If [stgKey] does not exist, returns an Empty Type
''  e.g. If IsEmpty(STG.Setting("invalidkey")) Then ....
''  If Setting SettingType is '3' (SettingTypeEnum.teDateTime = 3), the value will
''    be converted using 'CDate' before it is returned
''  If Setting SettingType is '1' (SettingTypeEnum.teNumeric = 1), the value will
''    be returned using:  [settingValue] = Val([settingValue])
''  If Setting SettingType is '2' (SettingTypeEnum.teBoolean = 2), the value will
''      be return using: [settingValue] = CBool([settingValue])
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get Setting(ByVal stgKey)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Dim tmpIdx As Long: tmpIdx = Index(stgKey)
        Dim resp
        If tmpIdx > 0 Then
            Dim stgArr As Variant
            stgArr = l_lo.listRows(tmpIdx).Range.value
            resp = FormatValue(stgArr(1, pbSTGColumns.scSettingValue), stgArr(1, pbSTGColumns.scSettingOptions))
        End If
        Setting = resp
    End Property
    

    
    Private Function OSSpecific(options As pbSettingOption) As Boolean
        OSSpecific = privEnumCompare(options, optOSSpecific + optUserAndOS, ecOR)
    End Function
    Private Function USERSpecific(options As pbSettingOption) As Boolean
        USERSpecific = privEnumCompare(options, optUserSpecific + optUserAndOS, ecOR)
    End Function
    Private Function IsArraySet(options As pbSettingOption) As Boolean
        IsArraySet = privEnumCompare(options, optValueAsArray, ecOR)
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  ArraySettingAppend
''  If [stgKey] does not exist, an exception will be raised
''      will be '0' (SettingTypeEnum.teGeneral = 0)
''  This Method Will check each value provided, and append to existing
''  setting with "|" (or modified default array delimeter) as delimiter
''
'' Returns modified variant array
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ArraySettingAppendTo(ByVal stgKey, options As pbSettingOption, ParamArray appendVals() As Variant)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        'if stg is not type array, then raise error
        Dim validKey: validKey = CheckKey(stgKey, options)
        Dim keyIndex As Long: keyIndex = Index(validKey)
        Dim existOptions As pbSettingOption
        If keyIndex > 0 Then
            existOptions = SettingOptionsByIndex(keyIndex)
            If Not privEnumCompare(existOptions, optValueAsArray) Then
                Err.Raise 13, Source:=Me, Description:="pbSettings.ArraySettingAppendTo: " & validKey & " key is not configured as an array"
                Exit Function
            End If
        Else
            Err.Raise 13, Source:=Me, Description:="ArraySettingAppendTo requires existing key.  " & stgKey & " was not found"
            Exit Function
        End If

        Dim tempText As String
        tempText = SettingValueRaw(validKey, keyIndex)
        
        
        Dim arrDelim As String: arrDelim = stdArrayDelimiter
        Dim itemsColl As New Collection
        If Len(tempText) > 0 Then
            Dim tmpItems: tmpItems = Setting(validKey)
            Dim existItem, newItem, shouldAdd As Boolean
            For Each existItem In tmpItems
                ConvertItemFormat existItem, existOptions
                itemsColl.add existItem, Key:=CStr(existItem)
            Next existItem
        End If
        For Each newItem In appendVals
            ConvertItemFormat newItem, existOptions
            ''Add New value to collect if it doesn't already exist, OR if 'optAllowDuplicationArrayValues' option exist on setting
            If privEnumCompare(existOptions, pbSettingOption.optAllowDuplicateArrayValues) Or Not privCollectionKeyExists(itemsColl, CStr(newItem)) Then
                itemsColl.add newItem, Key:=CStr(newItem)
            End If
        Next newItem
        Dim newArr()
        ReDim newArr(1 To itemsColl.Count)
        Dim iCounter As Long
        For Each newItem In itemsColl
            iCounter = iCounter + 1
            newArr(iCounter) = newItem
        Next
        ArraySetting(stgKey) = newArr
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  ArraySettingRemoveItems
''  If [stgKey] does not exist, it will be created, and the SettingType
''      will be '0' (SettingTypeEnum.teGeneral = 0)
''  This Method Will check each value provided, and append to existing
''  setting with "|" as delimiter
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ArraySettingRemoveFrom(ByVal stgKey, options As pbSettingOption, ParamArray removeVals() As Variant)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        Dim validKey: validKey = CheckKey(stgKey, options)
        Dim keyIndex As Long: keyIndex = Index(validKey)
        Dim existOptions As pbSettingOption
        If keyIndex > 0 Then
            existOptions = l_lo.listRows(keyIndex).Range(1, pbSTGColumns.scSettingOptions).value
            If Not privEnumCompare(existOptions, optValueAsArray) Then
                Err.Raise 13, Source:=Me, Description:="pbSettings.ArraySettingRemoveFrom: " & validKey & " key is not configured as an array"
                Exit Function
            End If
        Else
            Err.Raise 13, Source:=Me, Description:="ArraySettingRemoveFrom requires existing key.  " & stgKey & " was not found"
            Exit Function
        End If
        
        Dim arrDelim As String: arrDelim = stdArrayDelimiter
        Dim itemsColl As New Collection
        Dim tmpItems: tmpItems = Setting(validKey)
        Dim existItem, newItem, shouldAdd As Boolean
        For Each existItem In tmpItems
            itemsColl.add existItem, Key:=CStr(existItem)
        Next existItem
        For Each newItem In removeVals
            ConvertItemFormat newItem, existOptions
            If privCollectionKeyExists(itemsColl, CStr(newItem)) Then
                itemsColl.remove CStr(newItem)
            End If
        Next newItem
        Dim newArr()
        ReDim newArr(1 To itemsColl.Count)
        Dim iCounter As Long
        For Each newItem In itemsColl
            iCounter = iCounter + 1
            newArr(iCounter) = newItem
        Next
        ArraySetting(validKey) = newArr
    End Function
    
    Private Function ConcatCollectionItems(ByRef col As Collection, Optional delimitBy As String = "|") As String
        Dim resp As String
        Dim colItem
        For Each colItem In col
            If resp = vbNullString Then
                resp = CStr(colItem)
            Else
                resp = privConcatWithDelim(delimitBy, resp, colItem)
            End If
        Next colItem
        ConcatCollectionItems = resp
    End Function
    
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    '   Returns item from collection by Key
    '   If [key] does not exist in collection, error object with
    '   error code 1004 is return
    '   suggested use:
    '
    '   Dim colItem as Variant
    '   colItem = CollectionItemByKey([collection], [expectedKey])
    '
    '   'If expecting object, use 'Set'
    '    Set colItem = CollectionItemByKey([collection], [expectedKey])
    '
    '   If Not IsError(colItem) Then
    '       'value was returned
    '   Else
    '       'error was returned
    '   End if
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function privCollectionItemByKey(ByRef col As Collection, ByVal Key)
    On Error Resume Next
        Key = CStr(Key)
        If IsObject(col(Key)) Then
            If Err.number = 0 Then
                Set privCollectionItemByKey = col(Key)
            End If
        Else
            If Err.number = 0 Then
                privCollectionItemByKey = col(Key)
            End If
        End If
        If Err.number <> 0 Then
            Err.Clear
            privCollectionItemByKey = CVErr(1004)
        End If
    End Function
    Private Function privCollectionKeyExists(ByRef col As Collection, ByVal Key)
    On Error Resume Next
        Key = CStr(Key)
        If IsObject(col(Key)) Then
            If Err.number = 0 Then
                privCollectionKeyExists = True
            Else
                privCollectionKeyExists = False
            End If
        Else
            If Err.number = 0 Then
                privCollectionKeyExists = True
            Else
                privCollectionKeyExists = False
            End If
        End If
        If Err.number <> 0 Then
            Err.Clear
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Setting Value Property Setter
''  If [stgKey] does not exist, it will be created, and the SettingType
''      will be '0' (SettingTypeEnum.teGeneral = 0)
''  Typically you should not have to modify the SettingType unless for
''      some reason you'd be create the setting with a 'Value2' value, and
''      would need to return it with a different format
''      e.g. if you created a setting with the value: 45263.8997685185,
''          but would need to return it as a date: 12/3/23 9:35:40 PM
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let Setting(ByVal stgKey, ByVal stgVal)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        On Error Resume Next
        Dim tmpIdx As Long: tmpIdx = Index(stgKey)
        Dim validKey: validKey = CheckKey(stgKey)
        Dim resp, respType
        Dim options As pbSettingOption
        'Dim stType As SettingTypeEnum
        
        ''  Disable App Events
        Dim evts As Boolean: evts = Me.Events: Me.Events = False
        If tmpIdx > 0 Then
            'Update Setting
            options = l_lo.listRows(tmpIdx).Range(1, pbSTGColumns.scSettingOptions).value
            
            'stType = SettingTypeByOptions(options)
            stgVal = FormatValue(stgVal, options)
            
            If privEnumCompare(VarType(stgVal), VbVarType.vbArray) And Not privEnumCompare(options, pbSettingOption.optValueAsArray) Then
                RaiseEvent OnAlert(False, "pbSettings.Setting (LET): [stgVal] was found to be of type 'Array', but the setting (" & stgKey & ") was originally set to another data type. SettingOption was forcibly changed to an array")
                options = privEnumModify(options, pbSettingOption.optValueAsArray, FlagEnumModifyPriv.feVerifyEnumExists)
                 l_lo.listRows(tmpIdx).Range(1, pbSTGColumns.scSettingOptions).value = options
                ArraySetting(validKey) = stgVal
                Exit Property
            End If
            If privEnumCompare(options, pbSettingOption.optValueAsArray) Then
                ArraySetting(validKey) = stgVal
                Exit Property
            End If
            RaiseEvent OnModifySetting(validKey, stgVal, False)
            If Not StringsMatchPriv(TypeName(stgVal), TypeName(l_lo.listRows(tmpIdx).Range(1, pbSTGColumns.scSettingValue).value)) Then
                l_lo.listRows(tmpIdx).Range(1, pbSTGColumns.scSettingValue).numberFormat = "General"
            End If
            l_lo.listRows(tmpIdx).Range(1, pbSTGColumns.scSettingValue).value = ConvertItemFormat(stgVal, options)
            l_lo.listRows(tmpIdx).Range(1, pbSTGColumns.scUpdated).value = Now()
        Else
            If privEnumCompare(VarType(stgVal), VbVarType.vbArray) Then
                options = optValueAsArray + optValueAsGeneral
            Else
                options = optValueAsGeneral
            End If
            RaiseEvent OnModifySetting(validKey, stgVal, True)
            Dim stgArr() As Variant
            Dim newRow As ListRow
            ReDim stgArr(1 To 1, 1 To l_lo.ListColumns.Count)
            stgArr(1, pbSTGColumns.scSettingKey) = "'" & validKey
            stgArr(1, pbSTGColumns.scSettingValue) = ConvertItemFormat(stgVal, options)
            stgArr(1, pbSTGColumns.scUpdated) = Now()
            stgArr(1, pbSTGColumns.scSettingOptions) = options
            Set newRow = l_lo.listRows.add
            newRow.Range.value = stgArr
            newRow.Range(1, pbSTGColumns.scSettingValue).numberFormat = "General"
            CheckSort
        End If
        l_isDirty = True
        If Err.number <> 0 Then
            Err.Clear
        End If
        Me.Events = evts
    End Property
    
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
    
    Public Property Let ArraySetting(ByVal stgKey, ByVal stgVal)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        On Error Resume Next
        Dim tmpIdx As Long: tmpIdx = Index(stgKey)
        Dim validKey: validKey = CheckKey(stgKey)
        Dim resp, respType
        ''  Disable App Events
        Dim evts As Boolean: evts = Me.Events: Me.Events = False
        ''  If [stgVal] is not array, will convert to array
        
        Dim arrVal As Variant, arrDelim As String, newStgVal As String
        arrDelim = stdArrayDelimiter
        If privEnumCompare(VarType(stgVal), VbVarType.vbArray) = False Then
            arrVal = Array(stgVal)
        Else
            arrVal = stgVal
        End If
        
        Dim options As pbSettingOption
        If tmpIdx > 0 Then
            options = l_lo.listRows(tmpIdx).Range(1, pbSTGColumns.scSettingOptions).value
        Else
            options = pbSettingOption.optValueAsArray + pbSettingOption.optValueAsGeneral
        End If
        
        
        Dim tempVal As Variant, arrIdx As Long
        arrIdx = LBound(arrVal)
        For Each tempVal In arrVal
            arrVal(arrIdx) = ConvertItemFormat(tempVal, options)
            arrIdx = arrIdx + 1
        Next tempVal
        
        newStgVal = Join(arrVal, arrDelim)
        
        If tmpIdx > 0 Then
            'Update Setting
            If Not privEnumCompare(options, optValueAsArray) Then
                Err.Raise 13, Description:="Cannot add array values to non array-type setting "
                Exit Property
            End If
            RaiseEvent OnModifySetting(validKey, newStgVal, False)
            l_lo.listRows(tmpIdx).Range(1, pbSTGColumns.scSettingValue).value = newStgVal
            l_lo.listRows(tmpIdx).Range(1, pbSTGColumns.scUpdated).value = Now()
        Else
            RaiseEvent OnModifySetting(validKey, newStgVal, True)
            Dim stgArr() As Variant
            ReDim stgArr(1 To 1, 1 To l_lo.ListColumns.Count)
            stgArr(1, pbSTGColumns.scSettingKey) = stgKey
            stgArr(1, pbSTGColumns.scSettingValue) = newStgVal
            stgArr(1, pbSTGColumns.scUpdated) = Now()
            stgArr(1, pbSTGColumns.scSettingOptions) = options
            
            l_lo.listRows.add().Range.value = stgArr
            CheckSort
        End If
        l_isDirty = True
        If Err.number <> 0 Then
            Err.Clear
        End If
        Me.Events = evts
    End Property
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Setting Property Getter for OS-specific setting
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Property Get SettingForOS(ByVal stgKey)
        SettingForOS = Setting(BuildOSSpecificKey(stgKey))
    End Property
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Setting Property Setter for OS-specific setting
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Property Let SettingForOS(ByVal stgKey, ByVal stgVal)
        Setting(BuildOSSpecificKey(stgKey)) = stgVal
    End Property
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Setting Property Getter for USER-specific setting
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get SettingForUser(ByVal stgKey)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        SettingForUser = Setting(BuildUSERSpecificKey(stgKey))
    End Property
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Setting Property Setter for USER-specific setting
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Let SettingForUser(ByVal stgKey, ByVal stgVal)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Setting(BuildUSERSpecificKey(stgKey)) = stgVal
    End Property
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Attempt to convert Long to string key for StandardSettingEnum
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function GetStandardSettingKey(ByVal cKey As StandardSettingEnum) As String
        Dim resp As String, tmpCheckKey As Long
        If IsNumeric(cKey) Then
            tmpCheckKey = CLng(cKey)
            If tmpCheckKey >= StandardSettingEnum.[_StandardSettingKeyFirst] And _
                tmpCheckKey <= StandardSettingEnum.[_StandardSettingKeyLast] Then
                Select Case tmpCheckKey
                    Case StandardSettingEnum.stgArrayDoNotProtectSheets
                        resp = PBSTG_DO_NOT_PROTECT
                    Case StandardSettingEnum.stgArraySettingDelimiter
                        resp = PBSTG_ARRAY_DELIMITER
                    Case StandardSettingEnum.stgAutoHideSettingsSheet
                        resp = PBSTG_HIDE
                    Case StandardSettingEnum.stgDEVIsDevMode
                        resp = PBSTG_DEV_MODE
                    Case StandardSettingEnum.stgDEVNames
                        resp = PBSTG_DEV_NAMES
                    Case StandardSettingEnum.stgDEVRunAsNoNetwork
                        resp = PBSTG_DEV_NO_NETWORK
                    Case StandardSettingEnum.stgDEVShowLogTrace
                        resp = PBSTG_DEV_LOG_TO_IMMEDIATE
                    Case StandardSettingEnum.stgDisableAutoSave
                        resp = PBSTG_DISABLE_AUTO_SAVE
                    Case StandardSettingEnum.stgEnableConnections
                        resp = PBSTG_ENABLE_CONNECTIONS
                    Case StandardSettingEnum.stgSaveLinkedValues
                        resp = PBSTG_SAVE_LINK_VALUES
                    Case StandardSettingEnum.stgSettingsSheetMaxColWidth
                        resp = PBSTG_MAX_SETTINGS_COL_WIDTH
                    Case StandardSettingEnum.stgShowDevTools
                        resp = PBSTG_DEV_SHOW_TOOLS
                    Case Else
                        resp = vbNullString
                End Select
            Else
            GetStandardSettingKey = resp
            End If
        End If
        GetStandardSettingKey = resp
    End Function
    
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Exposed for convenience -- return name of setting key
''  The Only 'pbSettingOption' values that are used are:
''      pbSettingOption.optOSSpecific
''      pbSettingOption.optUserSpecific
''      pbSettingOption.optUserAndOS
''  If [stgKey] is Numeric, and a valid StandardSettingEnum option, a string key
''      will be returned.  e.g.
''      myKey = stg.CheckKey(StandardSettingEnum.stgAutoHideSettingsSheet)
''          Result is myKey = "PBSTG_AUTOHIDE"
''  if [stgkey] is Numeric and is not a valid StandardSettingEnum, or
''      if [stgKey] is any other non-string value, like date, boolena, it will be converted
''      to a string.  e.g.  6000 becomes "6000", TRUE becomes "TRUE", etc.
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function CheckKey(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        Dim forceOSSpecific As Boolean, forceUSERSpecific As Boolean
        forceOSSpecific = privEnumCompare(options, optOSSpecific + optUserAndOS, ecOR)
        forceUSERSpecific = privEnumCompare(options, optUserSpecific + optUserAndOS, ecOR)
        
        If Not StringsMatchPriv(TypeName(stgKey), "String") Then
            On Error Resume Next
            Dim stdKey: stdKey = GetStandardSettingKey(CLng(stgKey))
            If stdKey = vbNullString Then
                stgKey = CStr(stgKey)
            Else
                stgKey = stdKey
            End If
            On Error GoTo 0
        End If
        
        stgKey = CStr(stgKey)
        If forceUSERSpecific Or StringsMatchPriv(stgKey, "USER_", smStartsWithStr) Then
            stgKey = BuildUSERSpecificKey(stgKey)
        End If
        If forceOSSpecific Or StringsMatchPriv(stgKey, "_OS", smEndWithStr) Then
            stgKey = BuildOSSpecificKey(stgKey)
        End If
        CheckKey = stgKey
    End Function
'''    Public Function CheckKey(ByVal stgKey, Optional forceOSSpecific As Boolean = False, Optional forceUSERSpecific As Boolean = False)
'''        If Not ValidConfig Then Exit Function
'''        '' if [stgKey] is not a string, then:
'''        ''  1. If it is a Long, then check if between 1 and StandardSettingEnum.[_stgMaxVal]
'''        ''  2. Otherwise Convert To String
'''        Dim tmpKey
'''        If IsNumeric(stgKey) Then
'''            Dim stdSettingKey As String
'''            stdSettingKey = GetStandardSettingKey(stgKey)
'''            If Len(stdSettingKey) > 0 Then
'''                stgKey = stdSettingKey
'''            End If
'''        End If
'''        tmpKey = CStr(stgKey)
'''        'If isOSSpecific And isUSERSpecific Then Err.Raise 1004, Description:="A pbSetting Key cannot be both OS and User specific"
'''        ''
'''        '' If [stgKey] ends with "_OS", will automatically be changed to end with "_OS_MAC" or "_OS_PC"
'''        '' if [stgKey] starts with "USER_", will automatically be change to start with "USER_[username]_"
'''        If forceUSERSpecific Or StringsMatchPriv(tmpKey, "USER_", smStartsWithStr) Then
'''            tmpKey = BuildUSERSpecificKey(tmpKey)
'''        End If
'''        If forceOSSpecific Or StringsMatchPriv(tmpKey, "_OS", smEndWithStr) Then
'''            tmpKey = BuildOSSpecificKey(tmpKey)
'''        End If
'''        CheckKey = tmpKey
'''    End Function
    
'    Private Function CheckValueFormat(ByVal stgVal, ByVal stgType)
'        On Error Resume Next
'        Dim resp: resp = stgVal
'        If Not stgType = vbNullString Then
'            Select Case CLng(stgType)
''                Case SettingTypeEnum
'.teBoolean
''                    resp = CBool(resp)
''                    If Err.number <> 0 Then
''                        Err.Clear
''                        resp = False
''                    End If
''                Case SettingTypeEnum
'.teDate
''                    resp = dtPartPriv(dtDate_NoTime, CDate(resp))
''                    If Err.number <> 0 Then
''                        Err.Clear
''                        resp = vbNullString
''                    End If
''                Case SettingTypeEnum
'.teDatetime
''                    resp = CDate(resp)
''                    If Err.number <> 0 Then
''                        Err.Clear
''                        resp = vbNullString
''                    End If
''                Case SettingTypeEnum
'.teLong
''                    If CDbl(resp) <> CLng(resp) Then
''                        Dim lngString As String: lngString = CStr(resp)
''                        resp = CLng(Mid(lngString, 1, InStr(1, lngString, ".", vbTextCompare) - 1))
''                    Else
''                        resp = CLng(resp)
''                    End If
''                    If Err.number <> 0 Then
''                        Err.Clear
''                        resp = CLng(0)
''                    End If
''                Case SettingTypeEnum
'.teDouble
''                    resp = CDbl(resp)
''                    If Err.number <> 0 Then
''                        Err.Clear
''                        resp = CDbl(0)
''                    End If
''                Case SettingTypeEnum
'.teArray
''                    If privEnumCompare(VarType(resp), VbVarType.vbArray) = False Then
''                        resp = vbNullString
''                    End If
''            End Select
''        End If
''        CheckValueFormat = resp
''    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns timestamp when Setting [stgKey] was last modified
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function SettingUpdated(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        If Exists(stgKey) Then
            SettingUpdated = CDate(l_lo.listRows(Index(stgKey)).Range(1, pbSTGColumns.scUpdated))
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Returns True if SettingOptions includes 'pbSettingOption.optKeyRequired'
''      A Setting Cannot Be Deleted If it includes optKeyRequired
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function SettingIsRequired(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault) As Boolean
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        If Exists(stgKey) Then
            SettingIsRequired = privEnumCompare(SettingOptions(stgKey), pbSettingOption.optKeyRequired)
        End If
    End Function
    
    
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns SettingType
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function SettingTypeByOptions(options As pbSettingOption) As SettingTypeEnum
            If privEnumCompare(options, optValueAsBoolean) Then
                SettingTypeByOptions = teBoolean
            ElseIf privEnumCompare(options, optValueAsDate) Then
                SettingTypeByOptions = teDate
            ElseIf privEnumCompare(options, optValueAsDateTime) Then
                SettingTypeByOptions = teDatetime
            ElseIf privEnumCompare(options, optValueAsDouble) Then
                SettingTypeByOptions = teDouble
            ElseIf privEnumCompare(options, optValueAsLong) Then
                SettingTypeByOptions = teLong
            ElseIf privEnumCompare(options, optValueAsText) Then
                SettingTypeByOptions = teText
            Else
                SettingTypeByOptions = teGeneral
            End If
    End Function
    
    Private Function SettingValueRaw(ByVal validKey, Optional rowIdx As Long)
        If rowIdx <= 0 Or rowIdx > Count Then
            rowIdx = Index(validKey)
        End If
        If rowIdx > 0 Then
            SettingValueRaw = l_lo.listRows(rowIdx).Range(1, pbSTGColumns.scSettingValue).value
        End If
    End Function
    
    Private Function SettingOptionsByIndex(ByVal rowIdx As Long, Optional options As pbSettingOption = pbSettingOption.optDefault) As pbSettingOption
        SettingOptionsByIndex = l_lo.listRows(rowIdx).Range(1, pbSTGColumns.scSettingOptions).value
    End Function
    
    Public Property Get SettingOptions(ByVal stgKey) As pbSettingOption
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Dim tOptions As pbSettingOption
        stgKey = CheckKey(stgKey)
        Dim stgIDX As Long: stgIDX = Index(stgKey)
        If stgIDX > 0 Then
            SettingOptions = l_lo.listRows(stgIDX).Range(1, pbSTGColumns.scSettingOptions).value
        End If
    End Property
    
    Public Property Let SettingOptions(ByVal stgKey, stgOptions As pbSettingOption)
        If Not ValidConfig Then Exit Property
        stgKey = CheckKey(stgKey, stgOptions)
        Dim stgIDX As Long: stgIDX = Index(stgKey)
        If stgIDX > 0 Then
            l_lo.listRows(stgIDX).Range(1, pbSTGColumns.scSettingOptions).value = stgOptions
        ElseIf privEnumCompare(stgOptions, pbSettingOption.optCreateKeyIfMissing) Then
            Setting(stgKey) = vbNullString
            If Exists(stgKey) Then
                SettingOptions(stgKey) = stgOptions
            End If
        End If
    End Property
    
    Public Property Get SettingOption(ByVal stgKey, ByVal stgOpt As pbSettingOption) As Boolean
        stgKey = CheckKey(stgKey)
        If Exists(stgKey) Then
            SettingOption = privEnumCompare(SettingOptions(stgKey), stgOpt, ecAnd)
        Else
            SettingOption = False
        End If
    End Property
    Public Property Let SettingOption(ByVal stgKey, ByVal stgOpt As pbSettingOption, includeOption As Boolean)
        stgKey = CheckKey(stgKey)
        If Exists(stgKey) Then
            Dim curOptions As pbSettingOption, modOptions As pbSettingOption
            curOptions = SettingOptions(stgKey)
            modOptions = curOptions
            privEnumModify modOptions, stgOpt, IIf(includeOption, FlagEnumModifyPriv.feVerifyEnumExists, FlagEnumModifyPriv.feVerifyEnumRemoved)
            If curOptions <> modOptions Then
                SettingOptions(stgKey) = modOptions
            End If
        Else
            RaiseEvent OnAlert(True, "pbSettings.SettingOption 'Let' failed because Setting Key did not exist (" & stgKey & ")")
        End If
    End Property
    
    Public Function SettingType(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault) As SettingTypeEnum
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        If Exists(stgKey) Then
            Dim tOpt As pbSettingOption
            tOpt = l_lo.listRows(Index(stgKey)).Range(1, pbSTGColumns.scSettingOptions)
            If privEnumCompare(tOpt, optValueAsBoolean) Then
                SettingType = teBoolean
            ElseIf privEnumCompare(tOpt, optValueAsDate) Then
                SettingType = teDate
            ElseIf privEnumCompare(tOpt, optValueAsDateTime) Then
                SettingType = teDatetime
            ElseIf privEnumCompare(tOpt, optValueAsDouble) Then
                SettingType = teDouble
            ElseIf privEnumCompare(tOpt, optValueAsLong) Then
                SettingType = teLong
            ElseIf privEnumCompare(tOpt, optValueAsText) Then
                SettingType = teText
            Else
                SettingType = teGeneral
            End If
        Else
            SettingType = teUNKNOWN
        End If
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Check [stgKey] and append OS-specific string to key, if missing
''      e.g. on a PC,BuildOSSpecificKey("TEST") = "TEST_OS_PC"
''      e.g. on a PC,BuildOSSpecificKey("TEST_OS_PC") = "TEST_OS_PC"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function BuildOSSpecificKey(ByVal stgKey)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        Dim osKey
        Dim fullSuffix As String: fullSuffix = "_OS_" & PBSTG_OS
        If StringsMatchPriv(stgKey, fullSuffix, smEndWithStr) Then
            osKey = stgKey
        Else
            If StringsMatchPriv(stgKey, "_OS", smEndWithStr) Then
                osKey = Mid(stgKey, 1, Len(stgKey) - 3)
                osKey = osKey & fullSuffix
            Else
                osKey = stgKey & fullSuffix
            End If
        End If
        BuildOSSpecificKey = osKey
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Check [stgKey] and append USER-specific string to key, if missing
''      e.g. if user = 'browerp',BuildUSERSpecificKey("TEST") =
''          'TEST_USER_browerp"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function BuildUSERSpecificKey(ByVal stgKey)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        Dim userKey: userKey = CStr(stgKey)
        Dim fullPrefix As String: fullPrefix = "USER_" & UserNameOrLogin & "_"
        If Not StringsMatchPriv(userKey, fullPrefix, smStartsWithStr) Then
            If StringsMatchPriv(userKey, "USER_", smStartsWithStr) Then
                userKey = fullPrefix & Mid(userKey, 6)
            Else
                userKey = fullPrefix & userKey
            End If
        End If
        BuildUSERSpecificKey = userKey
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  HIDE ALL SHEETS THAT ARE:
''   - Not the current ActiveSheet
''   - Not In the DO_NOT_HIDE_SHEETCODENAMES Setting
''  @onlyHide - supply a single worksheet to hide.  (Will be ignored if it is the only
''      visible worksheet
''  @hideVeryHidden - True (Default) prevents users from unhiding. Set to false to
''      hide, but users are allowed to unhide
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function HideSheets(Optional onlyHide As Worksheet, Optional hideVeryHidden As Boolean = True)
        If Not onlyHide Is Nothing Then
            If Not onlyHide.Parent Is stgwb Then
                RaiseEvent OnAlert(True, "(pbSettings.HideSheets) " & onlyHide.Name & " is not a worksheet in workbook: " & stgwb.Name)
                Exit Function
            End If
            If WorksheetVisibleCount >= 2 Then
                onlyHide.visible = IIf(hideVeryHidden, xlSheetVeryHidden, xlSheetHidden)
            End If
        Else
            Dim tws As Worksheet
            
            For Each tws In stgwb.Worksheets
                If tws.visible <> xlSheetVisible And WorksheetVisibleCount > 1 Then
                    If CanHideSheet(tws) Then
                        tws.visible = IIf(hideVeryHidden, xlSheetVeryHidden, xlSheetHidden)
                    End If
                End If
            Next tws
        End If
    End Function
    
    Public Function CanHideSheet(wksht As Worksheet) As Boolean
        Dim doNotHide
        doNotHide = GetValue(STG_DO_NOT_HIDE)
        If isArray(doNotHide) Then
            Dim tsht
            For Each tsht In doNotHide
                If StringsMatchPriv(tsht, wksht.CodeName) Then
                    CanHideSheet = False
                    Exit Function
                End If
            Next tsht
        End If
        CanHideSheet = True
    End Function
    
    Public Function WorksheetVisibleCount(Optional countType As XlSheetVisibility = XlSheetVisibility.xlSheetVisible) As Long
            Dim resp As Long
            Dim tws As Worksheet
            For Each tws In stgwb.Worksheets
                If tws.visible = countType Then
                    resp = resp + 1
                End If
            Next tws
            WorksheetVisibleCount = resp
    End Function
    
'Public Function GrantRightsForMac(sPath)
'    #If Mac Then
'        Dim sScript
'        sScript = "set userName to short user name of (system info)" & vbNewLine & "return userName"
'
'    #End If
'End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Get current user name, based on OS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get UserNameOrLogin() As String
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        On Error Resume Next
        #If Mac Then
            Dim sScript
            sScript = "set userName to short user name of (system info)" & vbNewLine & "return userName"
            UserNameOrLogin = MacScript(sScript)
            '' If script failed, try LOGNAME
            If Err.number <> 0 Then
                RaiseEvent OnAlert(True, "pbSetting.UserNameOrLogin (Mac OS) - failed to get Mac OS UserName from Macscript (Error: " & Err.number & " - " & Err.Description)
                Err.Clear
                UserNameOrLogin = Environ("LOGNAME")
            End If
        #Else
           UserNameOrLogin = Environ("UserName")
        #End If
    End Property

    Private Function ResetDataTypeOption(ByVal currentOptions As pbSettingOption, ByVal dataType As pbSettingOption, ByVal isArray As Boolean) As pbSettingOption
        Dim resp As pbSettingOption
        resp = currentOptions
        If isArray Then
            resp = privEnumModify(resp, optValueAsArray, FlagEnumModifyPriv.feVerifyEnumExists)
        End If
        resp = privEnumModify(resp, optValueAsBoolean, FlagEnumModifyPriv.feVerifyEnumRemoved)
        resp = privEnumModify(resp, optValueAsDate, FlagEnumModifyPriv.feVerifyEnumRemoved)
        resp = privEnumModify(resp, optValueAsDateTime, FlagEnumModifyPriv.feVerifyEnumRemoved)
        resp = privEnumModify(resp, optValueAsDouble, FlagEnumModifyPriv.feVerifyEnumRemoved)
        resp = privEnumModify(resp, optValueAsGeneral, FlagEnumModifyPriv.feVerifyEnumRemoved)
        resp = privEnumModify(resp, optValueAsLong, FlagEnumModifyPriv.feVerifyEnumRemoved)
        resp = privEnumModify(resp, optValueAsText, FlagEnumModifyPriv.feVerifyEnumRemoved)
                        
         'If clng(dataType) >= clng(pbSettingOption.[_optValueMin]) and clng(dataType) <= clng(pbSettingOption[_optValueMax]) Then
        Select Case dataType
            Case pbSettingOption.optValueAsBoolean
                resp = privEnumModify(resp, optValueAsBoolean, FlagEnumModifyPriv.feVerifyEnumExists)
            Case pbSettingOption.optValueAsDate
                resp = privEnumModify(resp, optValueAsDate, FlagEnumModifyPriv.feVerifyEnumExists)
            Case pbSettingOption.optValueAsDateTime
                resp = privEnumModify(resp, optValueAsDateTime, FlagEnumModifyPriv.feVerifyEnumExists)
            Case pbSettingOption.optValueAsDouble
                resp = privEnumModify(resp, optValueAsDouble, FlagEnumModifyPriv.feVerifyEnumExists)
            Case pbSettingOption.optValueAsGeneral
                resp = privEnumModify(resp, optValueAsGeneral, FlagEnumModifyPriv.feVerifyEnumExists)
            Case pbSettingOption.optValueAsLong
                resp = privEnumModify(resp, optValueAsLong, FlagEnumModifyPriv.feVerifyEnumExists)
            Case pbSettingOption.optValueAsText
                resp = privEnumModify(resp, optValueAsText, FlagEnumModifyPriv.feVerifyEnumExists)
            Case Else
                resp = privEnumModify(resp, optValueAsGeneral, FlagEnumModifyPriv.feVerifyEnumExists)
        End Select
        ResetDataTypeOption = resp
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Force Setting Value to Go Through 'Val()' Function and set get values with or without
''  Decimals (teLong = no decimals, teDouble = decimals allowed)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ForceNumericFormat(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        Dim stgIDX As Long: stgIDX = Index(stgKey)
        Dim existOptions As pbSettingOption
        If stgIDX > 0 Then
            l_isDirty = True
            existOptions = SettingOptionsByIndex(stgIDX)
            existOptions = ResetDataTypeOption(existOptions, optValueAsDouble, IsArraySet(existOptions))
            l_lo.listRows(Index(stgKey)).Range(1, pbSTGColumns.scSettingOptions) = existOptions
        End If
    End Function
    Public Function ForceTextFormat(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        Dim stgIDX As Long: stgIDX = Index(stgKey)
        Dim existOptions As pbSettingOption
        If stgIDX > 0 Then
            l_isDirty = True
            existOptions = SettingOptionsByIndex(stgIDX)
            existOptions = ResetDataTypeOption(existOptions, optValueAsText, IsArraySet(existOptions))
            l_lo.listRows(Index(stgKey)).Range(1, pbSTGColumns.scSettingOptions) = existOptions
        End If
    End Function
    Public Function ForceDoubleFormat(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        Dim stgIDX As Long: stgIDX = Index(stgKey)
        Dim existOptions As pbSettingOption
        If stgIDX > 0 Then
            l_isDirty = True
            existOptions = SettingOptionsByIndex(stgIDX)
            existOptions = ResetDataTypeOption(existOptions, optValueAsDouble, IsArraySet(existOptions))
            l_lo.listRows(Index(stgKey)).Range(1, pbSTGColumns.scSettingOptions) = existOptions
        End If

    End Function
    Public Function ForceLongFormat(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        Dim stgIDX As Long: stgIDX = Index(stgKey)
        Dim existOptions As pbSettingOption
        If stgIDX > 0 Then
            l_isDirty = True
            existOptions = SettingOptionsByIndex(stgIDX)
            existOptions = ResetDataTypeOption(existOptions, optValueAsLong, IsArraySet(existOptions))
            l_lo.listRows(Index(stgKey)).Range(1, pbSTGColumns.scSettingOptions) = existOptions
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Force Setting Value to Go Through 'CDate()' Function and Allow Time
''  Note: A setting value that cannot be converted to a date will cause
''      an error
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ForceDateTimeFormat(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        Dim stgIDX As Long: stgIDX = Index(stgKey)
        Dim existOptions As pbSettingOption
        If stgIDX > 0 Then
            l_isDirty = True
            existOptions = SettingOptionsByIndex(stgIDX)
            existOptions = ResetDataTypeOption(existOptions, optValueAsDateTime, IsArraySet(existOptions))
            l_lo.listRows(Index(stgKey)).Range(1, pbSTGColumns.scSettingOptions) = existOptions
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Force Setting Value to Go Through 'CDate()' Function and Only Allow Date (No Time)
''  Note: A setting value that cannot be converted to a date will cause
''      an error
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ForceDateFormat(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        Dim stgIDX As Long: stgIDX = Index(stgKey)
        Dim existOptions As pbSettingOption
        If stgIDX > 0 Then
            l_isDirty = True
            existOptions = SettingOptionsByIndex(stgIDX)
            existOptions = ResetDataTypeOption(existOptions, optValueAsDate, IsArraySet(existOptions))
            l_lo.listRows(Index(stgKey)).Range(1, pbSTGColumns.scSettingOptions) = existOptions
        End If

    End Function
    
    
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Force Setting Value to Go Through 'CBool()' Function
''  Note: A setting value that cannot be converted to a Boolean will cause
''      an error
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ForceBooleanFormat(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        Dim stgIDX As Long: stgIDX = Index(stgKey)
        Dim existOptions As pbSettingOption
        If stgIDX > 0 Then
            l_isDirty = True
            existOptions = SettingOptionsByIndex(stgIDX)
            existOptions = ResetDataTypeOption(existOptions, optValueAsBoolean, IsArraySet(existOptions))
            l_lo.listRows(Index(stgKey)).Range(1, pbSTGColumns.scSettingOptions) = existOptions
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Force Setting Value to return as Array
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ForceArrayFormat(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault, Optional dataType As pbSettingOption = pbSettingOption.optValueAsGeneral)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        Dim stgIDX As Long: stgIDX = Index(stgKey)
        Dim existOptions As pbSettingOption
        If stgIDX > 0 Then
            l_isDirty = True
            existOptions = SettingOptionsByIndex(stgIDX)
            existOptions = ResetDataTypeOption(existOptions, dataType, True)
            l_lo.listRows(Index(stgKey)).Range(1, pbSTGColumns.scSettingOptions) = existOptions
        End If
        
    End Function
    
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Force Setting Value to be returned as it was created (This is the Default)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ForceGeneralFormat(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        Dim stgIDX As Long: stgIDX = Index(stgKey)
        Dim existOptions As pbSettingOption
        If stgIDX > 0 Then
            l_isDirty = True
            existOptions = SettingOptionsByIndex(stgIDX)
            existOptions = ResetDataTypeOption(existOptions, optValueAsGeneral, IsArraySet(existOptions))
            l_lo.listRows(Index(stgKey)).Range(1, pbSTGColumns.scSettingOptions) = existOptions
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Delete Setting with key [stgKey], if it exists
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function Delete(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        If SettingIsRequired(stgKey) = True Then
            RaiseEvent OnAlert(True, "pbSetting: '" & stgKey & "' could not be deleted because it contains option: optKeyRequired")
        Else
            If Index(stgKey) > 0 Then
                Dim cancelOperation As Boolean
                RaiseEvent OnDeleteSetting(stgKey, cancelOperation)
                DoEvents
                If cancelOperation = True Then
                    'cancel
                Else
                    Dim evts As Boolean: evts = Me.Events: Me.Events = False
                    ''  Delete Setting
                    l_lo.listRows(Index(stgKey)).Delete
                    ''  Restore App Events to prior setting
                    Me.Events = evts
                    l_isDirty = True
                End If
            End If
        End If
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns Index of List for [settingKey]
''  Returns 0 if Missing
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function Index(ByVal stgKey, Optional options As pbSettingOption = pbSettingOption.optDefault) As Long
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        stgKey = CheckKey(stgKey, options)
        If ValidConfig And Count > 0 Then
            Dim tmpIdx As Variant
            On Error Resume Next
            tmpIdx = keyIndex(stgKey)
            If Err.number = 0 Then
                Index = CLng(tmpIdx)
            Else
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End Function
    Public Function keyIndex(ByVal checkedKey) As Long
        If Count > 0 Then
            On Error Resume Next
            Dim tmpIdx As Long
            tmpIdx = WorksheetFunction.XMatch(checkedKey, l_lo.ListColumns(pbSTGColumns.scSettingKey).DataBodyRange, XMatch.exactMatch, XSearch.xsBinaryAsc)
            If Err.number = 0 Then
                keyIndex = CLng(tmpIdx)
            Else
                Err.Clear
            End If
        End If
    End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Exposed for Convenience - not needed for using pbSettings
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''  SETTINGS WORKSHEET REFERENCE
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get pbSettingsSheet() As Worksheet
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Set pbSettingsSheet = l_ws
    End Property
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''  SETTINGS LIST OBJECT REFERENCE
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get pbSettingsListObj() As ListObject
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Set pbSettingsListObj = l_lo
    End Property
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''  Earliest Modified Date for any Setting
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get ModifiedEarliestDate() As Variant
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        If ValidConfig Then
            Dim modified As Variant
            modified = CDate(WorksheetFunction.Min(l_lo.ListColumns(pbSTGColumns.scUpdated).DataBodyRange))
            ModifiedEarliestDate = modified
        End If
    End Property
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    ''  Earliest Modified Date for any Setting
    ' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get ModifiedLatestDate() As Variant
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        If ValidConfig Then
            Dim modified As Variant
            modified = CDate(WorksheetFunction.Max(l_lo.ListColumns(pbSTGColumns.scUpdated).DataBodyRange))
            ModifiedLatestDate = modified
        End If
    End Property
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns the Count of Unique Settings
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get Count() As Long
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
        Count = l_lo.listRows.Count
    End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Exports Settings that match [wildcardsearch] if provided, otherwise all
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Function ExportSettings(Optional wildcardSearch As Variant)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        If Count > 0 Then
            Dim exportArr() As Variant, iRow, iCol, newWB As Workbook, newWS As Worksheet
            Dim tmpIndexes As New Collection
            exportArr = l_lo.Range.value
            For iRow = LBound(exportArr) To UBound(exportArr)
                If iRow = 1 Then
                    'Add Column Headers Index
                    tmpIndexes.add iRow
                Else
                    If IsMissing(wildcardSearch) Then
                        tmpIndexes.add iRow
                    Else
                        For iCol = LBound(exportArr, 2) To UBound(exportArr, 2)
                            If StringsMatchPriv(exportArr(iRow, iCol), wildcardSearch, smContains) Then
                                tmpIndexes.add iRow
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
            Set newWB = Workbooks.add
            Set newWS = newWB.Worksheets(1)
            With newWS
                .Name = "SettingsExport"
                .Range("A5").Resize(RowSize:=tmpIndexes.Count, ColumnSize:=l_lo.ListColumns.Count).value = newArr
                .UsedRange.EntireColumn.AutoFit
                Dim checkCol As Long
                For checkCol = 1 To .UsedRange.Columns.Count
                    If .Cells(1, checkCol).EntireColumn.ColumnWidth > 100 Then
                        .Cells(1, checkCol).EntireColumn.ColumnWidth = 100
                        .Cells(1, checkCol).EntireColumn.WrapText = True
                    End If
                Next checkCol
                .UsedRange.VerticalAlignment = xlVAlignCenter
                .ListObjects.add SourceType:=xlSrcRange, Source:=.UsedRange, XLListObjectHasHeaders:=xlYes
                .Activate
                .Parent.Windows(1).DisplayGridlines = False
                
                .Range("A1").value = "Settings Exported From: " & stgwb.FullName
                .Range("A2").value = "Exported on: " & CStr(Now())
                If IsMissing(wildcardSearch) Then
                    wildcardSearch = "[None]"
                Else
                    wildcardSearch = """" & wildcardSearch & """"
                End If
                .Range("A3").value = "Search Expression: " & wildcardSearch
            End With
            newWB.Activate
            Set newWS = Nothing
            Set newWB = Nothing
            Beep
        End If
    
    End Function

    Private Property Get SettingsWorkbook() As Workbook
        Set SettingsWorkbook = stgwb
    End Property
    Private Property Set SettingsWorkbook(wb As Workbook)
        If stgwb Is Nothing Then
            Set stgwb = wb
        End If
    End Property

    Public Function DEV_TEST()
        If l_InitAttempted Then
            MsgBox "In order to test pbSettings, 'DEV_TEST must be the first method called after instantiation, and pbSettings will create a new workbookk with settings configured to test."
            Exit Function
        End If
        Dim wkbk As Workbook
        Set wkbk = Application.Workbooks.add
        DoEvents
        If Not wkbk Is Nothing Then
            wkbk.Worksheets(1).Name = "pbSettingsTest"
        End If
        wkbk.SaveAs fileName:=ThisWorkbook.path & Application.PathSeparator & "DEV_TEST_pbSettings.xlsm", FileFormat:=XlFileFormat.xlOpenXMLWorkbookMacroEnabled, addToMRU:=False
'        wkbk.Worksheets(1).Name = "pbSettingsTest_" & UCase(Format(Now(), "_yyyyMMMdd_hhnnss"))
        wkbk.VBProject.VBComponents.add vbext_ct_StdModule
        wkbk.Save
        Set SettingsWorkbook = wkbk
        StartSettings
    End Function

    Private Function StartSettings()
        On Error Resume Next
        If SettingsWorkbook Is Nothing Then
            Set SettingsWorkbook = ThisWorkbook
        End If
        l_InitAttempted = True
        l_validCfg = VerifySettingsSheet
        If Not l_validCfg Then
            Dim createMsg
             createMsg = "The Worksheet and Table needed for pbSettings was not found. Would you like to create that automatically now?"
            If MsgBox(createMsg, vbYesNo + vbDefaultButton1 + vbQuestion, "pbSettings Configuration") = vbYes Then
                l_validCfg = CreateSettingsSheet
            End If
        End If
        If l_validCfg Then
            CheckSettingsHealth
            CheckDefaultCommonSettings
            ConfigureWorkbook
        End If
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns True If SETTINGS_SHEET_CODENAME worksheet exists
''  and the worksheet contains a valid listobject for managing settings
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  THIS IS THE ONLY METHOD THAT CAN BE USED IF THE CONFIGURATION IS NOT VALID
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Public Property Get ValidConfig() As Boolean
        If Not l_InitAttempted Then
            StartSettings
        End If
        
        If l_validCfg = False And l_invalidMessageShown = False Then
            l_invalidMessageShown = True
            Beep
            MsgBox "There appears to be a problem with configuring pbSettings. If you have manually modified the 'pbSettings' worksheet, you may wish to delete that sheet entirely, and try again -- 'pbSettings' " _
            & " will rebuild the worksheet and listobject structure if needed.", vbOKOnly + vbExclamation, "OOPS"
        End If
        ValidConfig = l_validCfg
    End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  If found, returns the worksheet identified by the constant:
''      SETTINGS_SHEET_CODENAME
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function VerifySettingsSheet() As Boolean
        Dim tmpWS As Worksheet
        Dim tmpLO As ListObject
        Dim lstCol As listColumn
        '' First Try to Find by 'CodeName'
        Dim findCodeName: findCodeName = SETTINGS_SHEET_CODENAME
        For Each tmpWS In stgwb.Worksheets
            If StringsMatchPriv(tmpWS.CodeName, findCodeName) Then
                Set l_ws = tmpWS
                Exit For
            End If
        Next
        If l_ws Is Nothing Then
            For Each tmpWS In stgwb.Worksheets
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  If found, returns the validated pbSettings list object from [wksht]
''  The list object name must match the constant SETTINGS_LO exactly
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Note: This method is used to get the reference to the settings list object
''      in the current workbook, and it is also used to validate the pbSettings
''      ListObject structure in the template that is used for FIRST RUN
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function GetSettingsLO(wksht As Worksheet) As ListObject
            Dim tmpLO As ListObject
            Dim validLO As ListObject, lstCol As listColumn
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Returns the explicit name of ListColumns as a Base-1 1D Array,
''       that are expected to be in the settings ListObject
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Property Get ExpectedColsArr() As Variant
        Dim tmpArr As Variant: tmpArr = Split(SETTINGS_COL_NAMES, delimiter:="|", compare:=vbTextCompare)
        Dim i As Long, colName As Variant, tmpArrSize As Long
        tmpArrSize = ArrSize(tmpArr)
        If tmpArrSize <> pbSTGColumns.[_TOTALCOLS] Then
                Err.Raise 13, Source:=Me, Description:="pbSettings.ExpectedColsArr size must match pbStgColumns.[_TOTALCOLS]"
        End If
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

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Called when class is initialized, and any time a new settings Key
'   has been added
'   Will check the following:
'       1 - Ensure there are no duplicate settings keys
'             (Will remove earliest modified row if duplicate keys are found)
'       2 - Ensure ListObject is sorted by the Key Column
'             (Use: pbSTGColumns.scSettingKey to refer to Key Column Index)
'       3 - Remove Any Invalid 'SettingType'
'             (Valid values are: String, Boolean, Number, DateTime)
'             (This helps ensure retrieved settings formats will populate
'              into Variant or explicitly declared types correctly --
'              e.g. if you added '45262.9420949074' as the setting value
'              for a setting declared as 'DateTime', the setting would return
'              '12/2/23 10:36:37 PM' as a Variant of type Date)
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function CheckSettingsHealth()
        Dim evts As Boolean: evts = Me.Events: Me.Events = False
        If l_lo.ShowTotals Then l_lo.ShowTotals = False
        CheckSort
        CheckRequiredSettings
        If stdAutoHide And l_ws.visible = xlSheetVisible And OtherSheetsVisible Then
            l_ws.visible = xlSheetVeryHidden
        End If
        CheckDuplicates
        SizeColumns
        Me.Events = evts
    End Function
    Private Function DeleteColumnsIfExist(ParamArray colNames() As Variant)
        On Error Resume Next
        Dim colName
        Dim evts As Boolean: evts = Me.Events: Me.Events = False
        
        For Each colName In colNames
            Dim colIdx As Long
            For colIdx = l_lo.ListColumns.Count To 1 Step -1
                If StringsMatchPriv(colName, l_lo.ListColumns(colIdx).Name) Then
                    l_lo.ListColumns(colIdx).Delete
                End If
            Next colIdx
        Next colName
        Me.Events = evts
    End Function


' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Used to verify there is another worksheet that's visible, so that
''   we can autohide the pbSettings sheet if it is visible and the
''  setting "STG_AUTOHIDE" is set to TRUE
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function CheckDuplicates()
        On Error Resume Next
        Dim uniqueKeys As Variant
        Dim loName, keyColName, expr
        loName = l_lo.Name
        keyColName = l_lo.ListColumns(pbSTGColumns.scSettingKey).Name
        expr = "COUNTA(UNIQUE(" & loName & "[" & keyColName & "]))"
        uniqueKeys = Evaluate(expr)
        If VAL(uniqueKeys) < Count Then
            l_isDirty = True
            RemoveDuplicates
        End If
        If Err.number <> 0 Then
            Err.Clear
            '' If there was an error checking, then look through each row
            RemoveDuplicates
        End If
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Removes duplicates
''  Will keep the setting with the latest updated date if duplicates are found
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function RemoveDuplicates()
        Dim deleteItems As New Collection
        Dim srcArr() As Variant
        srcArr = l_lo.DataBodyRange.value
        Dim IDX As Long, startIdx As Long
        Dim keepIdx, keepKey, keepDt
        For IDX = LBound(srcArr) To UBound(srcArr)
            If CollectionItemExists(deleteItems, IDX) = False Then
                keepIdx = IDX
                keepKey = srcArr(IDX, pbSTGColumns.scSettingKey)
                keepDt = CDate(srcArr(IDX, pbSTGColumns.scUpdated))
                For startIdx = IDX + 1 To UBound(srcArr)
                    If CollectionItemExists(deleteItems, startIdx) = False Then
                        If StringsMatchPriv(keepKey, srcArr(startIdx, pbSTGColumns.scSettingKey)) Then
                            If CDate(srcArr(startIdx, pbSTGColumns.scUpdated)) >= keepDt Then
                                deleteItems.add keepIdx
                                keepIdx = startIdx
                                keepDt = CDate(srcArr(startIdx, pbSTGColumns.scUpdated))
                            Else
                                deleteItems.add startIdx
                            End If
                        End If
                    End If
                Next startIdx
            End If
        Next
        Dim deleteItem
        ''  Disable App Events
        Dim evts As Boolean: evts = Me.Events: Me.Events = False
        For Each deleteItem In deleteItems
            l_lo.listRows(deleteItem).Range(1, pbSTGColumns.scSettingKey) = "PB_SETTING_DUPLICATE_DELETE"
        Next
        Sort
        Dim delRng As Range
        Set delRng = l_lo.ListColumns(pbSTGColumns.scSettingKey).Range.Find("PB_SETTING_DUPLICATE_DELETE")
        If Not delRng Is Nothing Then
            Do While Not delRng Is Nothing
                delRng.Resize(ColumnSize:=l_lo.ListColumns.Count).Delete xlShiftUp
                Set delRng = l_lo.ListColumns(pbSTGColumns.scSettingKey).Range.Find("PB_SETTING_DUPLICATE_DELETE")
            Loop
        End If
        ''  Restore App Events to prior setting
        Me.Events = evts
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Used to verify there is another worksheet that's visible, so that
''   we can autohide the pbSettings sheet if it is visible and the
''  setting "STG_AUTOHIDE" is set to TRUE
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function OtherSheetsVisible() As Boolean
        Dim vWS As Worksheet
        For Each vWS In stgwb.Worksheets
            If vWS.visible = xlSheetVisible And Not vWS Is l_ws Then
                OtherSheetsVisible = True
                Exit For
            End If
        Next
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Check Required Settings is for settings that need to exist that
''   are used to manage pbSettings.
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function CheckRequiredSettings()
        '' If 'STG_AUTOHIDE' key does not exist, create it with default: True
        If Exists(PBSTG_HIDE) = False Then
            stdAutoHide = True
        End If
        If Exists(PBSTG_WS_CODENAME) = False Then
            Setting(PBSTG_WS_CODENAME) = l_ws.CodeName
        End If
        
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Check is settings ListObject is ordered -- needed to support the
'       binary searching of keys
'   This method is called when class is initialzed and whever a settings
'       KEY is added
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function CheckSort()
        If Not Sorted Then Sort
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Ensure Settings are sorted by Keys, so that binary search can be
'   used to most quickly find the setting
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function Sort()
        With l_lo.Sort
            ''  Disable App Events
            Dim evts As Boolean: evts = Me.Events: Me.Events = False
            .SortFields.Clear
            .SortFields.add l_lo.ListColumns(pbSTGColumns.scSettingKey).DataBodyRange, SortOn:=xlSortOnValues, Order:=XlSortOrder.xlAscending
            .Apply
            ''  Restore App Events to prior setting
            Me.Events = evts
        End With
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Build and evaluate Array Formulat to determine if sorted by Key
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function Sorted() As Boolean
        If Not ValidConfig Then Exit Function
        With l_lo
            If .listRows.Count <= 1 Then
                Sorted = True
            Else
                Dim expr As String, rng1 As Range, rng2 As Range
                Set rng1 = .ListColumns(pbSTGColumns.scSettingKey).DataBodyRange.Resize(.listRows.Count - 1)
                Set rng2 = rng1.Offset(rowOffset:=1)
                expr = "AND(" & "'[" & stgwb.Name & "]" & rng1.Worksheet.Name & "'!" & rng1.Address & "<='[" & stgwb.Name & "]" & rng2.Worksheet.Name & "'!" & rng2.Address & ")"
                Sorted = Evaluate(expr)
            End If
        End With
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Returns the count of  elements from the first dimension of an array
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function ArrSize(inputArr As Variant) As Long
        If privEnumCompare(VarType(inputArr), VbVarType.vbArray) Then
            ArrSize = UBound(inputArr) - LBound(inputArr) + 1
        End If
    End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Private (class) version of the dtPart Function available in
'   some of the common modules at:
'     https://github.com/lopperman/just-VBA
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

    Private Function dtPartPriv(thePart As DateDiffType_priv, dt1 As Variant, _
        Optional ByVal firstDayOfWeek As VbDayOfWeek = vbSunday, _
        Optional ByVal firstWeekOfYear As VbFirstWeekOfYear = VbFirstWeekOfYear.vbFirstJan1) As Variant
        Select Case thePart
            Case DateDiffType_priv.dtpDate_NoTime
                dtPartPriv = DateSerial(dtPartPriv(dtpYear, dt1), dtPartPriv(dtpMonth, dt1), dtPartPriv(dtpDay, dt1))
            Case DateDiffType_priv.dtpDay
                dtPartPriv = DatePart("d", dt1, firstDayOfWeek, firstWeekOfYear)
            Case DateDiffType_priv.dtpDayOfYear
                dtPartPriv = DatePart("y", dt1, firstDayOfWeek, firstWeekOfYear)
            Case DateDiffType_priv.dtpHour
                dtPartPriv = DatePart("h", dt1, firstDayOfWeek, firstWeekOfYear)
            Case DateDiffType_priv.dtpMinute
                dtPartPriv = DatePart("n", dt1, firstDayOfWeek, firstWeekOfYear)
            Case DateDiffType_priv.dtpMonth
                dtPartPriv = DatePart("m", dt1, firstDayOfWeek, firstWeekOfYear)
            Case DateDiffType_priv.dtpQuarter
                dtPartPriv = DatePart("q", dt1, firstDayOfWeek, firstWeekOfYear)
            Case DateDiffType_priv.dtpSecond
                dtPartPriv = DatePart("s", dt1, firstDayOfWeek, firstWeekOfYear)
            Case DateDiffType_priv.dtpWeek
                dtPartPriv = DatePart("ww", dt1, firstDayOfWeek, firstWeekOfYear)
            Case DateDiffType_priv.dtpWeekday
                dtPartPriv = DatePart("w", dt1, firstDayOfWeek, firstWeekOfYear)
            Case DateDiffType_priv.dtpYear
                dtPartPriv = DatePart("yyyy", dt1, firstDayOfWeek, firstWeekOfYear)
        End Select
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
    
    Private Function ConvertItemFormat(ByRef item, options As pbSettingOption)
        On Error Resume Next
        
        ''  VALUE AS BOOLEAN
        If privEnumCompare(options, pbSettingOption.optValueAsBoolean) Then
            item = CBool(item)
            If Err.number <> 0 Then
                Err.Clear
                item = False
            End If
        
        ''  VALUE AS DOUBLE
        ElseIf privEnumCompare(options, pbSettingOption.optValueAsDouble) Then
            item = VAL(item)
            If Err.number <> 0 Then
                Err.Clear
                item = 0
            End If
            
        ''  VALUE AS TEXT
        ''  IF YOU ARE GOING TO BE 'SENDING' NON-TEXT VALUES TO SETTINGS
        ''  THAT NEED THE VALUE RETURNED AS A STRING, MAKE SURE
        ''  options includes: pbSettingOption.optValueAsText
        ''  Example would be a product code "001"
        ElseIf privEnumCompare(options, pbSettingOption.optValueAsText) Then
            item = CStr(item)
        ''  VALUE AS LONG
        ''  NO ROUNDING WILL BE PERFORMED
        ''  IF A SETTING TYPED AS 'optValueAsLong' is set to value "1", or "1.0001" or "1.9999999"
        ''      the setting value saved will be: 1
        ''  IF ROUNDING IS NEEDED, MANAGE THAT OUTSIDE THE SETTINGS CLASS
        ElseIf privEnumCompare(options, pbSettingOption.optValueAsLong) Then
            If CDbl(item) <> CLng(item) Then
                Dim lngString As String: lngString = CStr(item)
                item = CLng(Mid(lngString, 1, InStr(1, lngString, ".", vbTextCompare) - 1))
            Else
                item = CLng(item)
            End If
            If Err.number <> 0 Then
                Err.Clear
                item = CLng(0)
            End If
        
        ''  VALUE AS DATE (NO TIME)
        ElseIf privEnumCompare(options, pbSettingOption.optValueAsDate) Then
            item = dtPartPriv(dtpDate_NoTime, CDate(item))
        
        ''  VALUE AS DATETIME
        ElseIf privEnumCompare(options, pbSettingOption.optValueAsDateTime) Then
            item = CDate(item)
        End If
        
        ConvertItemFormat = item
    End Function
    
    Public Function FormatValue(ByVal item, ByVal options As pbSettingOption)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
        Dim isArray As Boolean: isArray = privEnumCompare(options, pbSettingOption.optValueAsArray)
        
        If Not isArray Then
            FormatValue = ConvertItemFormat(item, options)
        Else
            Dim dlm As String: dlm = stdArrayDelimiter
            Dim tempArray As Variant
            Dim retArray()
            tempArray = Split(item, dlm, , vbTextCompare)
            
            
            '' If data types
            If Not privEnumCompare(options, optValueAsGeneral) Then
                Dim tmpItem As Variant, arrLBound As Long, arrIdx As Long
                arrLBound = LBound(tempArray)
                arrIdx = arrLBound
                ReDim retArray(arrLBound To UBound(tempArray))
                
                
                For Each tmpItem In tempArray
                    retArray(arrIdx) = ConvertItemFormat(tmpItem, options)
                    arrIdx = arrIdx + 1
                Next tmpItem
            End If
            FormatValue = retArray
        End If
    End Function
    
    Private Function ToArray(ByVal item, Optional SettingType As SettingTypeEnum = SettingTypeEnum.teGeneral)
        Dim dlm As String: dlm = stdArrayDelimiter
        Dim tempArray As Variant
        tempArray = Split(item, dlm, , vbTextCompare)
        ToArray = Split(item, dlm, , vbTextCompare)
    End Function
    
    Private Function privConcatWithDelim(ByVal delimeter As String, ParamArray items() As Variant) As String
        privConcatWithDelim = Join(items, delimeter)
    End Function
    

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   'FIRST RUN' METHODS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   Once the correctly structured pbSettings Worksheet and ListObject
'   Have been added, these methods are not longer used
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   ONLY USED IN 'FIRST RUN'
'   Create settings sheet
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function CreateSettingsSheet() As Boolean
        On Error Resume Next
        Dim tmplSheet As Worksheet
        Dim intr As Boolean, curs As Variant
        Dim evts As Boolean: evts = Me.Events: Me.Events = False
        intr = Application.Interactive
        curs = Application.Cursor
        Application.StatusBar = "Configuring new Settings Area ...": DoEvents
        Application.Interactive = False
        Application.Cursor = XlMousePointer.xlWait
        
        Set tmplSheet = stgwb.Worksheets.add(After:=stgwb.Worksheets(stgwb.Worksheets.Count))
        
        
        tmplSheet.Name = SETTINGS_SHEET_WEAKNAME
        tmplSheet.Activate
        stgwb.Windows(1).DisplayGridlines = False
        Dim infoRng As Range
        Set infoRng = tmplSheet.Range("A3:D3")
        infoRng.Merge Across:=True
        infoRng(1, 1).value = "MANUALLY EDITING THESE SETTING COULD CAUSE PERMANENT LOSS OF DATA OR COULD CORRUPT THIS WORKBOOK"
        infoRng.Interior.color = 13431551
        infoRng.Font.color = 192
        tmplSheet.Range("A4").value = "Created by pbSettings - Click to view in just-VBA on GitHub"
        tmplSheet.Range("A3:D3").EntireRow.RowHeight = tmplSheet.Range("A4:D4").EntireRow.RowHeight * 2
        tmplSheet.Range("A3:D3").Merge
        
        tmplSheet.Range("A4:D4").Merge
        With tmplSheet
            .Hyperlinks.add Anchor:=tmplSheet.Range("A4:D4"), _
            Address:="https://github.com/lopperman/just-VBA/blob/main/pbSettings/pbSettings.md", _
            ScreenTip:="View pbSettings on just-VBA Github repository"
            .Range("A3:D4").VerticalAlignment = xlVAlignCenter
            .Range("A3:D4").HorizontalAlignment = xlHAlignCenter
            .Range("A3:D4").WrapText = True
        End With
        
        stgwb.Save
        DoEvents
        ReCodeName stgwb, tmplSheet.CodeName, SETTINGS_SHEET_CODENAME
        
        Dim tRng As Range
        Set tRng = tmplSheet.Range("A6").Resize(ColumnSize:=ArrSize(ExpectedColsArr))
        tRng.value = ExpectedColsArr
        Set tRng = tRng.Resize(RowSize:=2)
        Dim lstObj As ListObject
        Set lstObj = tmplSheet.ListObjects.add(SourceType:=xlSrcRange, Source:=tRng, XLListObjectHasHeaders:=xlYes, tableStyleName:="TableStyleMedium2")
        lstObj.Name = SETTINGS_LO
        With lstObj.listRows(1)
            .Range(1, pbSTGColumns.scSettingKey) = PBSTG_HIDE
            .Range(1, pbSTGColumns.scSettingValue) = True
            .Range(1, pbSTGColumns.scSettingOptions) = pbSettingOption.optValueAsBoolean + pbSettingOption.optKeyRequired + pbSettingOption.optNonMigratableKey
            .Range(1, pbSTGColumns.scUpdated) = Now()
        End With
        lstObj.HeaderRowRange.HorizontalAlignment = xlHAlignCenter
        lstObj.HeaderRowRange.VerticalAlignment = xlVAlignCenter
        lstObj.HeaderRowRange.Font.Bold = True
        lstObj.ListColumns(pbSTGColumns.scSettingKey).DataBodyRange.Font.Bold = True
        lstObj.ListColumns(pbSTGColumns.scSettingKey).DataBodyRange.Font.color = 16724484
        Set tmplSheet = Nothing
        If Err.number = 0 Then
            Set l_ws = lstObj.Range.Worksheet
            Set l_lo = lstObj
        End If
        Me.Events = evts
        Application.Interactive = intr
        Application.Cursor = curs
        Application.StatusBar = False
        If Err.number <> 0 Then
            Err.Clear
            CreateSettingsSheet = False
        Else
            CreateSettingsSheet = True
        End If
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''  Utility Method to change the codeName of a Worksheet
''  From my article on Reddit:
''  https://www.reddit.com/r/vba/comments/18wo10x/how_to_change_the_codename_of_a_worksheet_using/
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    Private Function ReCodeName(wkbk As Workbook, currentCodeName, newCodeName) As Boolean
    On Error Resume Next
        If wkbk.HasVBProject = True Then
            Dim vbComp As VBComponent
            Set vbComp = wkbk.VBProject.VBComponents(currentCodeName)
            If Not vbComp Is Nothing Then
                vbComp.Properties("_CodeName") = CStr(newCodeName)
                If StringsMatchPriv(vbComp.Properties("_CodeName"), newCodeName) Then
                    Beep
                    ReCodeName = True
                End If
            End If
        End If
        If Err.number <> 0 Then
            RaiseEvent OnAlert(True, "Unable to execute 'pbSettings.ReCodeName', Error: " & Err.number & " - " & Err.Description)
            Err.Clear
        End If
        
    End Function
    
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   END OF 'FIRST RUN' METHODS
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  ENSURE DEFAULT COMMON SETTINGS EXIST
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Function CheckDefaultCommonSettings()
    On Error Resume Next
    Dim tmpOptions As pbSettingOption
    
    If Not Exists(PBSTG_VERSION) Then Setting(PBSTG_VERSION) = CDbl(0.1)
    If StringsMatchPriv(Setting(PBSTG_WS_CODENAME), l_ws.CodeName) = False Then
        Setting(PBSTG_WS_CODENAME) = l_ws.CodeName
    End If
    If Exists(PBSTG_ARRAY_DELIMITER) = False Then
        Setting(PBSTG_ARRAY_DELIMITER) = "|"
    ElseIf Len(Setting(PBSTG_ARRAY_DELIMITER)) = 0 Then
        Setting(PBSTG_ARRAY_DELIMITER) = "|"
    End If
    '' DEV SHOW LOG MSGS IN IMMEDIATE WINDOW
    If Exists(PBSTG_DEV_LOG_TO_IMMEDIATE) = False Then Setting(PBSTG_DEV_LOG_TO_IMMEDIATE) = False
    tmpOptions = SettingOptions(PBSTG_DEV_LOG_TO_IMMEDIATE)
    If Not SettingTypeByOptions(tmpOptions) = teBoolean Then
        ForceBooleanFormat PBSTG_DEV_LOG_TO_IMMEDIATE, tmpOptions
    End If
        
    '' PBSTG_RETURN_1BASED_ARRAYS
    '' WHEN TRUE (DEFAULT), AN ARRAY SETTING FIRST ITEM IS 'Array(1)' instead of 'Array(0)'
    If Exists(PBSTG_RETURN_1BASED_ARRAYS) = False Then Setting(PBSTG_RETURN_1BASED_ARRAYS) = True
    tmpOptions = SettingOptions(PBSTG_RETURN_1BASED_ARRAYS)
    If Not SettingTypeByOptions(tmpOptions) = teBoolean Then
        ForceBooleanFormat PBSTG_RETURN_1BASED_ARRAYS, tmpOptions
    End If
        
    ''DEV_NAMES
    If Exists(PBSTG_DEV_NAMES) = False Then
        ArraySetting(PBSTG_DEV_NAMES) = UserNameOrLogin
    End If
    tmpOptions = SettingOptions(PBSTG_DEV_NAMES)
    If Not SettingTypeByOptions(tmpOptions) = teText Then
        ForceTextFormat PBSTG_DEV_NAMES, tmpOptions
    End If
    
    'ADMIN_NAMES
    If Exists(PBSTG_ADMIN_NAMES) = False Then
        ArraySetting(PBSTG_ADMIN_NAMES) = UserNameOrLogin
    End If
    tmpOptions = SettingOptions(PBSTG_ADMIN_NAMES)
    If Not SettingTypeByOptions(tmpOptions) = teText Then
        ForceTextFormat PBSTG_ADMIN_NAMES, tmpOptions
    End If
    
    
    '' PBSETTING SHEET - MAX COLUMN WIDTH
    If Exists(PBSTG_MAX_SETTINGS_COL_WIDTH) = False Then Setting(PBSTG_MAX_SETTINGS_COL_WIDTH) = SETTINGS_MAX_COLUMN_WIDTH
    tmpOptions = SettingOptions(PBSTG_MAX_SETTINGS_COL_WIDTH)
    If Not SettingTypeByOptions(tmpOptions) = teDouble Then
        ForceDoubleFormat PBSTG_MAX_SETTINGS_COL_WIDTH, tmpOptions
    End If
    
    '' DEV_MODE
    If Exists(PBSTG_DEV_MODE) = False Then Setting(PBSTG_DEV_MODE) = False
    tmpOptions = SettingOptions(PBSTG_DEV_MODE)
    If Not SettingTypeByOptions(tmpOptions) = teBoolean Then
        ForceBooleanFormat PBSTG_DEV_MODE, tmpOptions
    End If
    
    '' DISABLE AUTO SAVE (DEFAULT TRUE)
    If Exists(PBSTG_DISABLE_AUTO_SAVE) = False Then Setting(PBSTG_DISABLE_AUTO_SAVE) = True
    tmpOptions = SettingOptions(PBSTG_DISABLE_AUTO_SAVE)
    If Not SettingTypeByOptions(tmpOptions) = teBoolean Then
        ForceBooleanFormat PBSTG_DISABLE_AUTO_SAVE, tmpOptions
    End If
    
    '' ENABLE WORKBOOK CONNECTIONS
    If Exists(PBSTG_ENABLE_CONNECTIONS) = False Then Setting(PBSTG_ENABLE_CONNECTIONS) = True
    tmpOptions = SettingOptions(PBSTG_ENABLE_CONNECTIONS)
    If Not SettingTypeByOptions(tmpOptions) = teBoolean Then
        ForceBooleanFormat PBSTG_ENABLE_CONNECTIONS, tmpOptions
    End If
    
    '' SAVE LINKED VALUES
    If Exists(PBSTG_SAVE_LINK_VALUES) = False Then Setting(PBSTG_SAVE_LINK_VALUES) = True
    tmpOptions = SettingOptions(PBSTG_SAVE_LINK_VALUES)
    If Not SettingTypeByOptions(tmpOptions) = teBoolean Then
        ForceBooleanFormat PBSTG_SAVE_LINK_VALUES, tmpOptions
    End If
    
    '' DEV: SHOW DEV TOOLS
    If Exists(PBSTG_DEV_SHOW_TOOLS) = False Then Setting(PBSTG_DEV_SHOW_TOOLS) = False
    tmpOptions = SettingOptions(PBSTG_DEV_SHOW_TOOLS)
    If Not SettingTypeByOptions(tmpOptions) = teBoolean Then
        ForceBooleanFormat PBSTG_DEV_SHOW_TOOLS, tmpOptions
    End If
    
    ''  DO NOT HIDE WORKSHEET CODENAMES
    If Exists(PBSTG_DO_NOT_HIDE) = False Then Setting(PBSTG_DO_NOT_HIDE) = vbNullString
    tmpOptions = SettingOptions(PBSTG_DO_NOT_HIDE)
    If Not SettingTypeByOptions(tmpOptions) = teText Or Not privEnumCompare(tmpOptions, optValueAsArray) Then
        ForceArrayFormat PBSTG_DO_NOT_HIDE, options:=tmpOptions, dataType:=optValueAsText
    End If
    
    ''  DEFAULT ZOOM FOR MAC/PC -- USED UNTIL 'AutoZoom' is disabled, or users set's custom preference
    If Not Exists(PBSTG_ZOOM_ENABLE) Then Setting(PBSTG_ZOOM_ENABLE) = True
    tmpOptions = SettingOptions(PBSTG_ZOOM_ENABLE)
    If Not SettingTypeByOptions(tmpOptions) = teBoolean Then
        ForceBooleanFormat PBSTG_ZOOM_ENABLE, tmpOptions
    End If
    
    ''  ZOOM DEFAULT MAC
    If Not Exists(PBSTG_ZOOM_MAC_DEF) Then Setting(PBSTG_ZOOM_MAC_DEF) = 100
    tmpOptions = SettingOptions(PBSTG_ZOOM_MAC_DEF)
    If Not SettingTypeByOptions(tmpOptions) = teDouble Then
        ForceDoubleFormat PBSTG_ZOOM_MAC_DEF, tmpOptions
    End If
    
    ''  ZOOM DEFAULT PC
    If Not Exists(PBSTG_ZOOM_PC_DEF) Then Setting(PBSTG_ZOOM_PC_DEF) = 80
    tmpOptions = SettingOptions(PBSTG_ZOOM_PC_DEF)
    If Not SettingTypeByOptions(tmpOptions) = teDouble Then
        ForceDoubleFormat PBSTG_ZOOM_PC_DEF, tmpOptions
    End If
    
    ''  ZOOM THRESHOLDS
    If Not Exists(PBSTG_ARRAY_ZOOM_LIMITS) Then
        ArraySetting(PBSTG_ARRAY_ZOOM_LIMITS) = Array(50, 300)
    End If
    tmpOptions = SettingOptions(PBSTG_ARRAY_ZOOM_LIMITS)
    If Not SettingTypeByOptions(tmpOptions) = teDouble Or Not privEnumCompare(tmpOptions, optValueAsArray) Then
        ForceArrayFormat PBSTG_ARRAY_ZOOM_LIMITS, dataType:=optValueAsDouble
    End If
    
End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''   Set Zoom Level
''   [zoomLevel] = Zoom Level.  Default is: 100.
''   If [wksht] arguement is missing, will be applied to all
''   worksheets in [wkbk]
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function Zoom( _
    Optional useUserSettingIfExists As Boolean = True _
    , Optional forceZoomValue _
    , Optional forceOSDefault As Boolean = False _
    , Optional singleWorksheet As Worksheet)

    On Error Resume Next
    '' CHECK VALID CONFIGURATION
    If Not ValidConfig Then Exit Function
    ''
    Dim applyZoom
    Dim failed As Boolean
    Dim evts As Boolean: evts = Me.Events: Me.Events = False
    If forceOSDefault Then
        applyZoom = ZoomOSDefault()
    ElseIf Not IsMissing(forceZoomValue) Then
        applyZoom = forceZoomValue
    ElseIf useUserSettingIfExists Then
        If Exists(PBSTG_USER_ZOOM) Then
            applyZoom = ZoomUserDefault()
        Else
            applyZoom = ZoomOSDefault()
        End If
    End If
    If applyZoom < 1 Then applyZoom = applyZoom * 100
    Dim zoomThresh As Variant, lbNum As Long
    zoomThresh = ZoomGetMinMaxThresholds
    lbNum = LBound(zoomThresh)
    If CDbl(applyZoom) < CDbl(zoomThresh(lbNum)) Then
        applyZoom = CDbl(zoomThresh(lbNum))
    End If
    If CDbl(applyZoom) > CDbl(zoomThresh(lbNum + 1)) Then
        applyZoom = CDbl(zoomThresh(lbNum))
    End If
    Dim visStatus As Variant, tmpWS As Worksheet, doZoom As Boolean, curWS As Worksheet
    Set curWS = stgwb.activeSheet
    Dim shtView As WorksheetView
    Dim checkWB As Workbook
    If Not singleWorksheet Is Nothing Then
        Set checkWB = singleWorksheet.Parent
    Else
        Set checkWB = stgwb
    End If
    For Each shtView In checkWB.Windows(1).SheetViews
        doZoom = False
        If Not singleWorksheet Is Nothing Then
            If shtView.Sheet Is singleWorksheet Then
                doZoom = True
            End If
        Else
            doZoom = True
        End If
        If doZoom Then
            visStatus = shtView.Sheet.visible
            shtView.Sheet.visible = xlSheetVisible
            shtView.Sheet.Activate
            shtView.Sheet.Parent.Windows(1).Zoom = applyZoom
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False
            shtView.Sheet.visible = visStatus
        End If
    Next shtView
    
Finalize:
    On Error Resume Next
        curWS.Activate
        Me.Events = evts
    If Err.number <> 0 Then Err.Clear
End Function

Public Property Get Events() As Boolean
    Events = Application.EnableEvents
End Property
Public Property Let Events(appEventsOn As Boolean)
    Application.EnableEvents = appEventsOn
End Property

Public Property Get EnableAutoZoom() As Boolean
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
    EnableAutoZoom = Setting(PBSTG_ZOOM_ENABLE)
End Property

Public Property Let EnableAutoZoom(zoomEnabled As Boolean)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
    Setting(PBSTG_ZOOM_ENABLE) = zoomEnabled
End Property

''  Get Default Zoom For Current OS
Public Property Get ZoomOSDefault() As Double
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
    Dim zoomKey: zoomKey = "PBSTG_" & PBSTG_OS & "_ZOOM_DEF"
    ZoomOSDefault = CDbl(Setting(zoomKey))
End Property
Public Property Let ZoomOSDefault(defZoom As Double)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
    Dim zoomKey: zoomKey = "PBSTG_" & PBSTG_OS & "_ZOOM_DEF"
    Setting(zoomKey) = defZoom
End Property


''  Default Zoom By User
Public Property Get ZoomUserDefault() As Double
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
    Dim zKey: zKey = CheckKey(PBSTG_USER_ZOOM)
    If Not Exists(zKey) Then
        Setting(zKey) = ZoomOSDefault
    End If
    ZoomUserDefault = Setting(zKey)
End Property
Public Property Let ZoomUserDefault(zoomPerc As Double)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Property
        ''
    Dim zKey: zKey = CheckKey(PBSTG_USER_ZOOM)
    Setting(zKey) = zoomPerc
    If Not SettingType(zKey) = teLong Then ForceLongFormat zKey
End Property


Public Function ZoomSetOSDefaults(Optional defaultZoom_PC, Optional defaultZoom_MAC)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
   If Not IsMissing(defaultZoom_PC) Then Setting(PBSTG_ZOOM_PC_DEF) = CDbl(defaultZoom_PC)
    If Not IsMissing(defaultZoom_MAC) Then Setting(PBSTG_ZOOM_MAC_DEF) = CDbl(defaultZoom_MAC)
End Function
Public Function ZoomSetMinMaxThresholds(minZoom, maxZoom)
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
    ArraySetting(PBSTG_ARRAY_ZOOM_LIMITS) = Array(CDbl(minZoom), CDbl(maxZoom))
End Function
Public Function ZoomGetMinMaxThresholds() As Variant
        '' CHECK VALID CONFIGURATION
        If Not ValidConfig Then Exit Function
        ''
    If Not Exists(PBSTG_ARRAY_ZOOM_LIMITS) Then
        ZoomSetMinMaxThresholds 50, 300
    End If
    ZoomGetMinMaxThresholds = Setting(PBSTG_ARRAY_ZOOM_LIMITS)
End Function

'' ENSURE WORKBOOK SETTINGS MATCH CONFIGURED VALUES
Private Function ConfigureWorkbook()
    On Error Resume Next
    If stgwb.ConnectionsDisabled And Setting(PBSTG_ENABLE_CONNECTIONS) = True Then
        stgwb.EnableConnections
    End If
    If stgwb.SaveLinkValues = False And Setting(PBSTG_SAVE_LINK_VALUES) = True Then
        stgwb.SaveLinkValues = True
    End If
    ''do not turn dev tools off, but turn on if true
    If Application.showDevTools = False And Setting(PBSTG_DEV_SHOW_TOOLS) = True Then
        Application.showDevTools = True
    End If
    Application.AutomationSecurity = msoAutomationSecurityLow
    Application.EnableCancelKey = xlErrorHandler
    Application.ExtendList = True
    Application.GenerateTableRefs = xlGenerateTableRefStruct
    Application.ReferenceStyle = xlA1
    
    
''  TODO: INVESTIGATE MANAGING PER WORKBOOK SESSION
    ''  Application.ShowMenuFloaties = True
    

End Function

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''
''   DEVELOPER UTILITIES
''  (IsDeveloper Must True To Run)
''
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
    '' LIST ALL SETTINGS KEY, AND DATA TYPE OF FORMATTED SETTING VALUE
    Public Function DEV_ListSettingsDataTypes(Optional keyNameLike As String = vbNullString)
        If Not IsDeveloper Then Exit Function
        If Count = 0 Then Exit Function
        Dim tARR() As Variant
        tARR = l_lo.DataBodyRange.value
        Dim i As Long
        Dim pbk As pbKey
        Dim tOutput As Boolean
        For i = LBound(tARR) To UBound(tARR)
            tOutput = True
            pbk = NewKey(tARR(i, pbSTGColumns.scSettingKey), optDefault)
            If Not keyNameLike = vbNullString Then
                tOutput = StringsMatchPriv(pbk.Key, keyNameLike, smContains)
            End If
            If tOutput Then
                Debug.Print "SETTING KEY: " & pbk.Key & " ** ValueType: " & TypeName(pbk.OldValFormatted)
            End If
        Next i
    End Function

    '' LIST ALL SETTINGS WITH SETTINGS OPTIONS SUMMARY AND CURRENT VALUES
    Public Function DEV_SettingsSummary(Optional keyNameLike As String = vbNullString)
        If Not IsDeveloper Then Exit Function
        If Count = 0 Then Exit Function
        On Error Resume Next
        Dim tARR() As Variant
        Dim pbKeys As New Collection
        '' Adding Array :   Key |ValueRaw|ValueFormatted|SettingOptions|Updated
        tARR = l_lo.DataBodyRange.value
        Dim i As Long
        Dim pbk As pbKey
        Dim tOutput As Boolean
        For i = LBound(tARR) To UBound(tARR)
            tOutput = True
            pbk = NewKey(tARR(i, pbSTGColumns.scSettingKey), optDefault)
            If Not keyNameLike = vbNullString Then
                tOutput = StringsMatchPriv(pbk.Key, keyNameLike, smContains)
            End If
            If tOutput Then
                 pbKeys.add Array(pbk.Key, pbk.OldVal, pbk.OldValFormatted, pbk.OldOptions, pbk.Updated)
            End If
        Next i
        If pbKeys.Count > 0 Then
            Dim tWB As Workbook, tws As Worksheet
            Set tWB = Workbooks.add
            Set tws = tWB.Worksheets(1)
            tws.Name = "pbSettings-Summary"
            tWB.Windows(1).DisplayGridlines = False
            Dim tRng As Range
            Set tRng = tws.Range("A1:K3")
            tRng.Merge
            tRng.HorizontalAlignment = xlHAlignLeft
            tRng.VerticalAlignment = xlVAlignCenter
            tRng.Font.Bold = True
            tRng.Font.color = 16724484
            tRng.Cells(1, 1) = CStr(Now) & "   pbSettings Summary From " & stgwb.FullName
            tRng.BorderAround
            tRng.WrapText = True
            Set tRng = tws.Range("B4").Resize(ColumnSize:=6)
            tRng(1, 1) = "SETTING KEY"
            tRng(1, 2) = "SETTING VALUE (RAW)"
            tRng(1, 3) = "SETTING UPDATED"
            tRng(1, 4) = "SETTING OPTIONS VALUE"
            tRng(1, 5) = "VALUE TYPE"
            tRng(1, 6) = "VALUE"
            tRng.BorderAround
            tRng.Font.Bold = True
            tRng.Font.color = 16724484
            
            Dim arrItem, lb As Long
            Set tRng = tRng.Resize(RowSize:=1)
            
            For Each arrItem In pbKeys
                lb = LBound(arrItem)
                Set tRng = tRng.Offset(rowOffset:=tRng.rows.Count)
                Set tRng = tRng.Resize(RowSize:=1)
                tRng(1, 1) = arrItem(lb)
                tRng(1, 2) = arrItem(lb + 1)
                tRng(1, 3) = arrItem(lb + 5)
                tRng(1, 4) = arrItem(lb + 4)
                tRng(1, 5) = TypeName(GetValue(arrItem(lb)))
                If Not privEnumCompare(arrItem(lb + 4), pbSettingOption.optValueAsArray) Then
                    tRng(1, 6) = GetValue(arrItem(lb))
                Else
                    tRng(1, 6) = SettingValueRaw(arrItem(lb))
                End If
                If Not privEnumCompare(arrItem(lb + 4), pbSettingOption.optValueAsArray) Then
                    tRng(2, 2) = "(Value Type):"
                    tRng(2, 3) = TypeName(GetValue(arrItem(lb)))
                    tRng(2, 4) = "(Current Value):"
                    tRng(2, 5) = GetValue(arrItem(lb))
                    
                End If
            Next arrItem
            tws.UsedRange.EntireColumn.AutoFit
            Dim colIdx As Long
            For colIdx = 1 To tws.UsedRange.Columns.Count
                If tws.UsedRange(1, colIdx).ColumnWidth > GetValue(StandardSettingEnum.stgSettingsSheetMaxColWidth) Then
                    tws.UsedRange(1, colIdx).ColumnWidth = GetValue(StandardSettingEnum.stgSettingsSheetMaxColWidth)
                End If
            Next
            tws.UsedRange.WrapText = True
            tws.UsedRange.EntireRow.AutoFit
        End If
    End Function



' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
'   CLASS INITIALIZE
'   Verify Settings Worksheet
'   Start FIRST RUN process (CreateSettingsSheet) If Necessary
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Private Sub Class_Initialize()

'    l_validCfg = VerifySettingsSheet
'    If Not ValidConfig Then
'        l_validCfg = CreateSettingsSheet
'    End If
'    If ValidConfig Then
'        CheckSettingsHealth
'        CheckDefaultCommonSettings
'        ConfigureWorkbook
'    End If
'    Set stgwb = ThisWorkbook

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


Private Sub stgwb_SheetDeactivate(ByVal Sh As Object)
    If Sh Is pbSettingsSheet Then
        If stdAutoHide = True Then
            pbSettingsSheet.visible = xlSheetVeryHidden
        End If
    End If
End Sub



Private Property Get SettingKeyColumn() As listColumn
    Set SettingKeyColumn = pbSettingsListObj.ListColumns(COLUMN_NAME_KEY)
End Property
Private Property Get SettingValueColumn() As listColumn
    Set SettingValueColumn = pbSettingsListObj.ListColumns(COLUMN_NAME_VALUE)
End Property
Private Property Get SettingOptionsColumn() As listColumn
    Set SettingOptionsColumn = pbSettingsListObj.ListColumns(COLUMN_NAME_OPTIONS)
End Property
Private Property Get SettingUpdatedColumn() As listColumn
    Set SettingUpdatedColumn = pbSettingsListObj.ListColumns(COLUMN_NAME_UPDATED)
End Property

' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''      Returns collection of [SettingKey] for matched options
''      If [findNonMatching] = False (Default), then settings are matched on all pbSettingOption
''          values in the enum
''      If [findNonMatching] = True, then settings are matched where any setting does not contain
''          any of the enum values
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
Public Function DEV_ListMatchingSettings(options As pbSettingOption, Optional findNonMatching As Boolean = False) As Collection
    Dim resp As New Collection
    Dim sKey
    For Each sKey In SettingKeyColumn.DataBodyRange
        If findNonMatching = False Then
            If privEnumCompare(SettingOptions(sKey), options, ecAnd) = True Then
                resp.add sKey
            End If
        Else
            If privEnumCompare(SettingOptions(sKey), options, ecOR) = False Then
                resp.add sKey
            End If
        End If
    Next sKey
    For Each sKey In resp
        Debug.Print sKey
    Next sKey
    Set DEV_ListMatchingSettings = resp
End Function

Public Property Get IgnoreDev() As Boolean
'   Returns True if conditional compilation argument is: IGNORE_DEV = 1
    #If IGNORE_DEV Then
        IgnoreDev = True
    #End If
End Property

