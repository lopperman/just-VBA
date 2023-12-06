# pbSettings Help

***

ABOUT pbSettings
pbSettings ([pbSettings.cls](https://github.com/lopperman/just-VBA/blob/main/pbSettings/pbSettings.cls)) is a VBA class module, with no dependencies, that can be added to any MS Excel VBA Workbook.
Upon first use, a worksheet and listobject we be created automatically as the source of truth for setting keys and values.
Recommended method for working with pbSettings is to add the 2 following methods to any standard/basic module.  To use pbSettings, check the 'readiness' once to ensure pbSettings is configured, and then use 'pbStg.[Method]' for working with settings.

***

### Add pbSettingsReady and pbStg to a standard module
```
        Public Property Get pbSettingsReady() As Boolean
            On Error Resume Next
            If Not pbStg Is Nothing Then
                pbSettingsReady = pbStg.ValidConfig
            End If
        End Property

        Public Function pbStg() As pbSettings
            On Error Resume Next
            Static stgObj As pbSettings
            If stgObj Is Nothing Then
                Set stgObj = New pbSettings
            End If
            If Err.number = 0 Then
                If Not stgObj Is Nothing Then
                    If stgObj.ValidConfig Then
                        Set pbStg = stgObj
                    End If
                End If
            Else
                Err.Clear
            End If
        End Function
```
***

### To Confirm pbSettings is ready to use:
```
      If pbSettingReady = True Then
        'settings area is ready to be used
      End If
```
### EXAMPLES / SCENARIOS FOR USING pbSettings

_The examples below are actual implementations of how I (pbSettings author) use pbSettings_

**BUTTON_BEEPS** 
(and _MESSAGE_BEEPS_ and _INPUT_NEEDED_BEEPS_)

Every time a user presses a command button, the code that executes checks settings _for the current user_ to determine if a `Beep` will occur when they press a button.

I have a public constant that defines the button beep setting key:  
```
        Public Constant SETTING_BUTTON_BEEP as String = "BUTTON_BEEP"
```
In my 'startup' code settings are checked for default values for whatever user is logged in
```
        If pbStg.Exists(SETTING_BUTTON_BEEP,isUSERSpecific:=True) = False Then
            pbStg.SettingForUser(SETTING_BUTTON_BEEP) = True
        End if
```
If the login name for current users was "SmithJ", then the above code would create a new setting with setting key = "BUTTON_BEEPS_USER_SmithJ"

In the method that gets called when a button is pressed, this code executes first -- which checks if it should play a beep for the current user:
```
        If pbStg.SettingForUser(SETTING_BUTTON_BEEP) Then Beep
```

### Using pbSettings
Creating, modifying, and obtaining settings information can be performed by typing the following anywhere in your VBA Project:
```
      pbStg.[Public Method, Arguments]
```
See below for documentation and examples

***

## Object Model **Properties**
***
###  AutoHide
```
  Public Property Get AutoHide() As Boolean
  Public Property Let AutoHide(hideSheetDefault As Boolean)
```

_DESCRIPTION_

Get or Set whether the Worksheet containing settings information is visible or very hidden
When _setting_ AutoHide to `False` (Default Value), the Settings Worksheet will only be hidden if at least one other worksheet is currently Visible

_EXAMPLES:_
 
_Get Settings Visiblity_ 

```
    Dim isVisible as Boolean
    isVisible = pbStg.AutoHide
```

_Set Settings Visibility_ 

```
    pbStg.AutoHide = False
    pbStg.AutoHide = True
```

***

###  Count

    Public Property Get Count() As Long

_DESCRIPTION_

Read Only - Returns `Boolean`
Returns the number of settings being managed

_EXAMPLES:_
 
_Get Settings Count_ 

```
    Dim settingsCount as Long
    settingsCount = pbStg.Count
```

***

###  ModifiedEarliestDate

    Public Property Get ModifiedEarliestDate() As Variant

_DESCRIPTION_

Read Only - Returns `Variant Type Date`
Returns the earliest date from any setting 'ModifiedDate'

_EXAMPLES:_
 
_Get Earliest Modified Date_ 

```
    Dim modDt as Variant
    modDt = pbStg.ModifiedEarliestDate
```

***

###  ModifiedLatestDate

    Public Property Get ModifiedLatestDate() As Variant

_DESCRIPTION_

Read Only - Returns `Variant Type Date`
Returns the latest date from any setting 'ModifiedDate'

_EXAMPLES:_
 
_Get Latest Modified Date_ 

```
    Dim modDt as Variant
    modDt = pbStg.ModifiedLatestDate
```

***

###  pbSettingsListObj

    Public Property Get pbSettingsListObj() As ListObject

_DESCRIPTION_

Read Only - Returns `ListObject`
Returns the pbSettings ListObject

_EXAMPLES:_
 
_Get pbSettings ListObject_ 

```
    Dim lo as ListObject
    Set lo = pbStg.pbSettingsListObj
```

***

###  pbSettingsSheet

    Public Property Get pbSettingsSheet() As Worksheet

_DESCRIPTION_

Read Only - Returns `Worksheet`
Returns the pbSettings Worksheet

_EXAMPLES:_
 
_Get pbSettings Worksheet_ 

```
    Dim ws as Worksheet
    Set ws = pbStg.pbSettingsSheet
```

***

###  Setting

    Public Property Get Setting(ByVal stgKey)
    Public Property Let Setting(ByVal stgKey, ByVal stgVal)

_DESCRIPTION_

Gets or sets a Setting Value
Returns `Variant` -- If 'SettingType' IS `teNumeric` or `teDateTime` or `teBoolean`, the returned Variant will be of type `Double, Date, or Boolean`

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


_EXAMPLES:_
 
_Create, Update, Get Settings Values_ 

```
    Dim settingKey as String
    settingKey = "testDateSetting"
    pbStg.Setting(settingKey) = Now()

    Dim settingValue as Variant
    settingValue = pbStg.Setting(settingKey)

    If pbStg.Setting(settingKey) < Now() Then
        pbStg.Setting(settingKey) = Now()
    End If
```

***

###  SettingForOS

    Public Property Get SettingForOS(ByVal stgKey)
    Public Property Let SettingForOS(ByVal stgKey, ByVal stgVal)

_DESCRIPTION_

Gets or sets a Setting Value for a setting specific to PC or MAC OS
For example a PC user and Mac user might need differerent DEFAULT_ZOOM settings.  Example below shows how to implement this.

_EXAMPLES:_
 
_Managed different setting values between MAC and PC_

```
    'For a PC user, this example would create a `DEFAULT_ZOOM_OS_PC' Key
    'For a MAC user, this example would create a `DEFAULT_ZOOM_OS_MAC' Key

    Dim settingKey as String
    settingKey = "DEFAULT_ZOOM"
    pbStg.SettingForOS(settingKey) = 100

    Dim settingValue as Variant
    settingValue = pbStg.SettingForOS(settingKey)

    'to see how the key was created, you can use the 'CheckKey' method

```

***

###  SettingForUser

    Public Property Get SettingForUser(ByVal stgKey)
    Public Property Let SettingForUser(ByVal stgKey, ByVal stgVal)

_DESCRIPTION_

Gets or sets a Setting Value for a setting specific to current User
For example a user might need differerent DEFAULT_ZOOM settings.  Example below shows how to implement this.

_EXAMPLES:_
 
_Managed different setting values between users_

```
    'For a user JohnSmith, this example would create a `DEFAULT_ZOOM_USER_JOHNSMITH' Key
    'For a JetLi, this example would create a `DEFAULT_ZOOM_USER_JETLI' Key

    Dim settingKey as String
    settingKey = "DEFAULT_ZOOM"
    pbStg.SettingForUser(settingKey) = 100

    Dim settingValue as Variant
    settingValue = pbStg.SettingForUser(settingKey)

    'to see how the key was created, you can use the 'CheckKey' method

```

***

###  ValidConfig

    Public Property Get ValidConfig() As Boolean

_DESCRIPTION_

ReadOnly - Returns True if pbSettings Is Ready to Use

_EXAMPLES:_
 
_Return pbSettings ValidConfiguration Status_

```
    Dim isValid as Boolean
    isValid = pbStg.ValidConfig
```

***

## Object Model **Methods**
***
###  CheckKey
```
      Public Function CheckKey(ByVal stgKey, isOSSpecific As Boolean, isUSERSpecific As Boolean)
```

_DESCRIPTION_
```
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ '
''  Exposed for convenience -- return name of setting key
''      e.g. CheckKey("TEST",true,false) = "TEST_OS_MAC" OR "TEST_OS_PC"
''      e.g. CheckKey("TEST",false,true) =
''          "TEST_USER_[LoginName]"
' ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~ ~~~
```

###  Delete

    Public Function Delete(ByVal stgKey, Optional isOSSpecific As Boolean = False, Optional isUSERSpecific As Boolean = False)

_DESCRIPTION_

Deletes a Setting, if it exists

_EXAMPLES:_
 
```
    'DELETE SETTING KEY: "testSetting"
    pbStg.Delete "testSetting", False, False

    'DELETE SETTING KEY: "testSetting_OS_PC", or "testSetting_OS_MAC"
    pbStg.Delete "testSetting", True, False

    'DELETE SETTING KEY: "testSetting_USER_browerp"
    '** ASSUMES CURRENT USER LOGIN IS 'browerp'
    pbStg.Delete "testSetting", False, True

```
***

###  Exists

    Public Function Delete(ByVal stgKey, Optional isOSSpecific As Boolean = False, Optional isUSERSpecific As Boolean = False)

_DESCRIPTION_

Returns True if setting key exists

_EXAMPLES:_
 
```
    Dim keyExists as Boolean

    'Return True if 'testSetting' key exists
    keyExists = pbStg.Exists("testSetting")

    'Return True if 'testSetting_OS_PC' key exists, and is being tested on a PC
    keyExists = pbStg.Exists("testSetting",isOSSpecific:=True)

    'Return True if 'testSetting_USER_ browerp' key exists, and is being tested by user 'browerp'
    keyExists = pbStg.Exists("testSetting",isUserSpecific:=True)

```
***

###  ExportSettings

    Public Function ExportSettings(Optional wildcardSearch As Variant)

_DESCRIPTION_

Exports all settings to a new Workbook
If `wildcardSearch' has a value, then any setting where any column value contains `wildcardSearch` will be exported

_EXAMPLES:_
 
```
    'Export All Settings
    pbStg.ExportSettings

    'Export settings containing "_OS_"
    pbStg.ExportSettings "_OS_"
```
***





