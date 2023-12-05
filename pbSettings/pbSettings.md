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




