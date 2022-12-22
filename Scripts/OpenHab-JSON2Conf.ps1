<#
.SYNOPSIS

Converts OpenHab JSONDB files to .things and .item files. This easily allows to create item configurations in the UI
and then export them to files where they can be further modified, linked to groups and so on.

.DESCRIPTION

If you run the script without any parameter, it will do nothing. You need to tell if it should create .things, .items or both.

If you do so, it will look for JSONDB files in its current directory (names MUST be the original names) and convert to allthings.things and/or allitems.items

OpenHab files allow single or multi line notation for configurations and metadata. By default, all values are on their own lines
which makes the files easy to read but quite lengthy. Use the ...SingleLine parameters to compress the parts you like.

By default, you will get no output. If you want, add -verbose. Be warned, it's massive...

.PARAMETER JSONFolder

Point the script to a different folder to look for JSONDB files. This enables direct processing of the files your OpenHab installation uses
without copying them before. Be aware that the .things/.items files will also be written to this folder unless you specify an -OutFolder.

You can even specify multiple folders at once, but as said: The script will place the output in these folders as well without an -OutFolder.

.PARAMETER OutFileBaseName

If you don't like allitems/allthings, specify your own basename. The files will then be named <basename>.things and <basename>.items

.PARAMETER OutFolder

By default, the script will save its results to the same JSONFolder (or the current script folder). If you want to redirect, use this parameter.
Obviously, if you specified multiple JSONFolder parameters, these will be overwritten by their ancestors because different naming is not implementet right now.

.PARAMETER CreateThings

Create allthings.things

.PARAMETER IncludeDefaultChannels

By default, channels without config parameters will not be included in the .things file to minimize clutter. If you want to include all channels, add this switch.

.PARAMETER CreateItems

Create allitems.items

.PARAMETER BridgeConfigSingleLine

Put all configuration parameters of bridges in the same line as the bridge itself (partially recommended).

.PARAMETER ThingConfigSingleLine

Put all configuration parameters of things in the same line as the thing itself (partially recommended).

.PARAMETER ChannelConfigSingleLine

Put all configuration parameters of thing channels in the same line as the channel itself (recommended).

.PARAMETER ItemMetaSingleLine

Put all channel links and metadata of items on the same line as the item itself (NOT recommended).

.PARAMETER MetaConfigSingleLine

Put all channel link configuration parameters and metadata configuration parameters  of items on the same line as the channel link/metadata itself (partially recommended).

.PARAMETER Filter

Use a regex of your choice to filter for things/items of interest. Mainly for testing purposes, but also useful if you want to maintain dedicated files for specific groups of things/items.

.EXAMPLE

.\OpenHab-JSON2Conf.ps1 -CreateThings -CreateItems -ThingConfigSingleLine -ChannelConfigSingleLine

This is the recommended parameter combination. It gives a quite compact .things list which you might not touch that often, and a comprehensive .items list which is easy to read and modify.

.EXAMPLE

.\OpenHab-JSON2Conf.ps1 -JSONFolder @( '\\server1\share\openhab\userdata\jsondb', '\\server2\share\openhab\userdata\jsondb' ) -CreateThings -CreateItems

Create fully line separated .things and .items for both OpenHab JSONDB in the given folders.

.EXAMPLE

.\OpenHab-JSON2Conf.ps1 -CreateThings -ThingConfigSingleLine -ChannelConfigSingleLine -filter 'avm'

Create a .things file that only deals with things and bridges that match to 'avm'. Welcome all Fritz!Box users :-)

#>
[CmdletBinding()]
param (
    [Parameter( ValueFromPipeline=$true,Position=0)]
    [ValidateScript( { Test-Path $_ } )]
    [String[]] $JSONFolder,
    [ValidateScript( { Test-Path $_ } )]
    [String] $OutFolder,
    [Switch] $CreateThings,
    [Switch] $IncludeDefaultChannels,
    [Switch] $BridgeConfigSingleLine,
    [Switch] $ThingConfigSingleLine,
    [Switch] $ChannelConfigSingleLine,
    [Switch] $CreateItems,
    [Switch] $ItemConfigSingleLine,
    [Switch] $ItemMetaSingleLine,
    [Switch] $MetaConfigSingleLine,
    [String] $OutFileBasename,
    [String] $Filter = '.*'
)

# for better understanding, some basic hints aboutn my coding style:
#
# all custom classes override the .ToString() method and accept 2 parameters:
# .ToString( [int] $Indent, [bool] $SingleLine  )
# $Indent is required for Things - these can be childs of Brindges, and for nice formatting, we then need to indent the whole thing by 2 spaces
# We also want nice formatting over all, so we leverage the $Indent value to accomplish that.
# $SingleLine controls if the result (enclosed in either [] or {} ) is returned on a single line or on separate lines for each value.
# child items (if present) will inherit the $SingleLine property so they cannot be returned on a not-single line ir the parent item is on a single line
# That's a result of the semi-JSON definition language for .things and .items
# The .ToString() methods never returns leading spaces or line breaks. This makes it possible to append its result in both $SingleLine and not.
#
# In the main functions, I am always returning by ,$return to prevent powershell from unwanted array conversion.
# This conversion is not broadly known and happens on most data types that represent arrays/lists.
#
# For class constructors, I mostly prefer "none" and add each property individually. Constructors have no named paramters which makes them different
# to read if you instanciate your class with a full set of properties like [Thing]::new( $Label, $location, $GindingID, $TypeID ),
# and it also makes the class definition itself harder to read because of a lot of constructors. KISS - keep it simple, stupid!

begin {

    # OpenHab needs UTF-8 or CP1252, we need this for the streamwriter
    $Encoding = [Text.Encoding]::GetEncoding( 1252 )

    class ohobject {
        # basic class for all oh objects
        # .class property only for sorting the results - makes using Sort-Object easier...
        hidden [string] $_Class
        ohobject() {
            $this._Class = $this.GetType().Name
            $this | Add-Member -MemberType ScriptProperty -Name 'Class' -Value {
                return $this._Class
            }
        }

        # all oh object classes have these methods to return properly formatted strings, so put them in a base class
        [String] ToString() {
            return $this._ToString( 0, $false )
        }
        [String] ToString( [int] $Indent ) {
            return $this._ToString( $Indent, $false )
        }
        [String] ToString( [int] $Indent, [bool] $SingleLine ) {
            return $this._ToString( $Indent, $SingleLine )
        }
        hidden [String] _ToString( [int] $Indent, [bool] $SingleLine ) {
            return $this
        }
    }

    class ohitem : ohobject {
        # if it is a real oh item, it always has a .configuration property to store custom config params
        [Configuration] $configuration = [Configuration]::new()
    }

    class Bridge : ohitem {

        # generic bridge definition:
        # Bridge <binding_name>:<bridge_type>:<bridge_name> "Displayname" @ "Location" [ <parameters> ] {
        #   <array of things>
        # }

        [String] $BindingID
        [String] $BridgeType
        [String] $BridgeID
        [String] $label
        [String] $location
        [Collections.ArrayList] $Things = [Collections.ArrayList]::new()

        hidden [String] _ToString( [int] $Indent, [bool] $SingleLine ) {

            # create the .things bridge definition from the bridge properties

            [String] $Return = $this.Class + ' ' + $this.BindingID + ':' + $This.BridgeType + ':' + $This.BridgeID
            if ( $This.label ) { $Return += ' "' + $this.label + '"' }
            if ( $This.location ) { $Return += ' @ "' + $this.location + '"' }

            # if the current item has configuration values, insert them in square brackets
            # the [Configuration] class (defined below) handles this.

            if ( $This.Configuration.Items.Count -gt 0 ) {
                # The [configuration].ToString method controls formatting
                $Return += ' ' + $this.Configuration.ToString( 0, $script:BridgeConfigSingleLine )
            }

            # if things use this bridge, include them in the bridge definition within curly braces

            if ( $this.Things.Count -gt 0 ) {
                $Return += " {`r`n"
                foreach ( $Thing in $this.Things ) {
                    $Return += $Thing.ToString( 2, $True ) # indentation 2 spaces more and the thing is a child of the bridge
                }
                $Return += "}`r`n"
            } else {
                $Return += "`r`n"
            }

            # add a final empty line for better reading

            return $Return + "`r`n"
        }
    }

    class Thing : ohitem {

        # generic bridge thing definition
        # Thing <type_id> <thing_id> "Displayname" @ "Location" [ <parameters> ]
        # generic standalone thing definition
        # Thing <binding_id>:<type_id>:<thing_id> "Displayname" @ "Location" [ <parameters> ]

        [String] $BindingID
        [String] $TypeID
        [String] $BridgeID
        [String] $ThingID
        [String] $label
        [String] $Location
        [Collections.Arraylist] $Channels = [Collections.ArrayList]::new()
        
        hidden [String] _ToString( [int] $Indent, [bool] $IsBridgeChild ) {

            # create the .things item definition from the item properties
            # if the item is a child of a bridge, the string composition is different

            $Spacing = ' ' * $Indent
            if ( $IsBridgeChild ) { 
                [String] $Return = $Spacing + $This.Class + ' ' + $This.TypeID + ' ' + $This.ThingID
            } else {
                # special handler for shelly devices...
                if ( $This.BindingID -eq 'shelly' ) {
                    [String] $Return = $Spacing + $This.Class + ' ' + $This.BindingID + ':shellydevice:' + $This.ThingID
                } else {
                    [String] $Return = $Spacing + $This.Class + ' ' + $This.BindingID + ':' + $This.TypeID + ':' + $This.ThingID
                }
            }
            if ( $This.label ) { $Return += ' "' + $This.label + '"' }
            if ( $This.location ) { $Return += ' @ "' + $This.location + '"' }

            # if the current item has configuration values, insert them in square brackets
            # the [Configuration] class (defined below) handles this.
            
            if ( $This.Configuration.Items.Count -gt 0 ) {
                # The configuration.ToString() method controls formatting - for items, this is an overwritten method of the base configuration.ToString() method
                $Return += ' ' + $This.Configuration.ToString( $Indent, $script:ThingConfigSingleLine )
            }
            
            # if the thing has channels, include them in curly braces
            
            if ( $This.Channels.Count -gt 0 ) {
                $Return += " {`r`n"
                $Return += $Spacing + "  Channels:`r`n"
                foreach ( $Channel in $This.Channels ) {
                    $Return += $Channel.ToString( $Indent + 4 )
                }
                $Return += $Spacing + '}'
            }
            $Return += "`r`n"
            If ( -not $IsBridgeChild ) {

                # add a final empty line for better reading if it is a standalone thing

                $Return += "`r`n"
            }

            return $Return
        }

    }

    class Channel : ohitem {

        # generic thing channel definition
        # Channels:
        #   State String : customChannel1 "My Custom Channel" [
        #     configParameter="Value",
        #     configParameter="Value"
        #   ]

        [String] $Kind
        [String] $Type
        [String] $ID
        [String] $Name
        
        hidden [String] _ToString( [int] $Indent, [bool] $SingleLine ){

            [string] $Return = ' ' * $Indent + $This.Kind.Substring( 0, 1 ).ToUpper() + $This.Kind.Substring( 1 ).ToLower() + ' ' + $This.Type + ' : ' + $This.ID
            if ( $This.Name ) {
                $Return += ' "' + $This.Name + '"'
            }
            $Return += ' ' + $This.Configuration.ToString( $Indent, $script:ChannelConfigSingleLine ) + "`r`n"
            return $Return
        }

    }

    class Item : ohitem {

        # generic item definition
        # itemtype itemname "labeltext [stateformat]" <iconname> (group1, group2, ...) ["tag1", "tag2", ...] {bindingconfig}
        # generic group definition
        # Group[:itemtype[:function]] groupname ["labeltext"] [<iconname>] [(group1, group2, ...)] [[ "semanticClass"]] [{<channel links>}]
    
        [String] $itemType
        [String] $Name
        [String] $label
        [String] $category
        [String] $iconName
        [Collections.ArrayList] $groups = [Collections.ArrayList]::new()
        [Collections.ArrayList] $tags = [Collections.ArrayList]::new()
        # all classes have a .configuration property. For items, we cannot use it since item configurations are not in [], but in {}, and they need a
        # different formatting for a clean look of the .items file. So add a new child class of [configuration]...
        [ItemConfiguration] $itemConfiguration = [ItemConfiguration]::new()
    
        # required for aggregate groups
        [string] $baseItemType 
        [string] $functionName
        [Collections.ArrayList] $functionParams = [Collections.ArrayList]::new()
    
        hidden [String] _ToString( [int] $Indent, [bool] $SingleLine ) {
    
            # item definition string in .items files as documented, see above
            [String] $Return = $This.itemType
    
            # handle aggregate groups - only these have a baseItemType and optionally an aggregate function
            if ( $This.baseItemType ) {
                $Return += ':' + $This.baseItemType
                if ( $This.functionName ) {
                    $Return += ':' + $This.functionName
                    if ( $This.functionParams ) {
                        if ( $This.functionName -eq 'COUNT' ){
                            # COUNT has a single item channel over which it aggregates
                            $Return += '"' + $This.functionParams + '"'
                        } else {
                            # all aggregate functions except COUNT have individual params that specify the aggregate values
                            $Return += '('
                            foreach ( $functionParam in $This.functionParams ) {
                                $Return += $FunctionParam + ','
                            }
                            $Return = $Return.Substring( 0, $Return.Length - 1 ) + ')' # strip last comma, close section
                        }
    
                    }
                }
            }
    
            $Return += ' ' + $This.name + ' "' + $This.label + '"'
            if ( $This.iconName ) {
                $Return += ' <' + $This.iconName + '>'
            }
            if ( $This.Groups.Count -ge 1 ) {
                $Return += ' ( '
                foreach ( $Group in $This.Groups ) {
                    $Return += $Group + ', '
                }
                $Return = $Return.Substring( 0, $Return.Length - 2 ) + ' )' # strip last comma, close section
            }
            if ( $This.tags.Count -ge 1 ) {
                $Return += ' [ '
                foreach ( $Tag in $This.tags ) {
                    $Return += '"' + $Tag + '", '
                }
                $Return = $Return.Substring( 0, $Return.Length - 2 ) + ' ]' # strip last comma, close section
            }
    
            if ( $This.itemConfiguration.Items.Count -gt 0 ) {
                # both channel links and metadata go into the same $Thing.configuration.Items $Property
                # The [configuration].ToString method controls formatting
                # items have no indendation, so 0 as first param. Second controls line breaks.
                $Return += ' ' + $This.itemConfiguration.ToString( 0, $script:ItemConfigSingleLine )
            }
            # add final new line for better reading
            return $Return + "`r`n"
        }
    }

    class ItemChannelLink : ohitem {

        # generic binding (aka "item channel link") definition
        # channel="<bindingID>:<thing-typeID>:MyThing:myChannel" [profile="<profileID>", <profile-parameter>="MyValue", ...]
    
        [String] $name
        [String] $uid
        [String] $itemName
    
        hidden [String] _ToString( [int] $Indent, [bool] $SingleLine ) {
            # special handling for $SingleLine -eq $true: No indendation at all...
            if ( $SingleLine ) { $Indent = 0 }
            $Spacing = ' ' * $Indent
            [String] $Return = $Spacing + 'channel="' + $This.uid + '"'
    
            if ( $This.Configuration.Items.Count -gt 0 ) {
                $Return += ' ' + $This.Configuration.ToString( $Indent + 2, $SingleLine -or $script:MetaConfigSingleLine )
            }
            return $Return
        }
    
    }
    
    class Metadata : ohitem {
    
        # generic metadata definition
        # metatype="metaname" [paramter=value, parameter=value, ...]
        [String] $name
        [String] $type
        [String] $value
        [String] $itemName

        hidden [String] _ToString( [int] $Indent, [bool] $SingleLine ) {
            # special handling for $SingleLine -eq $true: No indendation at all...
            if ( $SingleLine ) { $Indent = 0 }
            $Spacing = ' ' * $Indent
            [String] $Return = $Spacing + $This.type + '="' + $This.value + '"'
            if ( $This.Configuration.Items.Count -gt 0 ) {
                $Return += ' ' + $This.Configuration.ToString( $Indent + 2, $SingleLine -or $script:MetaConfigSingleLine )
            }
            return $Return
        }
    
    }
    
    class Config : ohitem {

        # items, bindings, etc. might have config values. These have  a name, value and (optionally) type
        # to make things easier, this class handles them and their types
        [String] $ValueType
        [String] $ValueName
        [String] $ValueData
        # item metadata configuration behaves weird in terms of data types... bool and decimal must be enclosed in quotes.
        # Thus we need to keep track if this config is from a metadata configuration and return properly formatted values in the .ToString() method
        [Bool] $isMetaConfig
    
        config ( [String] $ValueName, [String] $ValueData ) {
            $This.Valuename = $ValueName
            $This.ValueData = $ValueData
            # if no ValueType was provided, let's do our best to derive it from ValueData
            switch -regex ( $This.ValueData ) {
                '^(true|false)$' {
                    $This.ValueType = 'bool'
                    break
                }
                '^\d+$' {
                    $This.ValueType = 'int'
                    break
                }
                '^-?\d+([,.]\d+)?$' {
                    $This.ValueType = 'decimal'
                    break
                }
                default {
                    $This.ValueType = 'string'
                }
            }
            $This.isMetaConfig = $false
        }
        config ( [String] $ValueType, [String] $ValueName, [String] $ValueData ) {
            $This.ValueType = $ValueType
            $This.Valuename = $ValueName
            $This.ValueData = $ValueData
            $This.isMetaConfig = $false
        }
    
        config ( [String] $ValueType, [String] $ValueName, [String] $ValueData, [Bool] $isMetaConfig ) {
            $This.ValueType = $ValueType
            $This.Valuename = $ValueName
            $This.ValueData = $ValueData
            $This.isMetaConfig = $isMetaConfig
        }
    
        hidden [String] _ToString( [int] $Indent, [bool] $SingleLine ) {
            $Spacing = ' ' * $Indent
            [String] $Return = $Spacing + $This.ValueName + '='
            switch ( $This.ValueType ) {
                'int' {
                    $Return += $This.ValueData
                    break
                }
                'decimal' {
                    # for decimals, we need dot separated values in the .items file. This depends on the current locale,
                    # we need to force the decimal to be converted first to a single float depending on the current separator
                    # that ConvertFrom-JSON inserts, and then to the en-US string format where the separator is a dot.
                    $DecimalValue = $This.ValueData
                    if ( $DecimalValue -match ',') {
                        $DecimalValue = $DecimalValue.ToSingle( [cultureinfo]::new( 'de-DE' ))
                    } else {
                        $DecimalValue = $DecimalValue.ToSingle( [cultureinfo]::new( 'en-US' ))
                    }
                    $DecimalValue = $DecimalValue.ToString( [cultureinfo]::new('en-US' ))
                    # for item metadata configurations, decimals must be enclosed in quotes
                    if ( $This.isMetaConfig ) {
                        $DecimalValue = '"' + $DecimalValue + '"'
                    }
                    $Return += $DecimalValue
                    break
                }
                'bool' {
                    $BoolValue = $This.ValueData.ToString().ToLower()
                    # for item metadata configurations, bool must be enclosed in quotes
                    if ( $This.isMetaConfig ) {
                        $BoolValue = '"' + $BoolValue + '"'
                    }
                    $Return += $BoolValue
                    break
                }
                'string' {
                    # need to escape \ and " for semi-JSON used in .items and .things
                    $Return += '"' + $This.ValueData.Replace( '\', '\\' ).Replace( '"', '\"' ) + '"'
                    break
                }
                default {
                    # failed to resolve data type, return at least something...
                    $Return += $This.ValueData.ToString()
                }
            }
            return $Return
        }
    
    }

    class Configuration : ohobject {
        # helper class to format configuration items in a single line or multiple lines
        [Collections.ArrayList] $Items = [Collections.ArrayList]::new()
        
        hidden [String] _ToString( [int] $Indent, [bool] $SingleLine ) {
            [String] $Return = ''
            if ( $This.Items.Count -ge 1 ) {
                if ( $SingleLine ) {
                    [string] $Return = '[ '
                    foreach ( $Item in $This.Items ) {
                        $Return += $Item.ToString() + ', ' 
                    }
                    $Return = $Return.Substring( 0, $Return.Length - 2 ) + ' ]'
                } else {
                    $NewLine = "`r`n"
                    $Spacing = ' ' * $Indent
                    [string] $Return = '[' + $NewLine
                    foreach ( $Item in $This.Items ) {
                        $Return += $Spacing + '  ' + $Item.ToString( ) + ',' + $NewLine
                    }
                    $Return = $Return.Substring( 0, $Return.Length - ( 1 + $NewLine.Length ) ) + $NewLine
                    $Return += $Spacing + ']'
                }
            }
            return $Return
        }
    }

    class ItemConfiguration : Configuration {
        # item configuration is different from other configurations, here we have
        # channel links and metadata in curly braces, each of them on their own line
        # (can be concatenated, but that's a mess to read)
        hidden [String] _ToString( [int] $Indent, [bool] $SingleLine ) {
            $Return = ''
            if ( $This.Items.Count ) {
                if ( $SingleLine ) {
                    [string] $Return = '{ '
                    foreach ( $Item in $This.Items ) {
                        $Return += $Item.ToString( $Indent + 2, $SingleLine ) + ', ' 
                    }
                    $Return = $Return.Substring( 0, $Return.Length - 2 ) + ' }'
                } else {
                    $NewLine = "`r`n"
                    $Spacing = ' ' * $Indent
                    [string] $Return = '{' + $NewLine
                    foreach ( $Item in $This.Items ) {
                        $Return += $Spacing + '  ' + $Item.ToString( $Indent + 2 ) + ',' + $NewLine
                    }
                    $Return = $Return.Substring( 0, $Return.Length - ( 1 + $NewLine.Length ) ) + $NewLine
                    $Return += $Spacing + '}'
                }
            }
            return $Return
        }
    }

}

process {

    function Convert-ConfigurationFromJSON {
        # convert configuration JSON to the object type we need
        [CmdletBinding()]
        param (
            [Object] $ConfigurationJSON,
            [Bool] $isMetaConfig = $false
        )
        $Configurations = [Collections.ArrayList]::new()
        foreach ( $Config in $ConfigurationJSON | Get-Member -MemberType NoteProperty ) {
            # parse the actual NoteProperty to its individual partes and add it to the $Configurations arraylist
            if ( $Config.Definition -match "^(?<ValueType>\w+)\s+$( $Config.Name )=(?<ValueData>.+)$" ) {
                Write-Verbose "Processing configuration: $( $Config.Definition )"
                $ConfigValue = [Config]::new( $Matches.ValueType, $Config.Name, $Matches.ValueData, $isMetaConfig )
                [void] $Configurations.Add( $ConfigValue )
            }
        }
        return ,$Configurations
    }

    function Get-Things {
        [CmdletBinding()]
        param (
            [Object] $ThingsJSON,
            [String] $Filter
        )

        $Things = [Collections.ArrayList]::new()
        $Bridges = [Collections.ArrayList]::new()
        $Properties = $ThingsJSON | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $Filter }

        $ThingsCounter = 0
        foreach ( $Property in $Properties ) {
            $ThingsCounter++
            $JSON = $ThingsJSON."$( $Property.Name )"

            Write-Progress "Processing bridge things ($ThingsCounter/$($Properties.Count))" -Status $Property.Name -PercentComplete ( $ThingsCounter * 100 / $Properties.Count )
        
            # make sure it is a bridge
        
            if ( $JSON.value.isBridge ) {
        
                Write-Verbose "Processing bridge: $( $JSON.value.UID )"

                $Bridge = [Bridge]::new()
                $Bridge.label = $JSON.value.label
                $Bridge.location = $JSON.value.location
                $Bridge.BindingID = $JSON.value.UID.Split( ':', 3 )[0]
                $Bridge.BridgeType = $JSON.value.UID.Split( ':', 3 )[1]
                $Bridge.BridgeID = $JSON.value.UID.Split( ':', 3 )[2]
                $Bridge.Configuration.Items = Convert-ConfigurationFromJSON -ConfigurationJSON $JSON.value.Configuration

                [void] $Things.Add( $Bridge )
                [void] $Bridges.Add( $Property.Name )
            }
        }

        $Properties = $ThingsJSON | Get-Member -MemberType NoteProperty | Where-Object { $Bridges -notcontains $_.Name } | Where-Object { $_.Name -match $Filter }
        $ThingsCounter = 0
        foreach ( $Property in $Properties ) {
            $ThingsCounter++
            $JSON = $ThingsJSON."$( $Property.Name )"
        
            Write-Verbose "Processing thing: $( $JSON.value.UID )"
            Write-Progress "Processing standalone things ($ThingsCounter/$($Properties.Count))" -Status $Property.Name -PercentComplete ( $ThingsCounter * 100 / $Properties.Count )

            $Thing = [Thing]::new()
            $Thing.label = $JSON.value.label
            $Thing.location = $JSON.value.location
            $Thing.BindingID = $JSON.value.UID.Split( ':', 4 )[0]
            $Thing.TypeID = $JSON.value.UID.Split( ':', 4 )[1]
            if ( $JSON.value.BridgeUID ) {
                # if the thing uses a bridge, the bridge ID will be part of its UID which thus contains 4 segments...
                $Thing.BridgeID = $JSON.value.UID.Split( ':', 4 )[2]
                $Thing.ThingID = $JSON.value.UID.Split( ':', 4 )[3]
            } else {
                # ...if it is a standalone thing, its UID will only contain 3 segments and we have no BridgeID
                $Thing.ThingID = $JSON.value.UID.Split( ':', 4 )[2]
            }
        
            $Thing.Location = $JSON.value.location
            $Thing.Configuration.Items = Convert-ConfigurationFromJSON -ConfigurationJSON $JSON.value.Configuration
       
            foreach ( $Ch in $JSON.value.channels ) {
                Write-Verbose "Processing thing channel: $( $Ch.UID )"
                $Channel = [Channel]::new()
                $Channel.Name = $Ch.Label
                $Channel.Kind = $Ch.Kind
                if ( $Ch.Kind -eq 'TRIGGER' ) {
                    # trigger channels do not define their type because it must be 'String'
                    $Channel.Type = 'String'
                } else {
                    $Channel.Type = $Ch.itemType
                }
                If ( $ch.uid -match ':(?<ID>[^:]+)$' ) {
                    # channel ID needs to be extracted from channel UID - always last segment (after last ':')
                    $Channel.ID = $Matches.ID
                }
                $Channel.Configuration.Items = Convert-ConfigurationFromJSON -ConfigurationJSON $Ch.Configuration
       
                # only add the channel if any configurations were found unless -IncludeDefaultChannels
                # all standard channels (without configuration) will be added anyway by the thing binding automatically
                if ( ( $Channel.Configuration.Items.Count -gt 0 ) -or $IncludeDefaultChannels ) {
                    [void] $Thing.Channels.Add( $Channel )
                }
            }

            if ( $JSON.value.BridgeUID ) {
                # this thing uses a bridge, so assign it to its bridge object
                $BindingID = $JSON.value.BridgeUID.Split( ':', 3 )[0]
                $BridgeType = $JSON.value.BridgeUID.Split( ':', 3 )[1]
                $BridgeID = $JSON.value.BridgeUID.Split( ':', 3 )[2]
                $Bridge = $Things | Where-Object { $_.BindingID -eq $BindingID -and $_.BridgeType -eq $BridgeType -and $_.BridgeID -eq $BridgeID }
                # no need to check if we found any bridge - we WILL find one and only one...
                [void] $Bridge.Things.Add( $Thing )
            } else {
                [Void] $Things.Add( $Thing )
            }
        }
        return ,$Things
    }

    function Get-ItemChannelLinks {
        [CmdletBinding()]
        param (
            [Object] $ItemChannelLinksJSON,
            [String] $Filter
        )

        $ItemChannelLinks = [Collections.ArrayList]::new()
        foreach ( $Property in $ItemChannelLinksJSON | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $Filter } ) {

            $JSON = $ItemChannelLinksJSON."$( $Property.Name )"
            Write-Verbose "Processing ItemChannelLink: $( $JSON.value.ChannelUID.UID )"

            $ItemChannelLink = [ItemChannelLink]::new()
            $ItemChannelLink.name = $Property.Name
            $ItemChannelLink.uid = $JSON.value.ChannelUID.UID
            $ItemChannelLink.itemName = $JSON.value.itemName
            $ItemChannelLink.Configuration.Items = Convert-ConfigurationFromJSON -ConfigurationJSON $JSON.value.Configuration.properties

            [void] $ItemChannelLinks.Add( $ItemChannelLink )
        }
        return ,$ItemChannelLinks
    }

    function Get-Metadata {
        [CmdletBinding()]
        param (
            [Object] $MetadataJSON,
            [String] $Filter
        )

        $Metadatas = [Collections.Arraylist]::new()
        foreach ( $Property in $MetadataJSON | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $Filter } ) {
    
            $JSON = $MetadataJSON."$( $Property.Name )"
            Write-Verbose "Processing metadata: $( $Property.Name )"
        
            $MetaData = [Metadata]::new()
            $MetaData.name = $Property.Name
            $Metadata.type = $Property.Name.Split( ':' )[0]
            $Metadata.itemName = $Property.Name.Split( ':', 2 )[1]
            $MetaData.value = $JSON.value.value
            $Metadata.Configuration.Items = Convert-ConfigurationFromJSON -ConfigurationJSON $JSON.value.configuration
        
            [void] $MetaDatas.Add( $MetaData )
        }
        return ,$Metadatas
    }

    function Get-Items {
        [CmdletBinding()]
        param (
            [Object] $ItemsJSON,
            [Object] $ItemChannelLinksJSON,
            [Object] $MetadataJSON,
            [String] $Filter
        )

        # we need bindings and metadata to assign them to their items later
        $ItemChannelLinks = Get-ItemChannelLinks -ItemChannelLinksJSON $ItemChannelLinksJSON -Filter $Filter
        $Metadata = Get-Metadata -MetadataJSON $MetadataJSON -Filter $Filter
        
        $Items = [Collections.ArrayList]::new()
        $ItemCounter = 0
        $Properties = $ItemsJSON | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $Filter }
        foreach ( $Property in $Properties ) {
            $ItemCounter++
            $JSON = $ItemsJSON."$( $Property.Name )"
            Write-Verbose "Processing item: $( $Property.Name )"
            Write-Progress "Processing items ($ItemCounter/$($Properties.Count))" -Status $Property.Name -PercentComplete ( $ItemCounter * 100 / $Properties.Count )

            $Item = [Item]::new()
            $Item.Name = $Property.Name
            $Item.itemType = $JSON.value.itemType
            $Item.label = $JSON.value.label
            $Item.iconname = $JSON.value.category
            $Item.baseItemType = $JSON.value.baseItemType
            $Item.functionName = $JSON.value.functionName
        
            foreach ( $functionParam in $JSON.value.functionParams ) {
                Write-Verbose "Processing item function param: $( $functionParam )"
                [void] $Item.functionParams.Add( $functionParam )
            }
            foreach ( $Group in $JSON.value.groupNames ) {
                Write-Verbose "Processing item group: $( $group )"
                [void] $Item.groups.Add( $Group )
            }
            foreach ( $Tag in $JSON.value.tags ) {
                Write-Verbose "Processing item tag: $( $tag )"
                [void] $Item.tags.Add( $Tag )
            }
            foreach ( $ItemChannelLink in $ItemChannelLinks | Where-Object { $_.itemName -eq $Item.Name } ) {
                Write-Verbose "Processing ItemChannelLink: $( $ItemChannelLink )"
                [void] $Item.itemConfiguration.Items.Add( $ItemChannelLink )
            }
            foreach ( $Meta in $Metadata | Where-Object { $_.itemName -eq $Item.Name } ) {
                Write-Verbose "Processing item metadata: $( $Meta )"
                [void] $Item.itemConfiguration.Items.Add( $Meta )
            }

            [void] $Items.Add( $Item )
        }
        return ,$Items
    }
    
    if ( -not $JSONFolder ) {
        $JSONFolder = $PSScriptRoot
    }
    if ( -not $OutFolder ) {
        $OutFolder = $JSONFolder
    }

    if ( $CreateThings ) {
        Write-Progress "Reading $JSONFolder\org.openhab.core.thing.Thing.JSON"
        $ThingsJSON = Get-Content "$JSONFolder\org.openhab.core.thing.Thing.JSON" | ConvertFrom-Json
        $Things = Get-Things -ThingsJSON $ThingsJSON -Filter $Filter
        if ( $OutFileBasename ) {
            $OutFile = "$OutFolder\$OutfileBasename.things"
        } else {
            $OutFile = "$OutFolder\allthings.things"
        }
        $streamWriter = [IO.StreamWriter]::new( $Outfile, $false, $Encoding )
        # sort by Class (Bridge, Thing) amd ID for pretty reading
        foreach ( $Thing in $Things | Sort-Object -Property Class, BindingID ) {
            $streamwriter.Write( $Thing.ToString() )
        }
        $streamWriter.Dispose()
    }
    If ( $CreateItems ) {
        Write-Progress "Reading $JSONFolder\org.openhab.core.items.Item.JSON"
        $ItemsJSON = Get-Content "$JSONFolder\org.openhab.core.items.Item.JSON" | ConvertFrom-Json
        Write-Progress "Reading $JSONFolder\org.openhab.core.thing.Link.ItemChannelLink.JSON"
        $ItemChannelLinksJSON = Get-Content "$JSONFolder\org.openhab.core.thing.Link.ItemChannelLink.JSON" | ConvertFrom-Json
        Write-Progress "Reading $JSONFolder\org.openhab.core.items.Metadata.JSON"
        $MetadataJSON = Get-Content "$JSONFolder\org.openhab.core.items.Metadata.JSON" | ConvertFrom-Json
        $Items = Get-Items -ItemsJSON $ItemsJSON -ItemChannelLinksJSON $ItemChannelLinksJSON -MetadataJSON $MetadataJSON -Filter $Filter
        if ( $OutFileBasename ) {
            $OutFile = "$OutFolder\$OutfileBasename.items"
        } else {
            $OutFile = "$OutFolder\allitems.items"
        }
        $streamWriter = [IO.StreamWriter]::new( $Outfile, $false, $Encoding )
        # sort by type and name for pretty reading
        forEach ( $Item in $Items | Sort-Object -Property ItemType, Name ) {
            $streamWriter.WriteLine( $Item.ToString() )
        }
        $streamWriter.Dispose()
    }

    
}

end {

}
