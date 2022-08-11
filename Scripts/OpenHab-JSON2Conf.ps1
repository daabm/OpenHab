<#
.SYNOPSIS

Converts OpenHab JSONDB files to .things and .item files. This easily allows to create item configurations in the UI
and then export them to files where they can be further modified, linked to groups and so on.

.DESCRIPTION

If you run the script without any parameter, it will do nothing. You need to tell if it should create .things, .items or both.

If you do so, it will look for JSONDB files in its current directory (names MUST be the original names) and convert to allthings.things and/or allitems.items

OpenHab files allow single or multi line notation for configurations and metadata. By default, all values are on their own lines
which makes the files easy to read but quite lengthy. Use the ...SingleLine parameters to compress the parts you like.

.PARAMETER JSONFolder

Point the script to a different folder to look for JSONDB files. This enables to directly process the current files your OH installation uses
without copying them before. Be aware that the things/items files will also be written to this folder.

You can even specify multiple folders at once, but as said: The script will place the output in these folders as well.

.PARAMETER CreateThings

Create allthings.things

.PARAMETER CreateItems

Create allitems.items

.PARAMETER OutFileBaseName

If you don't like allitems/allthings, specify your own basename. The files will then be named <basename>.things and <basename>.items

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

#>
[CmdletBinding()]
param (
    [Parameter( ValueFromPipeline=$true,Position=0)]
    [ValidateScript( { Test-Path $_ } )]
    [String[]] $JSONFolder,
    [Switch] $CreateThings,
    [Switch] $BridgeConfigSingleLine,
    [Switch] $ThingConfigSingleLine,
    [Switch] $ChannelConfigSingleLine,
    [Switch] $CreateItems,
    [Switch] $ItemMetaSingleLine,
    [Switch] $MetaConfigSingleLine,
    [String] $OutFileBasename,
    [String] $Filter = '.*'
)

begin {

    $Encoding = [Text.Encoding]::GetEncoding( 1252 )

    class Bridge {

        # generic bridge definition:
        # Bridge <binding_name>:<bridge_type>:<bridge_name> [ <parameters> ] {
        #   (array of things)
        # }

        # basic bridge properties
        [String] $BindingID
        [String] $BridgeType
        [String] $BridgeID
        [String] $label
        [String] $location
        [Configuration] $Configuration = [Configuration]::new()
        [Collections.ArrayList] $Things = [Collections.ArrayList]::new()

        [String] ToString() {
            Return $This.ToStringInternal()
        }

        [String] Hidden ToStringInternal() {

            # create the .things bridge definition from the bridge properties
            [String] $Return = $This.Class + ' ' + $This.BindingID + ':' + $This.BridgeType + ':' + $This.BridgeID
            If ( $This.label ) { $Return += ' "' + $This.label + '"' }
            If ( $This.location ) { $Return += ' @ "' + $This.location + '"' }

            # if the bridge has configuration values, insert them in square brackets
            # and take care of indentation - Openhab is quite picky about misalignment :-)

            If ( $This.Configuration.Items.Count -gt 0 ) {
                $Return += ' ' + $This.Configuration.ToString( 0, $script:BridgeConfigSingleLine )
            }

            # if there are things using this bridge, include them in the bridge definition within curly brackets

            If ( $This.Things.Count -gt 0 ) {
                $Return += " {`r`n"
                Foreach ( $Thing in $This.Things ) {
                    $Return += $Thing.ToString( 2, $True ) # indentation 2 spaces more and the thing is a child of the bridge
                }
                $Return += "}`r`n"
            } Else {
                $Return += "`r`n"
            }
            # add a final empty line for better reading
            Return $Return + "`r`n"
        }
    }

    Class Thing {

        # generic bridge thing definition
        # Thing <type_id> <thing_id> "Label" @ "Location" [ <parameters> ]
        # generic standalone thing definition
        # Thing <binding_id>:<type_id>:<thing_id> "Label" @ "Location" [ <parameters> ]

        # basic thing properties
        [String] $BindingID
        [String] $TypeID
        [String] $BridgeID
        [String] $ThingID
        [String] $label
        [String] $Location
        [Configuration] $Configuration = [Configuration]::new()
        [Collections.Arraylist] $Channels = [Collections.ArrayList]::new()
        
        [String] ToString( ){
            Return $This.ToStringInternal( 0, $False )
        }
        [String] ToString( [int] $Indent ) {
            Return $This.ToStringInternal( $Indent, $False )
        }
        [String] ToString( [int] $Indent, [bool] $IsBridgeChild ) {
            Return $This.ToStringInternal( $Indent, $IsBridgeChild )
        }
        [String] Hidden ToStringInternal( [int] $Indent, [bool] $IsBridgeChild ) {

            # create the .things item definition from the item properties

            # if the item is a child of a bridge, we want 2 spaces more at the beginning of each line
            # and we have a different string composition

            $Spacing = ' ' * $Indent
            If ( $IsBridgeChild ) { 
                [String] $Return = $Spacing + $This.Class + ' ' + $This.TypeID + ' ' + $This.ThingID
            } Else { 
                [String] $Return = $Spacing + $This.Class + ' ' + $This.BindingID + ':' + $This.TypeID + ':' + $This.ThingID
            }
            If ( $This.label ) { $Return += ' "' + $This.label + '"' }
            If ( $This.location ) { $Return += ' @ "' + $This.location + '"' }

            # if the thing has configuration values, add them in square brackets
    
            If ( $This.Configuration.Items.Count -gt 0 ) {
                $Return += ' ' + $This.Configuration.ToString( $Indent, $script:ThingConfigSingleLine )
            }
            
            # if the thing has channels, include them in curly brackets as well
            # again, take care of correct indendation
            
            If ( $This.Channels.Count -gt 0 ) {
                $Return += " {`r`n"
                $Return += $Spacing + "  Channels:`r`n"
                Foreach ( $Channel in $This.Channels ) {
                    $Return += $Channel.ToString( $Indent + 4 )
                }
                $Return += $Spacing + '}'
            }
            $Return += "`r`n"
            If ( -not $IsBridgeChild ) {
                # add a final empty line for better reading if it is a standalone thing
                $Return += "`r`n"
            }
            Return $Return
        }

    }

    Class Channel {

        # generic thing channel definition
        # Channels:
        #   State String : customChannel1 "My Custom Channel" [
        #     configParameter="Value"
        #   ]

        # channel definition in .item files as documented, see above
        [String] $Kind
        [String] $Type
        [String] $ID
        [String] $Name
        [Configuration] $Configuration = [Configuration]::new()
        
        [String] ToString( ){
            Return $This.ToStringInternal( 4 )
        }
        [String] ToString( [int] $Indent ) {
            Return $This.ToStringInternal( $Indent )
        }

        [String] Hidden ToStringInternal( [int] $Indent ){

            [String] $Return = ''

            # create the .things channel definition from the channel properties
            # if the channel has configuration values, append them in square brackets
            
            If ( $This.Configuration.Items.Count -gt 0 ) {
                $Return += ' ' * $Indent + $This.Kind.Substring( 0, 1 ).ToUpper() + $This.Kind.Substring( 1 ).ToLower() + ' ' + $This.Type + ' : ' + $This.ID
                If ( $This.Name ) {
                    $Return += ' "' + $This.Name + '"'
                }
                $Return += ' ' + $This.Configuration.ToString( $Indent, $script:ChannelConfigSingleLine ) + "`r`n"
            }
            Return $Return
        }

    }

    class Item {

        # generic item definition
        # itemtype itemname "labeltext [stateformat]" <iconname> (group1, group2, ...) ["tag1", "tag2", ...] {bindingconfig}
        # generic group definition
        # Group[:itemtype[:function]] groupname ["labeltext"] [<iconname>] [(group1, group2, ...)] [[ "semanticClass"]] [{<channel links>}]
    
        # basic item properties
        [String] $itemType
        [String] $Name
        [String] $label
        [String] $category
        [String] $iconName
        [Collections.ArrayList] $groups = [Collections.ArrayList]::new()
        [Collections.ArrayList] $tags = [Collections.ArrayList]::new()
        [ItemConfiguration] $configuration = [ItemConfiguration]::new()
    
        # required for aggregate groups
        [string] $baseItemType 
        [string] $functionName
        [Collections.ArrayList] $functionParams = [Collections.ArrayList]::new()
    
        [String] ToString( ) {
            Return $This.ToStringInternal( )
        }
    
        [String] Hidden ToStringInternal( ) {
    
            # item definition string in .items files as documented, see above
            [String] $Return = $This.itemType
    
            # handle aggregate groups - only these have a baseItemType and optionally an aggregate function
            If ( $This.baseItemType ) {
                $Return += ':' + $This.baseItemType
                If ( $This.functionName ) {
                    $Return += ':' + $This.functionName
                    If ( $This.functionParams ) {
                        If ( $This.functionName -eq 'COUNT' ){
                            # COUNT has a single item channel over which it aggregates
                            $Return += '"' + $This.functionParams + '"'
                        } Else {
                            # all aggregate functions except COUNT have individual params that specify the aggregate values
                            $Return += '('
                            Foreach ( $functionParam in $This.functionParams ) {
                                $Return += $FunctionParam + ','
                            }
                            $Return = $Return.Substring( 0, $Return.Length - 1 ) + ')' # strip last comma, close section
                        }
    
                    }
                }
            }
    
            $Return += ' ' + $This.name + ' "' + $This.label + '"'
            If ( $This.iconName ) {
                $Return += ' <' + $This.iconName + '>'
            }
            If ( $This.Groups.Count -ge 1 ) {
                $Return += ' ( '
                Foreach ( $Group in $This.Groups ) {
                    $Return += $Group + ', '
                }
                $Return = $Return.Substring( 0, $Return.Length - 2 ) + ' )' # strip last comma, close section
            }
            If ( $This.tags.Count -ge 1 ) {
                $Return += ' [ '
                Foreach ( $Tag in $This.tags ) {
                    $Return += '"' + $Tag + '", '
                }
                $Return = $Return.Substring( 0, $Return.Length - 2 ) + ' ]' # strip last comma, close section
            }
    
            If ( $This.configuration.Items.Count -gt 0 ) {
                # both bindings and metadata go into the same $Thing.metadata $Property
                # indent by 2 spaces
                $Return += ' ' + $This.Configuration.ToString( 0, $script:ItemMetaSingleLine )
            }
            # add final new line for better reading
            Return $Return + "`r`n"
        }
    }
    
    class Binding {

        # generic binding (aka "item channel") definition
        # channel="<bindingID>:<thing-typeID>:MyThing:myChannel"[profile="system:<profileID>", <profile-parameterID>="MyValue", ...]
    
        # basic binding properties
        [String] $name
        [String] $uid
        [String] $itemName
        [Configuration] $Configuration = [Configuration]::new()
    
        [String] ToString() {
            Return $This.ToStringInternal( 0 )
        }
        [String] ToString( [int] $Indent ) {
            Return $This.ToStringInternal( $Indent )
        }
    
        [String] Hidden ToStringInternal( [int] $Indent ) {
            # binding definition string in .items files as documented, see above
            $Spacing = ' ' * $Indent
            [String] $Return = $Spacing + 'channel="' + $This.uid + '"'
    
            If ( $This.Configuration.Items.Count -gt 0 ) {
                $Return += ' ' + $This.Configuration.ToString( $Indent + 2, $script:MetaConfigSingleLine )
            }
            Return $Return
        }
    
    }
    
    Class Metadata {
    
        # basic metadata properties
        [String] $name
        [String] $type
        [String] $value
        [String] $itemName
        [Configuration] $Configuration = [Configuration]::new()
    
        [String] ToString() {
            Return $This.ToStringInternal( 0 )
        }
        [String] ToString( [int] $Indent ) {
            Return $This.ToStringInternal( $Indent )
        }
        [String] Hidden ToStringInternal ( [int] $Indent ) {
            $Spacing = ' ' * $Indent
            [String] $Return = $Spacing + $This.type + '="' + $This.value + '"'
            If ( $This.Configuration.Items.Count -gt 0 ) {
                $Return += ' ' + $This.Configuration.ToString( $Indent, $script:MetaConfigSingleLine )
            }
            Return $Return
        }
    
    }
    
    Class Config {

        # items, bindings, etc. might have config values. These consist of a name, a value and (optionally) a type
        # to make things easier, this class handles them and their types
        [String] $ValueType
        [String] $ValueName
        [String] $ValueData
        # item meta configuration behaves weird in terms of data types... bool and decimal must be enclosed in quotes.
        [Bool] $isMetaConfig
    
        Config ( [String] $ValueName, [String] $ValueData ) {
            $This.Valuename = $ValueName
            $This.ValueData = $ValueData
            # if no ValueType was provided, let's do our best to derive it from ValueData
            Switch -regex ( $This.ValueData ) {
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
        Config ( [String] $ValueType, [String] $ValueName, [String] $ValueData ) {
            $This.ValueType = $ValueType
            $This.Valuename = $ValueName
            $This.ValueData = $ValueData
            $This.isMetaConfig = $false
        }
    
        Config ( [String] $ValueType, [String] $ValueName, [String] $ValueData, [Bool] $isMetaConfig ) {
            $This.ValueType = $ValueType
            $This.Valuename = $ValueName
            $This.ValueData = $ValueData
            $This.isMetaConfig = $isMetaConfig
        }
    
        [String] ToString() {
            Return $This.ToStringInternal( 0 )
        }
        [String] ToString( [int] $Indent) {
            Return $This.ToStringInternal( $Indent )
        }
        [String] Hidden ToStringInternal ( $Indent) {
            $Spacing = ' ' * $Indent
            [String] $Return = $Spacing + $This.ValueName + '='
            Switch ( $This.ValueType ) {
                'int' {
                    $Return += $This.ValueData
                    break
                }
                'decimal' {
                    # for decimals, we need dot separated values in the item file. This depends on the current locale,
                    # we need to force the decimal to be converted first to a single float depending on the current separator
                    # that ConvertFrom-JSON insert, and then to the en-US string format where the separator is a dot.
                    $DecimalValue = $This.ValueData
                    If ( $DecimalValue -match ',') {
                        $DecimalValue = $DecimalValue.ToSingle( [cultureinfo]::new( 'de-DE' ))
                    } Else {
                        $DecimalValue = $DecimalValue.ToSingle( [cultureinfo]::new( 'en-US' ))
                    }
                    # for item configurations, decimals cannot carry a decimal ValueType and must be enclosed
                    # in quotes
                    # $Return = '"' + $DecimalValue.ToString( [cultureinfo]::new( 'en-US' )) + '"'
                    $DecimalValue = $DecimalValue.ToString( [cultureinfo]::new('en-US' ))
                    If ( $This.isMetaConfig ) {
                        $DecimalValue = '"' + $DecimalValue + '"'
                    }
                    $Return += $DecimalValue
                    break
                }
                'bool' {
                    $BoolValue = $This.ValueData.ToString().ToLower()
                    If ( $This.isMetaConfig ) {
                        $BoolValue = '"' + $BoolValue + '"'
                    }
                    $Return += $BoolValue
                    break
                }
                'string' {
                    # need to escape \ and " for semi-JSON used in .items
                    $Return += '"' + $This.ValueData.Replace( '\', '\\' ).Replace( '"', '\"' ) + '"'
                    break
                }
                default {
                    $Return += $This.ValueData.ToString()
                }
            }
            Return $Return
        }
    
    }

    class Configuration {
        [Collections.ArrayList] $Items = [Collections.ArrayList]::new()
        
        [String] ToString(){
            Return $This.ToStringInternal( 0, $false )
        }
        [String] ToSTring( [int] $Indent ) {
            Return $This.ToStringInternal( $Indent, $false )
        }
        [String] ToSTring( [int] $Indent, [Bool] $SingleLine ) {
            Return $This.ToStringInternal( $Indent, $SingleLine )
        }

        [String] Hidden ToStringInternal( [int] $Indent, [bool] $SingleLine ) {
            If ( $SingleLine ) {
                [string] $Return = '[ '
                Foreach ( $Item in $This.Items ) {
                    $Return += $Item.ToString() + ', ' 
                }
                $Return = $Return.Substring( 0, $Return.Length - 2 ) + ' ]'
            } Else {
                $NewLine = "`r`n"
                $Spacing = ' ' * $Indent
                [string] $Return = '[' + $NewLine
                Foreach ( $Item in $This.Items ) {
                    $Return += $Spacing + '  ' + $Item.ToString( ) + ',' + $NewLine
                }
                $Return = $Return.Substring( 0, $Return.Length - ( 1 + $NewLine.Length ) ) + $NewLine
                $Return += $Spacing + ']'
            }
            Return $Return
        }
    }

    class ItemConfiguration : Configuration {

        [String] Hidden ToStringInternal( [int] $Indent, [bool] $SingleLine ) {
            $NewLine = "`r`n"
            $Spacing = ' ' * $Indent
            [string] $Return = '{' + $NewLine
            Foreach ( $Item in $This.Items ) {
                $Return += $Spacing + $Item.ToString( $Indent + 2 ) + ',' + $NewLine
            }
            $Return = $Return.Substring( 0, $Return.Length - ( 1 + $NewLine.Length ) ) + $NewLine
            $Return += $Spacing + '}'
            Return $Return
        }
    }

}


process {

    function Convert-ConfigurationFromJSON {
        param (
            [Object] $ConfigurationJSON,
            [Bool] $isMetaConfig = $false
        )
        $Configurations = [Collections.ArrayList]::new()
        Foreach ( $Config in $ConfigurationJSON | Get-Member -MemberType NoteProperty ) {
            If ( $Config.Definition -match "^(?<ValueType>\w+)\s+$( $Config.Name )=(?<ValueData>.+)$" ) {
                Write-Verbose "Processing configuration: $( $Config.Definition )"
                $ConfigValue = [Config]::new( $Matches.ValueType, $Config.Name, $Matches.ValueData, $isMetaConfig )
                [void] $Configurations.Add( $ConfigValue )
            }
        }
        Return ,$Configurations
    }

    Function Get-Things {
        [CmdletBinding()]
        param (
            [Object] $ThingsJSON,
            [String] $Filter
        )

        $Things = [Collections.ArrayList]::new()
        $Bridges = [Collections.ArrayList]::new()

        Foreach ( $Property in $ThingsJSON | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $Filter } ) {

            $JSON = $ThingsJSON."$( $Property.Name )"
        
            # make sure it is a bridge
        
            If ( $JSON.value.isBridge ) {
        
                Write-Verbose "Processing bridge: $( $JSON.value.UID )"

                # basic bridge data
        
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

        Foreach ( $Property in $ThingsJSON | Get-Member -MemberType NoteProperty | Where-Object { $Bridges -notcontains $_.Name } | Where-Object { $_.Name -match $Filter } ) {

            $JSON = $ThingsJSON."$( $Property.Name )"
        
            Write-Verbose "Processing thing: $( $JSON.value.UID )"

            # basic thing data
        
            $Thing = [Thing]::new()
            $Thing.label = $JSON.value.label
            $Thing.location = $JSON.value.location
            $Thing.BindingID = $JSON.value.UID.Split( ':', 4 )[0]
            $Thing.TypeID = $JSON.value.UID.Split( ':', 4 )[1]
            If ( $JSON.value.BridgeUID ) {
                # if the thing uses a bridge, the bridge ID will be part of its UID which thus contains 4 segments...
                $Thing.BridgeID = $JSON.value.UID.Split( ':', 4 )[2]
                $Thing.ThingID = $JSON.value.UID.Split( ':', 4 )[3]
            } Else {
                # ...if it is a standalone thing, its UID will only contain 3 segments and we have no BridgeID
                $Thing.ThingID = $JSON.value.UID.Split( ':', 4 )[2]
            }
        
            $Thing.Location = $JSON.value.location
            $Thing.Configuration.Items = Convert-ConfigurationFromJSON -ConfigurationJSON $JSON.value.Configuration
       
            Foreach ( $Ch in $JSON.value.channels ) {
                Write-Verbose "Processing thing channel: $( $Ch.UID )"
                $Channel = [Channel]::new()
                $Channel.Name = $Ch.Label
                $Channel.Kind = $Ch.Kind
                If ( $Ch.Kind -eq 'TRIGGER' ) {
                    # trigger channels do not define their type because it must be 'String'
                    $Channel.Type = 'String'
                } Else {
                    $Channel.Type = $Ch.itemType
                }
                If ( $ch.uid -match ':(?<ID>[^:]+)$' ) {
                    # channel ID needs to be extracted from channel UID - always last segment (after last ':')
                    $Channel.ID = $Matches.ID
                }
                $Channel.Configuration.Items = Convert-ConfigurationFromJSON -ConfigurationJSON $Ch.Configuration
       
                # only add the channel if any configurations were found
                # all standard channels (without configuration) will be added anyway by the thing binding automatically
                If ( $Channel.Configuration.Count -gt 0 ) {
                    [void] $Thing.Channels.Add( $Channel )
                }
            }

            If ( $JSON.value.BridgeUID ) {
                # this thing uses a bridge, so assign it to its bridge object
                $BindingID = $JSON.value.BridgeUID.Split( ':', 3 )[0]
                $BridgeType = $JSON.value.BridgeUID.Split( ':', 3 )[1]
                $BridgeID = $JSON.value.BridgeUID.Split( ':', 3 )[2]
                $Bridge = $Things | Where-Object { $_.BindingID -eq $BindingID -and $_.BridgeType -eq $BridgeType -and $_.BridgeID -eq $BridgeID }
                [void] $Bridge.Things.Add( $Thing )
            } Else {
                [Void] $Things.Add( $Thing )
            }
        }
        
        Return ,$Things
    }

    function Get-Bindings {
        [CmdletBinding()]
        param (
            [Object] $BindingsJSON,
            [String] $Filter
        )

        $Bindings = [Collections.ArrayList]::new()
        Foreach ( $Property in $BindingsJSON | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $Filter } ) {

            $JSON = $BindingsJSON."$( $Property.Name )"
            Write-Verbose "Processing binding: $( $JSON.value.ChannelUID.UID )"

            $Binding = [Binding]::new()
            $Binding.name = $Property.Name
            $Binding.uid = $JSON.value.ChannelUID.UID
            $Binding.itemName = $JSON.value.itemName
            $Binding.Configuration.Items = Convert-ConfigurationFromJSON -ConfigurationJSON $JSON.value.Configuration

            [void] $Bindings.Add( $Binding )
        }
        Return ,$Bindings
    }

    function Get-Metadata {
        [CmdletBinding()]
        param (
            [Object] $MetadataJSON,
            [String] $Filter
        )

        $Metadatas = [Collections.Arraylist]::new()
        Foreach ( $Property in $MetadataJSON | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $Filter } ) {
    
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
        Return ,$Metadatas
    }

    function Get-Items {
        [CmdletBinding()]
        param (
            [Object] $ItemsJSON,
            [Object] $BindingsJSON,
            [Object] $MetadataJSON,
            [String] $Filter
        )

        $Bindings = Get-Bindings -BindingsJSON $BindingsJSON -Filter $Filter
        $Metadata = Get-Metadata -MetadataJSON $MetadataJSON -Filter $Filter
        
        $Items = [Collections.ArrayList]::new()
        Foreach ( $Property in $ItemsJSON | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $Filter } ) {
    
            $JSON = $ItemsJSON."$( $Property.Name )"
            Write-Verbose "Processing item: $( $Property.Name )"
        
            $Item = [Item]::new()
            $Item.Name = $Property.Name
            $Item.itemType = $JSON.value.itemType
            $Item.label = $JSON.value.label
            $Item.iconname = $JSON.value.category
            $Item.baseItemType = $JSON.value.baseItemType
            $Item.functionName = $JSON.value.functionName
        
            Foreach ( $functionParam in $JSON.value.functionParams ) {
                Write-Verbose "Processing item function param: $( $functionParam )"
                [void] $Item.functionParams.Add( $functionParam )
            }
            Foreach ( $Group in $JSON.value.groupNames ) {
                Write-Verbose "Processing item group: $( $group )"
                [void] $Item.groups.Add( $Group )
            }
            Foreach ( $Tag in $JSON.value.tags ) {
                Write-Verbose "Processing item tag: $( $tag )"
                [void] $Item.tags.Add( $Tag )
            }
            Foreach ( $Binding in $Bindings | Where-Object { $_.itemName -eq $Item.Name } ) {
                Write-Verbose "Processing item binding: $( $Binding )"
                [void] $Item.configuration.Items.Add( $Binding )
            }
            Foreach ( $Meta in $Metadata | Where-Object { $_.itemName -eq $Item.Name } ) {
                Write-Verbose "Processing item metadata: $( $Meta )"
                [void] $Item.configuration.Items.Add( $Meta )
            }

            [void] $Items.Add( $Item )
        }
        Return ,$Items
    }
    
    If ( -not $JSONFolder ) {
        $JSONFolder = $PSScriptRoot
    }

    If ( $CreateThings ) {
        $ThingsJSON = Get-Content "$JSONFolder\org.openhab.core.thing.Thing.JSON" | ConvertFrom-Json
        If ( $OutFileBasename ) {
            $OutFile = "$JSONFolder\$OutfileBasename.things"
        } Else {
            $OutFile = "$JSONFolder\allthings.things"
        }
        $Things = Get-Things -ThingsJSON $ThingsJSON -Filter $Filter
        $streamWriter = [IO.StreamWriter]::new( $Outfile, $false, $Encoding )
        Foreach ( $Thing in $Things | Sort-Object -Property Class, BindingID ) {
            $streamwriter.Write( $Thing.ToString() )
        }
        $streamWriter.Dispose()
    }
    If ( $CreateItems ) {
        $ItemsJSON = Get-Content "$JSONFolder\org.openhab.core.items.Item.JSON" | ConvertFrom-Json
        $BindingsJSON = Get-Content "$JSONFolder\org.openhab.core.thing.Link.ItemChannelLink.JSON" | ConvertFrom-Json
        $MetadataJSON = Get-Content "$JSONFolder\org.openhab.core.items.Metadata.JSON" | ConvertFrom-Json
        If ( $OutFileBasename ) {
            $OutFile = "$JSONFolder\$OutfileBasename.items"
        } Else {
            $OutFile = "$JSONFolder\allitems.items"
        }
        $Items = Get-Items -ItemsJSON $ItemsJSON -BindingsJSON $BindingsJSON -MetadataJSON $MetadataJSON -Filter $Filter
        $streamWriter = [IO.StreamWriter]::new( $Outfile, $false, $Encoding )
        ForEach ( $Item in $Items | Sort-Object -Property ItemType, Name ) {
            $streamWriter.WriteLine( $Item.ToString() )
        }
        $streamWriter.Dispose()
    }

    
}

end {

}
