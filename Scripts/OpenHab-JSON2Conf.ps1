[CmdletBinding()]
param (
    [ValidateScript( { Test-Path $_ } )]
    [String[]] $JSONFolder,
    [Switch] $Things,
    [Switch] $Items,
    [String] $OutFileBasename,
    [String] $Filter = '.*'
)

begin {

    $Encoding = [Text.Encoding]::GetEncoding( 1252 )

    # generic bridge definition:
    # Bridge <binding_name>:<bridge_type>:<bridge_name> [ <parameters> ] {
    #   (array of things)
    # }
    # generic bridge thing definition
    # Thing <type_id> <thing_id> "Label" @ "Location" [ <parameters> ]
    # generic standalone thing definition
    # Thing <binding_id>:<type_id>:<thing_id> "Label" @ "Location" [ <parameters> ]

    class Bridge {

        # basic bridge properties
        [String] $BindingID
        [String] $BridgeType
        [String] $BridgeID
        [String] $label
        [String] $location
        [Collections.ArrayList] $Configuration = [Collections.ArrayList]::new()
        [Collections.ArrayList] $Things = [Collections.ArrayList]::new()
        
        [String] ToString() {
            Return $This.CreateItem()
        }

        [String] Hidden CreateItem() {

            # create the .things bridge definition from the bridge properties

            [String] $Return = $This.GetType().Name + ' ' + $This.BindingID + ':' + $This.BridgeType + ':' + $This.BridgeID
            If ( $This.label ) { $Return += ' "' + $This.label + '"' }
            If ( $This.location ) { $Return += ' @ "' + $This.location + '"' }

            # if the bridge has configuration values, insert them in square brackets
            # and take care of indentation - Openhab is quite picky about misalignment :-)

            If ( $This.Configuration.Count -gt 0 ) {
                $Return += " [`r`n"
                Foreach ( $Config in $This.Configuration ) {
                    $Return += '  ' + $Config.ValueName + '=' + $Config.ToString() + ",`r`n"
                }
                $Return = $Return.Substring( 0, $Return.Length - 3 ) + "`r`n]" # # strip last comma, close section
            }

            # if there are things using this bridge, include them in the bridge definition within curly brackets

            If ( $This.Things.Count -gt 0 ) {
                $Return += " {`r`n"
                Foreach ( $Thing in $This.Things ) {
                    # ToString( $true ) means "this thing is a child of a bridge" - this adds indendation
                    # for pretty formatting and adjusts the thing definition for bridge bound things according
                    # to the description above
                    $Return += $Thing.ToString( $True )
                }
                $Return += "}`r`n"
            }
            Return $Return
        }
    }

    Class Thing {

        # basic thing properties
        [String] $BindingID
        [String] $TypeID
        [String] $BridgeID
        [String] $ThingID
        [String] $label
        [String] $Location
        [Collections.ArrayList] $Configuration = [Collections.ArrayList]::new()
        [Collections.Arraylist] $Channels = [Collections.ArrayList]::new()
        
        [String] ToString( ){
            Return $This.CreateItem( $False )
        }
        [String] ToString( $IsBridgeChild ) {
            Return $This.CreateItem( $IsBridgeChild )
        }

        [String] Hidden CreateItem( $IsBridgeChild ) {

            # create the .things item definition from the item properties

            # if the item is a child of a bridge, we want 2 spaces more at the beginning of each line
            # and we have a different string composition

            If ( $IsBridgeChild ) { 
                $Indent = 2 
                [String] $Return = ' ' * $Indent + $This.GetType().Name + ' ' + $This.TypeID + ' ' + $This.ThingID
            } Else { 
                $Indent = 0 
                [String] $Return = ' ' * $Indent + $This.GetType().Name + ' ' + $This.BindingID + ':' + $This.TypeID + ':' + $This.ThingID
            }
            If ( $This.label ) { $Return += ' "' + $This.label + '"' }
            If ( $This.location ) { $Return += ' @ "' + $This.location + '"' }

            # if the thing has configuration values, add them in square brackets
    
            If ( $This.Configuration.Count -gt 0 ) {
                $Return += ' [ '
                Foreach ( $Config in $This.Configuration ) {
                    $Return += $Config.ValueName + '=' + $Config.ToString() + ', '
                }
                $Return = $Return.Substring( 0, $Return.Length - 2 ) + ' ]' # strip last comma, close section
            }
            
            # if the thing has channels, include them in curly brackets as well
            # again, take care of correct indendation
            
            If ( $This.Channels.Count -gt 0 ) {
                $Return += " {`r`n"
                $Return += ' ' * $Indent + "  Channels:`r`n"
                Foreach ( $Channel in $This.Channels ) {
                    $Return += $Channel.ToString( $Indent + 4 )
                }
                $Return += ' ' * $Indent + '}'
            }
            $Return += "`r`n"
            Return $Return 
        }

    }

    Class Channel {

        # channel definition in .item files as documented, see above
        [String] $Kind
        [String] $Type
        [String] $ID
        [String] $Name
        [Collections.ArrayList] $Configuration = [Collections.ArrayList]::new()
        
        [String] ToString( ){
            Return $This.CreateItem( 4 )
        }

        [String] ToString( [int] $Indent ) {
            Return $This.CreateItem( $Indent )
        }

        [String] Hidden CreateItem( [int] $Indent = 4 ){

            [String] $Return = ''

            # create the .things channel definition from the channel properties
            # if the channel has configuration values, append them in square brackets
            
            If ( $This.Configuration.Count -gt 0 ) {
                $Return += ' ' * $Indent + $This.Kind.Substring( 0, 1 ).ToUpper() + $This.Kind.Substring( 1 ).ToLower() + ' ' + $This.Type + ' : ' + $This.ID
                If ( $This.Name ) {
                    $Return += ' "' + $This.Name + '"'
                }
                
                $Return += ' [ '
                Foreach ( $Config in $This.Configuration ) {
                    $Return += $Config.ValueName + '=' + $Config.ToString() + ', '
                }
                $Return = $Return.Substring( 0, $Return.Length - 2 ) + " ]`r`n" # strip last comma, close section
            }

            Return $Return
        }

    }

    class Item {
    
        # basic item properties
        [String] $itemType
        [String] $Name
        [String] $label
        [String] $category
        [String] $iconName
        [Collections.ArrayList] $Bindings = [Collections.ArrayList]::new()
        [Collections.ArrayList] $groups = [Collections.ArrayList]::new()
        [Collections.ArrayList] $tags = [Collections.ArrayList]::new()
        [Collections.ArrayList] $metadata = [Collections.ArrayList]::new()
    
        # required for aggregate groups
        [string] $baseItemType 
        [string] $functionName
        [Collections.ArrayList] $functionParams = [Collections.ArrayList]::new()
    
        [String] ToString( ) {
            Return $This.CreateItem()
        }
    
        [String] Hidden CreateItem() {
    
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
                $Return = $Return.Substring( 0, $Return.Length - 2 ) + " )" # strip last comma, close section
            }
            If ( $This.tags.Count -ge 1 ) {
                $Return += ' [ '
                Foreach ( $Tag in $This.tags ) {
                    $Return += '"' + $Tag + '", '
                }
                $Return = $Return.Substring( 0, $Return.Length - 2 ) + " ]" # strip last comma, close section
            }
    
            If ( $This.Bindings.Count + $This.metadata.Count -gt 0 ) {
                # both bindings and metadata go into the same section, so we need to handle them together
                $Return += " {`r`n"
                If ( $This.Bindings.Count -ge 1 ) {
                    Foreach ( $Binding in $This.Bindings ) {
                        $Return += '  ' + $Binding.ToString() + ",`r`n"
                    }
                }
                If ( $This.metadata.Count -ge 1 ) {
                    Foreach ( $Meta in $This.metadata ) {
                        $Return += '  ' + $Meta.ToString() + ",`r`n"
                    }
                }
                $Return = $Return.Substring( 0, $Return.Length - 3 ) + "`r`n}" # strip last comma, close section
            }
    
            Return $Return
        }
    }
    
    class Binding {
    
        # basic binding properties
        [String] $name
        [String] $uid
        [String] $itemName
        [Collections.ArrayList] $Configuration = [Collections.ArrayList]::new()
    
        [String] ToString() {
            Return $This.CreateBinding()
        }
    
        [String] Hidden CreateBinding() {
    
            # binding definition string in .items files as documented, see above
            [String] $Return = 'channel="' + $This.uid + '"'
    
            If ( $This.Configuration.Count -gt 0 ) {
                $Return += ' [ '
                Foreach ( $Config in $This.Configuration ) {
                    $Return += $Config.ValueName + '=' + $Config.ToString() + ', '
                }
                $Return = $Return.Substring( 0, $Return.Length - 2 ) + ' ]' # strip last comma, close section
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
        [Collections.ArrayList] $Configuration = [Collections.ArrayList]::new()
    
        [String] ToString() {
            Return $This.CreateMeta()
        }
    
        [String] Hidden CreateMeta () {
            [String] $Return = $This.type + '="' + $This.value + '"'
            If ( $This.Configuration.Count -gt 0 ) {
                # if it has configurations, create a section and insert all of them
                $Return += ' [ '
                Foreach ( $Config in $This.Configuration ) {
                    $Return += $Config.ValueName + '=' + $Config.ToString() + ', '
                }
                $Return = $Return.Substring( 0, $Return.Length - 2 ) + ' ]' # strip last comma, close section
            }
            Return $Return
        }
    
    }
    
    Class Config {

        # items, bindings, etc. might have config values. These consist of a name, a value and (optionally) a type
        # to make things easier, this class provides them and handles their types
        [String] $ValueType
        [String] $ValueName
        [String] $ValueData
        # meta configuration behaves weird in terms of data types... bool and decimal must be enclosed in quotes.
        [Bool] $isMetaConfig
    
        Config () {}
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
            $This.isMeteConfig = $false
        }
        Config ( [String] $ValueType, [String] $ValueName, [String] $ValueData ) {
            $This.ValueType = $ValueType
            $This.Valuename = $ValueName
            $This.ValueData = $ValueData
            $This.isMeteConfig = $false
        }
    
        Config ( [String] $ValueType, [String] $ValueName, [String] $ValueData, [Bool] $isMetaConfig ) {
            $This.ValueType = $ValueType
            $This.Valuename = $ValueName
            $This.ValueData = $ValueData
            $This.isMetaConfig = $isMetaConfig
        }
    
        [String] ToString() {
            Return $This.ValueToString( $This.ValueData, $This.ValueType, $This.isMetaConfig )
        }
        [String] ToString( [String] $ValueData, [String] $ValueType ) {
            Return $This.ValueToString( $ValueData, $ValueType, $This.isMetaConfig )
        }
        [String] ToString( [String] $ValueData, [String] $ValueType, [Bool] $isMetaConfig ) {
            Return $This.ValueToString( $ValueData, $ValueType, $isMetaConfig )
        }
    
        [String] Hidden ValueToString ( [String] $ValueData, [String] $ValueType, [bool] $isMetaConfig ) {
            [String] $Return = ''
            Switch ( $ValueType ) {
                'int' {
                    $Return = $ValueData
                }
                'decimal' {
                    # for decimals, we need dot separated values in the item file. This depends on the current locale,
                    # we need to force the decimal to be converted first to a single float depending on the current separator
                    # that ConvertFrom-JSON insert, and then to the en-US string format where the separator is a dot.
                    $DecimalValue = $ValueData
                    If ( $DecimalValue -match ',') {
                        $DecimalValue = $DecimalValue.ToSingle( [cultureinfo]::new( 'de-DE' ))
                    } Else {
                        $DecimalValue = $DecimalValue.ToSingle( [cultureinfo]::new( 'en-US' ))
                    }
                    # for item configurations, decimals cannot carry a decimal ValueType and must be enclosed
                    # in quotes
                    # $Return = '"' + $DecimalValue.ToString( [cultureinfo]::new( 'en-US' )) + '"'
                    $Return = $DecimalValue.ToString( [cultureinfo]::new('en-US' ))
                    If ( $isMetaConfig ) {
                        $Return = '"' + $Return + '"'
                    }
                }
                'bool' {
                    $Return = $ValueData.ToString().ToLower()
                    If ( $isMetaConfig ) {
                        $Return = '"' + $Return + '"'
                    }
                }
                'string' {
                    # need to escape \ and " for semi-JSON used in .items
                    $Return = '"' + $ValueData.Replace( '\', '\\' ).Replace( '"', '\"' ) + '"'
                }
                default {
                    $Return = $ValueData.ToString()
                }
            }
            Return $Return
        }
    
    }

}

process {

    function Get-Configuration {
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
                $Bridge.Configuration = Get-Configuration -ConfigurationJSON $JSON.value.Configuration

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
            $Thing.Configuration = Get-Configuration -ConfigurationJSON $JSON.value.Configuration
       
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
                $Channel.Configuration = Get-Configuration -ConfigurationJSON $Ch.Configuration
       
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
            $Binding.Configuration = Get-Configuration -ConfigurationJSON $JSON.value.Configuration

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
            $Metadata.Configuration = Get-Configuration -ConfigurationJSON $JSON.value.configuration
        
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
                Write-Verbose "Processing item binding: $( $ItemBinding )"
                [void] $Item.Bindings.Add( $Binding )
            }
            Foreach ( $Meta in $Metadata | Where-Object { $_.itemName -eq $Item.Name } ) {
                Write-Verbose "Processing item metadata: $( $Meta )"
                [void] $Item.metadata.Add( $Meta )
            }

            [void] $Items.Add( $Item )
        }
        Return ,$Items
    }
    
    If ( -not $JSONFolder ) {
        $JSONFolder = $PSScriptRoot
    }

    If ( $Things ) {
        $ThingsJSON = Get-Content "$JSONFolder\org.openhab.core.thing.Thing.JSON" | ConvertFrom-Json
        If ( $OutFileBasename ) {
            $OutFile = "$JSONFolder\$OutfileBasename.things"
        } Else {
            $OutFile = "$JSONFolder\allthings.things"
        }
        $Things = Get-Things -ThingsJSON $ThingsJSON -Filter $Filter
        $streamWriter = [IO.StreamWriter]::new( $Outfile, $false, $Encoding )
        Foreach ( $Thing in $Things | Sort-Object -Property BindingID ) {
            $streamwriter.Write( $Thing.ToString() )
        }
        $streamWriter.Dispose()
    }
    If ( $Items ) {
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
