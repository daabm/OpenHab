$ItemsRaw = Get-Content $PSScriptRoot\org.openhab.core.items.Item.JSON | ConvertFrom-Json
$BindingsRaw = Get-Content $PSScriptRoot\org.openhab.core.thing.link.ItemChannelLink.json | ConvertFrom-Json
$MetadataRaw = Get-Content $PSScriptRoot\org.openhab.core.items.Metadata.json | ConvertFrom-Json
$OutFile = "$PSScriptRoot\allitems.items"
$ItemsFilter = '.*'
#$ItemsFilter = 'TSOG2Ku'
#$ItemsFilter = 'SomfyTemperatur'

$Items = [Collections.ArrayList]::new()
$Bindings = [Collections.ArrayList]::new()
$Metadatas = [Collections.ArrayList]::new()

# generic item definition:
# itemtype itemname "labeltext [stateformat]" <iconname> (group1, group2, ...) ["tag1", "tag2", ...] {bindingconfig}
# generic group definition:
# Group groupname ["labeltext"] [<iconname>] [(group1, group2, ...)]
# group with aggregate function:
# Group[:itemtype[:function]] groupname ["labeltext"] [<iconname>] [(group1, group2, ...)]

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
                        # COUNT has a single property over which it aggregates
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
    [Collections.ArrayList] $Properties = [Collections.ArrayList]::new()

    [String] ToString() {
        Return $This.CreateBinding()
    }

    [String] Hidden CreateBinding() {

        # binding definition string in .items files as documented, see above
        [String] $Return = 'channel="' + $This.uid + '"'

        If ( $This.Properties.Count -gt 0 ) {
            $Return += '[ '
            Foreach ( $Property in $This.Properties ) {
                $Return += $Property.ValueName + '=' + $Property.ToString() + ', '
            }
            $Return = $Return.Substring( 0, $Return.Length - 2 ) + ']' # strip last comma, close section
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
    }
    Config ( [String] $ValueType, [String] $ValueName, [String] $ValueData ) {
        $This.ValueType = $ValueType
        $This.Valuename = $ValueName
        $This.ValueData = $ValueData
    }

    [String] ToString() {
        Return $THis.ValueToString( $This.ValueData, $This.ValueType )
    }
    [String] ToString( $ValueData, $ValueType ) {
        Return $This.ValueToString( $ValueData, $ValueType )
    }

    [String] Hidden ValueToString ( [String] $ValueData, [String] $ValueType ) {
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
            }
            'bool' {
                $Return = $ValueData.ToString().ToLower()
            }
            'string' {
                $Return = '"' + $ValueData.Replace( '"', '\"' ) + '"'
            }
        }
        Return $Return
    }

}

# create a list of all bindings to assign them to the items later

Foreach ( $Property in $BindingsRaw | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $ItemsFilter } ) {

    $JSON = $BindingsRaw."$( $Property.Name )"

    $Binding = [Binding]::new()
    $Binding.name = $Property.Name
    $Binding.uid = $JSON.value.ChannelUID.UID
    $Binding.itemName = $JSON.value.itemName

    Write-Verbose "Processing binding: $( $Binding.UID )"

    Foreach ( $Configuration in $JSON.value.configuration ) {
        Foreach ( $Config in $Configuration.Properties | Get-Member -MemberType NoteProperty ) {
            If ( $Config.Definition -match "^(?<ValueType>\w+)\s+$( $Config.Name )=(?<ValueData>.+)$" ) {
                Write-Verbose "Processing binding configuration: $( $Config.Definition )"
                $Property = [Config]::new( $Matches.ValueType, $Config.Name, $Configuration.Properties."$( $Config.Name )".Replace( '\', '\\' ))
                [void] $Binding.Properties.Add( $Property )
            }
        }
    }
    [void] $Bindings.Add( $Binding )
}

# create a list of all metadata to assing them to the items later

Foreach ( $Property in $MetadataRaw | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $ItemsFilter } ) {

    $JSON = $MetadataRaw."$( $Property.Name )"

    $MetaData = [Metadata]::new()
    $MetaData.name = $Property.Name
    $Metadata.type = $Property.Name.Split( ':' )[0]
    $Metadata.itemName = $Property.Name.Split( ':', 2 )[1]
    $MetaData.value = $JSON.value.value

    Write-Verbose "Processing metadata: $( $Metadata.Name )"
    
    Foreach ( $MetaConfig in $JSON.value.configuration | Get-Member -MemberType NoteProperty  ) {
        Write-Verbose "Processing metadata configuration: $( $Metaconfig.Definition )"
        $Config = [Config]::new( $MetaConfig.Name, $JSON.value.configuration."$( $MetaConfig.Name )" )
        [void] $Metadata.Configuration.Add( $Config )
    }
    [void] $MetaDatas.Add( $MetaData )
}

# analyze all items

Foreach ( $Property in $ItemsRaw | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $ItemsFilter } ) {

    $JSON = $ItemsRaw."$( $Property.Name )"

    # basic item data

    $Item = [Item]::new()
    $Item.Name = $Property.Name
    $Item.itemType = $JSON.value.itemType # .Split( ':', 2 )[0]
    $Item.label = $JSON.value.label
    $Item.iconname = $JSON.value.category
    $Item.baseItemType = $JSON.value.baseItemType
    $Item.functionName = $JSON.value.functionName

    Write-Verbose "Processing item: $( $Item.Name )"

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

    Foreach ( $ItemBinding in $Bindings | Where-Object { $_.itemName -eq $Item.Name } ) {
        Write-Verbose "Processing item binding: $( $ItemBinding )"
        [void] $Item.Bindings.Add( $ItemBinding )
    }

    Foreach ( $Meta in $Metadatas | Where-Object { $_.itemName -eq $Item.Name } ) {
        Write-Verbose "Processing item metadata: $( $Meta )"
        [void] $Item.metadata.Add( $Meta )
    }

    [void] $Items.Add( $Item )
}

$encoding = [Text.Encoding]::GetEncoding( 1252 )
$streamWriter = [IO.StreamWriter]::new( $Outfile, $false, $Encoding )
$Items | Sort-Object -Property itemType, Name | Foreach-Object { $_.ToString() | ForEach-Object { $streamWriter.WriteLine( $_ ) } }
$streamWriter.Dispose()
