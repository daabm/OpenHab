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

# define properties for items
# generic item definition:
# itemtype itemname "labeltext [stateformat]" <iconname> (group1, group2, ...) ["tag1", "tag2", ...] {bindingconfig}
# generic group definition:
# Group groupname ["labeltext"] [<iconname>] [(group1, group2, ...)]
# group with aggregate function:
# Group[:itemtype[:function]] groupname ["labeltext"] [<iconname>] [(group1, group2, ...)]

class Item {
    
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

        # item definition in .items files as documented

        [String] $Return = $This.itemType

        # handle aggregate groups - only these have a baseItemType and optionally an aggregate function
        If ( $This.baseItemType ) {
            $Return += ':' + $This.baseItemType
            If ( $This.functionName ) {
                $Return += ':' + $This.functionName
                If ( $This.functionParams ) {
                    If ( $This.functionName -eq 'COUNT' ){
                        $Return += '"' + $This.functionParams + '"'
                    } Else {
                        $Return += '('
                        Foreach ( $functionParam in $This.functionParams ) {
                            $Return += $FunctionParam + ','
                        }
                        $Return = $Return.Substring( 0, $Return.Length - 1 ) + ')'
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
            $Return = $Return.Substring( 0, $Return.Length - 2 ) + " )"
        }
        If ( $This.tags.Count -ge 1 ) {
            $Return += ' [ '
            Foreach ( $Tag in $This.tags ) {
                $Return += '"' + $Tag + '", '
            }
            $Return = $Return.Substring( 0, $Return.Length - 2 ) + " ]"
        }

        If ( $This.Bindings.Count + $This.metadata.Count -gt 0 ) {
            $Return += " { "
            If ( $This.Bindings.Count -ge 1 ) {
                Foreach ( $Binding in $This.Bindings ) {
                    $Return += $Binding.ToString() + ', '
                }
            }
            If ( $This.metadata.Count -ge 1 ) {
                Foreach ( $Meta in $This.metadata ) {
                    $Return += $Meta.ToString() + ', '
                }
            }
            $Return = $Return.Substring( 0, $Return.Length - 2 ) + ' }'
        }

        Return $Return
    }
}

class Binding {
    
    [String] $name
    [String] $uid
    [String] $itemName
    [Collections.ArrayList] $Properties = [Collections.ArrayList]::new()

    [String] ToString() {
        Return $This.CreateBinding()
    }

    [String] Hidden CreateBinding() {
        [String] $Return = 'channel="' + $This.uid + '"'
        If ( $This.Properties.Count -gt 0 ) {
            $Return += '[ '
            Foreach ( $Property in $This.Properties ) {
                $Return += $Property.ValueName + '=' + $Property.ToString() + ', '
            }
            $Return = $Return.Substring( 0, $Return.Length - 2 ) + ']'
        }
        Return $Return
    }

}

Class Metadata {
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
            $Return += ' [ '
            Foreach ( $Config in $This.Configuration ) {
                $Return += $Config.ValueName + '=' + $Config.ToString() + ', '
            }
            $Return = $Return.Substring( 0, $Return.Length - 2 ) + ' ]'
        }
        Return $Return
    }

}

# create binding list
Class Config {
    [String] $ValueType
    [String] $ValueName
    [String] $ValueData

    Config () {}
    Config ( [String] $ValueName, [String] $ValueData ) {
        $This.Valuename = $ValueName
        $This.ValueData = $ValueData
        Switch -regex ( $This.ValueData ) {
            '^(true|false)$' {
                $This.ValueType = 'string'
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
                $DecimalValue = $ValueData
                If ( $DecimalValue -match ',') {
                    $DecimalValue = $DecimalValue.ToSingle( [cultureinfo]::new( 'de-DE' ))
                } Else {
                    $DecimalValue = $DecimalValue.ToSingle( [cultureinfo]::new( 'en-US' ))
                }
                $Return = '"' + $DecimalValue.ToString( [cultureinfo]::new( 'en-US' ) ) + '"'
            }
            'bool' {
                $Return = $ValueData.ToString().ToLower()
            }
            'string' {
                $Return = '"' + $ValueData + '"'
            }
        }
        Return $Return
    }

}

Foreach ( $Property in $BindingsRaw | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $ItemsFilter } ) {

    $JSON = $BindingsRaw."$( $Property.Name )"
    $Binding = [Binding]::new()
    $Binding.name = $Property.Name
    $Binding.uid = $JSON.value.ChannelUID.UID
    $Binding.itemName = $JSON.value.itemName
    Foreach ( $Configuration in $JSON.value.configuration ) {
        Foreach ( $Config in $Configuration.Properties | Get-Member -MemberType NoteProperty ) {
            If ( $Config.Definition -match "^(?<ValueType>\w+)\s+$( $Config.Name )=(?<ValueData>.+)$" ) {
                $Property = [Config]::new( $Matches.ValueType, $Config.Name, $Configuration.Properties."$( $Config.Name )".Replace( '\', '\\' ))
                [void] $Binding.Properties.Add( $Property )
            }
        }
    }
    [void] $Bindings.Add( $Binding )
}

# create metadata list

Foreach ( $Property in $MetadataRaw | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $ItemsFilter } ) {
    $JSON = $MetadataRaw."$( $Property.Name )"
    $MetaData = [Metadata]::new()
    $MetaData.name = $Property.Name
    $Metadata.type = $Property.Name.Split( ':' )[0]
    $Metadata.itemName = $Property.Name.Split( ':', 2 )[1]
    $MetaData.value = $JSON.value.value
    Foreach ( $MetaConfig in $JSON.value.configuration | Get-Member -MemberType NoteProperty  ) {
        $Config = [Config]::new( $MetaConfig.Name, $JSON.value.configuration."$( $MetaConfig.Name )" )
        [void] $Metadata.Configuration.Add( $Config )
    }
    [void] $MetaDatas.Add( $MetaData )
}

# analyze all items

Foreach ( $Property in $ItemsRaw | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $ItemsFilter } ) {

    $JSON = $ItemsRaw."$( $Property.Name )"

    $Item = [Item]::new()
    $Item.Name = $Property.Name
    $Item.itemType = $JSON.value.itemType # .Split( ':', 2 )[0]
    $Item.label = $JSON.value.label
    $Item.category = $JSON.value.category
    $Item.baseItemType = $JSON.value.baseItemType
    $Item.functionName = $JSON.value.functionName
    Foreach ( $functionParam in $JSON.value.functionParams ) {
        [void] $Item.functionParams.Add( $functionParam )
    }
    Foreach ( $Group in $JSON.value.groupNames ) {
        [void] $Item.groups.Add( $Group )
    }
    Foreach ( $Tag in $JSON.value.tags ) {
        [void] $Item.tags.Add( $Tag )
    }

    Foreach ( $ItemBinding in $Bindings | Where-Object { $_.itemName -eq $Item.Name } ) {
        [void] $Item.Bindings.Add( $ItemBinding )
    }

    Foreach ( $Meta in $Metadatas | Where-Object { $_.itemName -eq $Item.Name } ) {
        [void] $Item.metadata.Add( $Meta )
    }

    [void] $Items.Add( $Item )
}

$encoding = [Text.Encoding]::GetEncoding( 1252 )
$streamWriter = [IO.StreamWriter]::new( $Outfile, $false, $Encoding )
$Items | Sort-Object -Property itemType, Name | Foreach-Object { $_.ToString() | ForEach-Object { $streamWriter.WriteLine( $_ ) } }
#$Lines = ( $Items | Sort-Object -Property itemType, Name | Foreach-Object { $_.ToString() } )
#[IO.File]::WriteAllLines( $OutFile, $lines, [Text.Encoding]::GetEncoding(1252))
#$Lines | Out-File $Outfile -Encoding default -Force
$streamWriter.Dispose()
