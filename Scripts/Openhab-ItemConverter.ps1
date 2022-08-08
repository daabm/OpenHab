# $Things = Get-Content .\org.openhab.core.thing.Thing.json | ConvertFrom-Json
$ItemsRaw = Get-Content $PSScriptRoot\org.openhab.core.items.Item.JSON | ConvertFrom-Json
$BindingsRaw = Get-Content $PSScriptRoot\org.openhab.core.thing.link.ItemChannelLink.json | ConvertFrom-Json
$MetadataRaw = Get-Content $PSScriptRoot\org.openhab.core.items.Metadata.json | ConvertFrom-Json
$ItemsFilter = '.*'
$ItemsFilter = 'TSOG2Ku'

$Processed = [Collections.ArrayList]::new()
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


    [String] CreateOHItem( ) {
        Return $This.CreateItem()
    }

    [String] Hidden CreateItem() {

        # item definition in .items files as documented

        [String] $ItemReturn = $This.itemType

        # handle aggregate groups - only these have a baseItemType and optionally an aggregate function
        If ( $This.baseItemType ) {
            $ItemReturn += ':' + $This.baseItemType
            If ( $This.functionName ) {
                $ItemReturn += ':' + $This.functionName
                If ( $This.functionParams ) {
                    If ( $This.functionName -eq 'COUNT' ){
                        $ItemReturn += '"' + $This.functionParams + '"'
                    } Else {
                        $ItemReturn += '('
                        Foreach ( $functionParam in $This.functionParams ) {
                            $ItemReturn += $FunctionParam + ','
                        }
                        $ItemReturn = $ItemReturn.Substring( 0, $ItemReturn.Length - 1 ) + ')'
                    }

                }
            }
        }
        

        $ItemReturn += ' ' + $This.name + ' "' + $This.label + '"'
        If ( $This.iconName ) {
            $ItemReturn += ' <' + $This.iconName + '>'
        }
        If ( $This.Groups.Count -ge 1 ) {
            $ItemReturn += ' ( '
            Foreach ( $Group in $This.Groups ) {
                $ItemReturn += $Group + ', '
            }
            $ItemReturn = $ItemReturn.Substring( 0, $ItemReturn.Length - 2 ) + " )"
        }
        If ( $This.tags.Count -ge 1 ) {
            $ItemReturn += ' [ '
            Foreach ( $Tag in $This.tags ) {
                $ItemReturn += '"' + $Tag + '", '
            }
            $ItemReturn = $ItemReturn.Substring( 0, $ItemReturn.Length - 2 ) + " ]"
        }

        If ( $This.Bindings.Count + $This.metadata.Count -gt 0 ) {
            $ItemReturn += " { "
            If ( $This.Bindings.Count -ge 1 ) {
                Foreach ( $Binding in $This.Bindings ) {
                    $ItemReturn += $Binding.CreateOHBinding() + ', '
                }
            }
            If ( $This.metadata.Count -ge 1 ) {
                Foreach ( $Meta in $This.metadata ) {
                    $ItemReturn += $Meta.CreateOHMeta() + ', '
                }
            }
            $ItemReturn = $ItemReturn.Substring( 0, $ItemReturn.Length - 2 ) + ' }'
        }

        Return $ItemReturn
    }
}

class Binding {
    
    [String] $name
    [String] $uid
    [String] $itemName
    [String] $profile
    [Collections.HashTable] $profileParameters = [Collections.HashTable]::new()

    [String] CreateOhBinding() {
        Return $This.CreateBinding()
    }

    [String] Hidden CreateBinding() {
        [String] $BindingReturn = 'channel="' + $This.uid + '"'
        If ( $This.profile ) {
            $BindingReturn += ' [ profile="' + $This.profile + '"'
            Foreach ( $Key in $This.profileParameters.Keys ) {
                $BindingReturn += ', ' + $Key + '='
                $rtn = ''
                If ( [double]::TryParse( $This.profileParameters[ $Key ], [ref]$rtn )) { # check if we have a number, otherwise we need surrounding double quotes
                    $BindingReturn += $This.profileParameters[ $Key ]
                } Else {
                    $BindingReturn += '"' + $This.profileParameters[ $Key ] + '"'
                }
            }
            $BindingReturn += ' ]'
        }
        Return $BindingReturn
    }

}

Class Metadata {
    [String] $name
    [String] $type
    [String] $value
    [String] $itemName
    [Collections.HashTable] $Configuration = [Collections.HashTable]::new()

    [String] CreateOHMeta() {
        Return $This.CreateMeta()
    }

    [String] Hidden CreateMeta () {
        [String] $MetaReturn = $This.type + '="' + $This.value + '"'
        If ( $This.Configuration.Count -gt 0 ) {
            $MetaReturn += ' [ '
            Foreach ( $Key in $This.Configuration.Keys ) {
                $MetaReturn += $Key + '='
                $rtn = ''
                If ( [double]::TryParse( $This.Configuration[ $Key ], [ref]$rtn )) { # check if we have a number, otherwise we need surrounding double quotes
                    $MetaReturn += $This.Configuration[ $Key ]
                } Else {
                    $MetaReturn += '"' + $This.Configuration[ $Key ] + '"'
                }
                $MetaReturn += ', '
            }
            $MetaReturn = $MetaReturn.Substring( 0, $MetaReturn.Length - 2 ) + ' ]'
        }
        Return $MetaReturn
    }

}

# create binding list

Foreach ( $Property in $BindingsRaw | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $ItemsFilter } ) {
    $JSON = $BindingsRaw."$( $Property.Name )"
    $Binding = [Binding]::new()
    $Binding.name = $Property.Name
    $Binding.uid = $JSON.value.ChannelUID.UID
    $Binding.itemName = $JSON.value.itemName
    Foreach ( $Configuration in $JSON.value.configuration ) {
        Foreach ( $BindingProperty in $Configuration.Properties | Get-Member -MemberType NoteProperty ) {
            If ( $BindingProperty.Name -eq 'profile' ) {
                $Binding.profile = $Configuration.Properties.profile
            } Else {
                [void] $Binding.profileParameters.Add( $BindingProperty.Name, $Configuration.Properties."$( $BindingProperty.Name )" )
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
        [void] $MetaData.Configuration.Add( $MetaConfig.Name, $JSON.value.configuration."$( $MetaConfig.Name )" )
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
    [void] $Processed.Add( $Property.Name )
}

# $Results.CreateOhItem()
