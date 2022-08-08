# $Things = Get-Content .\org.openhab.core.thing.Thing.json | ConvertFrom-Json
$ItemsRaw = Get-Content $PSScriptRoot\org.openhab.core.items.Item.JSON | ConvertFrom-Json
$BindingsRaw = Get-Content $PSScriptRoot\org.openhab.core.thing.link.ItemChannelLink.json | ConvertFrom-Json
$MetadataRaw = Get-Content $PSScriptRoot\org.openhab.core.items.Metadata.json | ConvertFrom-Json
$ItemsFilter = '.*'
$ItemsFilter = 'TSOG2Ku'

$Processed = [Collections.ArrayList]::new()
$Items = [Collections.ArrayList]::new()
$Bindings = [Collections.ArrayList]::new()

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
        If ( $This.Bindings.Count -eq 1 ) {
            $ItemReturn += ' { ' + $This.Bindings[0].CreateOHBinding() + ' }'
        } ElseIf ( $This.Bindings.Count -gt 1 ) {
            $ItemReturn += " {`r`n"
            Foreach ( $Binding in $This.Bindings ) {
                $ItemReturn += $Binding.CreateOHBinding() + ",`r`n"
            }
            $ItemReturn = $ItemReturn.Substring( 0, $ItemReturn.Length - 3 ) + "`r`n}`r`n"
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
            If ( $This.ProfileParameters ) {
                Foreach ( $Key in $This.profileParameters.Keys ) {
                    $BindingReturn += ', ' + $Key + '='
                    $rtn = ''
                    If ( [double]::TryParse( $This.profileParameters[ $Key ], [ref]$rtn )) { # check if we have a number, otherwise we need surrounding double quotes
                        $BindingReturn += $This.profileParameters[ $Key ]
                    } Else {
                        $BindingReturn += '"' + $This.profileParameters[ $Key ] + '"'
                    }
                }
            }
            $BindingReturn += ' ]'
        }
        Return $BindingReturn
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

    [void] $Items.Add( $Item )
    [void] $Processed.Add( $Property.Name )
}

# $Results.CreateOhItem()
