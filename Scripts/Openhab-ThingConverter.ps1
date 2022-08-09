$ThingsRaw = Get-Content $PSScriptRoot\org.openhab.core.thing.Thing.JSON | ConvertFrom-Json
$Outfile = "$PSScriptRoot\allthings.things"

$ThingsFilter = '.*'
# $ThingsFilter = 'avm'
# $ThingsFilter = '099950388533'

$Processed = [Collections.ArrayList]::new()
$Things = [Collections.ArrayList]::new()

# define properties for bridges, things and channels

class Bridge {

    [String] $BindingID # first part of UID
    [String] $BridgeType # second part of UID
    [String] $BridgeID  # remaining UID parts
    [String] $label
    [String] $location
    
    [Collections.ArrayList] $Configuration = [Collections.ArrayList]::new()
    [Collections.ArrayList] $Things = [Collections.ArrayList]::new()
    
    # enable the bridge class to output an Openhab .things definition string - https://www.openhab.org/docs/configuration/things.html

    [String] ToString() {
        Return $This.CreateItem()
    }

    [String] Hidden CreateItem() {

        # bridge definition in .things files as documented

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
            $Return = $Return.Substring( 0, $Return.Length - 3 ) + "`r`n]" # remove last comma and reappend CR/LF]
        }

        # if there are things using this bridge, include them in the bridge definition within curly brackets

        If ( $This.Things.Count -gt 0 ) {
            $Return += " {`r`n"
            Foreach ( $Thing in $This.Things ) {
                # ToString( $true ) means "this thing is a child of a bridge" - this requires more indendation than a
                # standalone thing definition. And again, Openhab is picky about indendation :)
                $Return += $Thing.ToString( $True )
            }
            $Return += "}`r`n"
        }
        Return $Return
    }
}

Class Thing {

    [String] $BindingID # first part of UID
    [String] $TypeID    # second part of UID
    [String] $BridgeID  # third part of UID
    [String] $ThingID   # remaining UID parts
    [String] $label
    [String] $Location

    [Collections.ArrayList] $Configuration = [Collections.ArrayList]::new()
    [Collections.Arraylist] $Channels = [Collections.ArrayList]::new()
    
    # enable the things class to output an Openhab .things definition string - https://www.openhab.org/docs/configuration/things.html

    [String] ToString( ){
        Return $This.CreateItem( $False )
    }
    [String] ToString( $IsBridgeChild ) {
        Return $This.CreateItem( $IsBridgeChild )
    }

    [String] Hidden CreateItem( $IsBridgeChild ) {

        # if the item is a child of a bridge, we need 2 spaces more at the beginning of each line

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
            $Return = $Return.Substring( 0, $Return.Length - 2 ) + ' ]' # cutoff last comma
        }
        
        # if the thing has channels, include them in curly brackets as well
        # again, take care of correct indendation
        
        If ( $This.Channels.Count -gt 0 ) {
            $Return += " {`r`n"
            $Return += ' ' * $Indent + "  Channels:`r`n"
            Foreach ( $Channel in $This.Channels ) {
                $Return += $Channel.ToString( $Indent + 4 )
            }
            $Return += "}"
        }
        $Return += "`r`n"
        Return $Return 
    }

}

Class Channel {

    [String] $Kind
    [String] $Type
    [String] $ID
    [String] $Name

    [Collections.ArrayList] $Configuration = [Collections.ArrayList]::new()
    
    # enable the channel to output an Openhab .things definition string - https://www.openhab.org/docs/configuration/things.html
    
    [String] ToString( ){
        Return $This.CreateItem( 4 )
    }

    [String] ToString( [int] $Indent ) {
        Return $This.CreateItem( $Indent )
    }

    [String] Hidden CreateItem( [int] $Indent = 4 ){


        [String] $Return = ''

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
            $Return = $Return.Substring( 0, $Return.Length - 2 ) + " ]`r`n"
        }

        Return $Return
    }

}

Class Config {
    [String] $ValueType
    [String] $ValueName
    [String] $ValueData

    Config () {}
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
                $Return = $DecimalValue.ToString( [cultureinfo]::new( 'en-US' ) )
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

# first, grab all bridges - JSON in Powershell is pretty awkward for iterating over...

Foreach ( $Property in $ThingsRaw | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $ThingsFilter } ) {

    $JSON = $ThingsRaw."$( $Property.Name )"

    # make sure it is a bridge

    If ( $JSON.value.isBridge ) {

        # basic bridge data

        $Bridge = [Bridge]::new()
        $Bridge.label = $JSON.value.label
        $Bridge.location = $JSON.value.location
        $Bridge.BindingID = $JSON.value.UID.Split( ':', 3 )[0]
        $Bridge.BridgeType = $JSON.value.UID.Split( ':', 3 )[1]
        $Bridge.BridgeID = $JSON.value.UID.Split( ':', 3 )[2]

        Write-Verbose "Processing bridge: $( $JSON.value.UID )"

        # if the bridge has configuration values, add them to the bridge object

        If ( $JSON.value.Configuration ) {
            Foreach ( $Config in $JSON.value.Configuration | Get-Member -MemberType NoteProperty ) {
                If ( $Config.Definition -match "^(?<ValueType>\w+)\s+$( $Config.Name )=(?<ValueData>.+)$" ) {
                    Write-Verbose "Processing bridge configuration: $( $Config.Definition )"
                    $ConfigValue = [Config]::new( $Matches.ValueType, $Config.Name, $Matches.ValueData )
                    [void] $Bridge.Configuration.Add( $ConfigValue )
                }
            }
        }
        [void] $Things.Add( $Bridge )
        [void] $Processed.Add( $Property.Name )
    }
}

# now get all remaining stuff that are real things

Foreach ( $Property in $ThingsRaw | Get-Member -MemberType NoteProperty | Where-Object { $Processed -notcontains $_.Name } | Where-Object { $_.Name -match $ThingsFilter } ) {

    $JSON = $ThingsRaw."$( $Property.Name )"

    $Thing = [Thing]::new()
    $Thing.label = $JSON.value.label
    $Thing.location = $JSON.value.location
    $Thing.BindingID = $JSON.value.UID.Split( ':', 4 )[0]
    $Thing.TypeID = $JSON.value.UID.Split( ':', 4 )[1]

    Write-Verbose "Processing thing: $( $Thing.Name )"

    If ( $JSON.value.BridgeUID ) {
        # if the thing uses a bridge, the bridge ID will be part of its UID which thus contains 4 segments...
        $Thing.BridgeID = $JSON.value.UID.Split( ':', 4 )[2]
        $Thing.ThingID = $JSON.value.UID.Split( ':', 4 )[3]
    } Else {
        # ...if it is a standalone thing, its UID will only contain 3 segments.
        $Thing.ThingID = $JSON.value.UID.Split( ':', 4 )[2]
    }

    $Thing.Location = $JSON.value.location

    If ( $JSON.value.Configuration ) {
        Foreach ( $Config in $JSON.value.Configuration | Get-Member -MemberType NoteProperty ) {
            If ( $Config.Definition -match "^(?<ValueType>\w+)\s+$( $Config.Name )=(?<ValueData>.+)$" ) {
                Write-Verbose "Processing thing configuration: $( $Config.Definition )"
                $ConfigValue = [Config]::new( $Matches.ValueType, $Config.Name, $Matches.ValueData )
                [void] $Thing.Configuration.Add( $ConfigValue )
            }
        }
    }

    Foreach ( $Ch in $JSON.value.channels ) {
        $Channel = [Channel]::new()
        $Channel.Name = $Ch.Label
        $Channel.Kind = $Ch.Kind
        If ( $Ch.Kind -eq 'TRIGGER' ) {
            # trigger channels do not define their type because it must be 'String'
            $Channel.Type = 'String'
        } Else {
            $Channel.Type = $Ch.itemType
        }
        Write-Verbose "Processing thing channel: $( $Channel.Name ): $( $Channel.Kind ) of type $( $Channel.Type )"
        If ( $ch.uid -match ':(?<ID>[^:]+)$' ) {
            # channel ID needs to be extracted from channel UID - always last segment
            $Channel.ID = $Matches.ID
        }
        Foreach ( $Config in $Ch.Configuration | Get-Member -MemberType NoteProperty ) {
            If ( $Config.Definition -match "^(?<ValueType>\w+)\s+$( $Config.Name )=(?<ValueData>.+)$" ) {
                Write-Verbose "Processing channel configuration: $( $Config.Definition )"
                $ConfigValue = [Config]::new( $Matches.ValueType, $Config.Name, $Matches.ValueData )
                [void] $Channel.Configuration.Add( $ConfigValue )
            }
        }

        # only add the channel if any configurations were found
        # all standard channels (without configuration) will be added anyway by the binding
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
    [void] $Processed.Add( $Property.Name )
}

$encoding = [System.Text.Encoding]::GetEncoding(1252)
$streamWriter = [IO.StreamWriter]::new( $Outfile, $false, $Encoding)
$Things | ForEach-Object { $_.ToString() | ForEach-Object { $streamWriter.WriteLine( $_ ) } }
$streamWriter.Dispose()
