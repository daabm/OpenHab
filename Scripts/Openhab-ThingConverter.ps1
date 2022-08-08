$ThingsRaw = Get-Content $PSScriptRoot\org.openhab.core.thing.Thing.JSON | ConvertFrom-Json
$Outfile = "$PSScriptRoot\allthings.things"

$ThingsFilter = '.*'
# $ThingsFilter = 'tesla'

$Processed = [Collections.ArrayList]::new()
$Things = [Collections.ArrayList]::new()

# define properties for bridges, things and channels

class Bridge {

    [String] $BindingID # first part of UID
    [String] $BridgeType # second part of UID
    [String] $BridgeID  # remaining UID parts
    [String] $Name
    
    [Collections.Hashtable] $Configuration = [Collections.Hashtable]::new()
    [Collections.ArrayList] $Things = [Collections.ArrayList]::new()
    
    # enable the bridge class to output an Openhab .things definition string - https://www.openhab.org/docs/configuration/things.html

    [String] CreateOHItem() {
        Return $This.CreateItem()
    }

    [String] Hidden CreateItem() {

        # bridge definition in .things files as documented

        [String] $BridgeReturn = $This.GetType().Name + ' ' + $This.BindingID + ':' + $This.BridgeType + ':' + $This.BridgeID

        # if the bridge has configuration values, insert them in square brackets
        # and take care of indentation - Openhab is quite picky about misalignment :-)

        If ( $This.Configuration.Count -gt 0 ) {
            $BridgeReturn += " [`r`n"
            Foreach ( $Key in $This.Configuration.Keys ) {
                $BridgeReturn += '  ' + $Key + '=' + $This.Configuration[ $Key ] + ",`r`n"
            }
            $BridgeReturn = $BridgeReturn.Substring( 0, $BridgeReturn.Length - 3 ) + "`r`n]" # remove last comma and reappend CR/LF]
        }

        # if there are things using this bridge, include them in the bridge definition within curly brackets

        If ( $This.Things.Count -gt 0 ) {
            $BridgeReturn += " {`r`n"
            Foreach ( $Thing in $This.Things ) {
                # CreateOHThing( $true ) means "this thing is a child of a bridge" - this requires more indendation than a
                # standalone thing definition. And again, Openhab is picky about indendation :)
                $BridgeReturn += $Thing.CreateOHItem( $True )
            }
            $BridgeReturn += "}`r`n"
        }
        Return $BridgeReturn
    }
}

Class Thing {

    [String] $BindingID # first part of UID
    [String] $TypeID    # second part of UID
    [String] $BridgeID  # third part of UID
    [String] $ThingID   # remaining UID parts
    [String] $Name
    [String] $Location

    [Collections.Hashtable] $Configuration = [Collections.Hashtable]::new()
    [Collections.Arraylist] $Channels = [Collections.ArrayList]::new()
    
    # enable the things class to output an Openhab .things definition string - https://www.openhab.org/docs/configuration/things.html

    [String] CreateOHItem( ){
        Return $This.CreateItem( $False )
    }
    [String] CreateOHItem( $IsBridgeChild ) {
        Return $This.CreateItem( $IsBridgeChild )
    }

    [String] Hidden CreateItem( $IsBridgeChild ) {

        # if the item is a child of a bridge, we need 2 spaces more at the beginning of each line

        If ( $IsBridgeChild ) { 
            $Indent = 2 
            [String] $ThingReturn = ' ' * $Indent + $This.GetType().Name + ' ' + $This.TypeID + ' ' + $This.ThingID
        } Else { 
            $Indent = 0 
            [String] $ThingReturn = ' ' * $Indent + $This.GetType().Name + ' ' + $This.BindingID + ':' + $This.TypeID + ':' + $This.ThingID
        }

        # if the thing has configuration values, add them in square brackets
 
        If ( $This.Configuration.Count -gt 0 ) {
            # $ThingReturn += " [`r`n"
            $ThingReturn += ' [ '
            Foreach ( $Key in $This.Configuration.Keys ) {
                # $ThingReturn += ' ' * $Indent + '  ' + $Key + '=' + $This.Configuration[ $Key ] + ",`r`n"
                $ThingReturn += $Key + '=' + $This.Configuration[ $Key ] + ', '
            }
            $ThingReturn = $ThingReturn.Substring( 0, $ThingReturn.Length - 2 ) + ' ]' # cutoff last comma
        }
        
        # if the thing has channels, include them in curly brackets as well
        # again, take care of correct indendation
        
        If ( $This.Channels.Count -gt 0 ) {
            $ThingReturn += " {`r`n"
            [String] $ThingReturn += ' ' * $Indent + "  Channels:`r`n"
            Foreach ( $Channel in $This.Channels ) {
                $ThingReturn += $Channel.CreateOHItem( $Indent + 4 )
            }
            $ThingReturn += "}"
        }
        $ThingReturn += "`r`n"
        Return $ThingReturn 
    }

}

Class Channel {

    [String] $Kind
    [String] $Type
    [String] $ID
    [String] $Name

    [Hashtable] $Configuration = [Collections.Hashtable]::new()
    
    # enable the channel to output an Openhab .things definition string - https://www.openhab.org/docs/configuration/things.html
    
    [String] CreateOHItem( ){
        Return $This.CreateItem( 4 )
    }
    [String] CreateOHItem( [int] $Indent ) {
        Return $This.CreateItem( $Indent )
    }

    [String] Hidden CreateItem( [int] $Indent = 4 ){

        [String] $ChannelReturn = ''

        Foreach ( $Key in $This.Configuration.Keys ) {

            $ChannelReturn += ' ' * $Indent + $This.Kind.Substring( 0, 1 ).ToUpper() + $This.Kind.Substring( 1 ).ToLower() + ' ' + $This.Type + ' : ' + $This.ID
            If ( $This.Name ) {
                $ChannelReturn += ' "' + $This.Name + '"'
            }
            
            # if the channel has configuration values, append them in square brackets
            
            If ( $This.Configuration.Count -gt 0 ) {
                $ChannelReturn += ' [ '
                Foreach ( $Key in $This.Configuration.Keys ) {
                    $rtn = ''
                    If ( [double]::TryParse( $This.Configuration[ $Key ], [ref] $rtn )) { # check if we have a number, otherwise we need surrounding double quotes
                        $ChannelReturn += $This.Configuration[ $Key ].ToString() -replace ',', '.'
                    } Else {
                        $ChannelReturn += '"' + $This.Configuration[ $Key ] + '"'
                    }
                }
                $ChannelReturn = $ChannelReturn.Substring( 0, $ChannelReturn.Length - 2 ) + " ]`r`n"
            }
        }
        Return $ChannelReturn
    }

}

# first, grab all bridges - JSON in Powershell is pretty awkward for iterating over...

Foreach ( $Property in $ThingsRaw | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match $ThingsFilter } ) {

    $JSON = $ThingsRaw."$( $Property.Name )"

    # make sure it is a bridge

    If ( $JSON.value.isBridge ) {

        # basic bridge data

        $Bridge = [Bridge]::new()
        $Bridge.Name = $JSON.value.label
        $Bridge.BindingID = $JSON.value.UID.Split( ':', 3 )[0]
        $Bridge.BridgeType = $JSON.value.UID.Split( ':', 3 )[1]
        $Bridge.BridgeID = $JSON.value.UID.Split( ':', 3 )[2]

        # if the bridge has configuration values, add them to the bridge object

        If ( $JSON.value.Configuration ) {
            Foreach ( $Config in $JSON.value.Configuration | Get-Member -MemberType NoteProperty ) {
                If ( $COnfig.Definition -match "^(?<Type>\w+)\s+$( $Config.Name )=(?<Value>.+)$" ) {
                    $ConfigValue = $Matches.Value
                    If ( $Matches.Type -ne 'decimal' ) { $ConfigValue = """$ConfigValue""" }
                }
                $Bridge.Configuration.Add( $Config.Name, $ConfigValue )
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
    $Thing.Name = $JSON.value.label
    $Thing.BindingID = $JSON.value.UID.Split( ':', 4 )[0]
    $Thing.TypeID = $JSON.value.UID.Split( ':', 4 )[1]

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
            If ( $COnfig.Definition -match "^(?<Type>\w+)\s+$( $Config.Name )=(?<Value>.+)$" ) {
                $ConfigValue = $Matches.Value
                If ( $Matches.Type -ne 'decimal' ) { $ConfigValue = """$ConfigValue""" }
            }
            $Thing.Configuration.Add( $Config.Name, $ConfigValue )
        }
    }

    Foreach ( $Ch in $JSON.value.channels ) {
        $Channel = [Channel]::new()
        $Channel.Kind = $Ch.Kind
        If ( $Ch.Kind -eq 'TRIGGER' ) {
            # trigger channels do not define their type because it must always be 'String'
            $Channel.Type = 'String'
        } Else {
            # the channel type is the first part of the itemType if the channel defines an item subtype
            # if there's no subtype, this works as well
            $Channel.Type = $Ch.itemType.Split( ':', 2 )[0]
        }
        If ( $ch.uid -match ':(?<ID>[^:]+)$' ) {
            # channel ID needs to be extracted from channel UID - always last segment of UID
            $Channel.ID = $Matches.ID
        }
        $Channel.Name = $Ch.Label
        Foreach ( $Config in $Ch.Configuration | Get-Member -MemberType NoteProperty ) {
            If ( $Config.Definition -match "^(?<Type>\w+)\s+$( $Config.Name )=(?<Value>.+)$" ) {
                $ConfigValue = $Matches.Value
                # for non decimal values, surround with double hyphens
                If ( $Matches.Type -ne 'decimal' ) { $ConfigValue = """$ConfigValue""" }
            }
            $Channel.Configuration.Add( $Config.Name, $ConfigValue )
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
$Things.CreateOHItem() | ForEach-Object { $streamWriter.WriteLine( $_ ) }
$streamWriter.Dispose()
