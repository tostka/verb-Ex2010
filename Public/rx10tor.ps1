#*------v rx10tor.ps1 v------
function rx10tor {
    <#
    .SYNOPSIS
    rx10tor - Reonnect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Reconnect-EX2010 - Reonnect to specified on-prem Exchange
    .EXAMPLE
    rx10tor
    #>
    Reconnect-EX2010 -cred $credTorSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ rx10tor.ps1 ^------