#*------v rx10cmw.ps1 v------
function rx10cmw {
    <#
    .SYNOPSIS
    rx10cmw - Reonnect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Reconnect-EX2010 - Reonnect to specified on-prem Exchange
    .EXAMPLE
    rx10cmw
    #>
    Reconnect-EX2010 -cred $credCMWSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ rx10cmw.ps1 ^------