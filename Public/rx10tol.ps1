#*------v rx10tol.ps1 v------
function rx10tol {
    <#
    .SYNOPSIS
    rx10tol - Reonnect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Reconnect-EX2010 - Reonnect to specified on-prem Exchange
    .EXAMPLE
    rx10tol
    #>
    Reconnect-EX2010 -cred $credtolSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ rx10tol.ps1 ^------