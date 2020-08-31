#*------v cx10cmw.ps1 v------
function cx10cmw {
    <#
    .SYNOPSIS
    cx10tol - Connect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Connect-EX2010 - Connect-EX2010 to specified on-prem Exchange
    .EXAMPLE
    cx10cmw
    #>
    Connect-EX2010 -cred $credCMWSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ cx10cmw.ps1 ^------
