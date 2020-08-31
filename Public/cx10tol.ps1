#*------v cx10tol.ps1 v------
function cx10tol {
    <#
    .SYNOPSIS
    cx10tol - Connect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Connect-EX2010 - Connect-EX2010 to specified on-prem Exchange
    .EXAMPLE
    cx10tol
    #>
    Connect-EX2010 -cred $credtolSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ cx10tol.ps1 ^------
