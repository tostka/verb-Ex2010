#*------v cx10tor.ps1 v------
function cx10tor {
    <#
    .SYNOPSIS
    cx10tor - Connect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Connect-EX2010 - Connect-EX2010 to specified on-prem Exchange
    .EXAMPLE
    cx10tor
    #>
    Connect-EX2010 -cred $credTorSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ cx10tor.ps1 ^------
