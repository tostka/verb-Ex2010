#*------v Function Get-ExchServerInLYN v------
Function Get-ExchServerInLYN {
  <#
    .SYNOPSIS
    Get-ExchServerInLYN - Returns the name of an Exchange server in the LYN site (much simpler, pulls random box from ad.toro.com\Exchange Servers grp, & filters names for LYN and hubcas nameschemes).
    .NOTES
    Author: Todd Kadrie
    Website:	http://tintoys.blogspot.com
    REVISIONS   :
    * 6:59 PM 1/15/2020 cleanup
    # 10:44 AM 9/2/2016 - initial tweak
    .PARAMETER  Site
    Optional: Ex Servers from which Site (defaults to AD lookup against local computer's Site)
    .DESCRIPTION
    Get-ExchServerInLYN - Returns the name of an Exchange server in the local LYN AD site.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns the name of an Exchange server in the local AD site.
    .EXAMPLE
    .\Get-ExchServerInLYN
    .LINK
    #>
  (Get-ADGroupMember -Identity 'Exchange Servers' -server $DomTORParentfqdn | ? { $_.distinguishedname -match $rgxLocalHubCAS }).name | get-random | write-output ;
} #*------^ END Function Get-ExchServerInLYN ^------