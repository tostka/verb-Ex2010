#*------v Get-ExchServerFromExServersGroup.ps1 v------
Function Get-ExchServerFromExServersGroup {
  <#
    .SYNOPSIS
    Get-ExchServerFromExServersGroup - Returns the name of an Exchange server by drawing a random box from ad.DOMAIN.com\Exchange Servers grp & regex matches for desired site hubCas names.
    .NOTES
    Author: Todd Kadrie
    Website:	http://tintoys.blogspot.com
    REVISIONS   :
    * 10:02 AM 5/15/2020 pushed the post regex into a infra string & defaulted param, so this could work with any post-filter ;ren Get-ExchServerInLYN -> Get-ExchServerFromExServersGroup
    * 6:59 PM 1/15/2020 cleanup
    # 10:44 AM 9/2/2016 - initial tweak
    .PARAMETER  ServerRegex
    Server filter Regular Expression[-ServerRegex '^CN=(SITE1|SITE2)BOX1[0,1].*$']
    .DESCRIPTION
    Get-ExchServerFromExServersGroup - Returns the name of an Exchange server by drawing a random box from ad.DOMAIN.com\Exchange Servers grp & regex matches for desired site hubCas names.
    Leverages the ActiveDirectory module Get-ADGroupMember cmdlet
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns the name of an Exchange server in the local AD site.
    .EXAMPLE
    .\Get-ExchServerFromExServersGroup
    Draw random matching ex server with defaulted settings
    .EXAMPLE
    .\Get-ExchServerFromExServersGroup -ServerRegex '^CN=SITEPREFIX.*$' 
    Draw random matching ex server with explicit ServerRegex match
    .LINK
    #>
    #Requires -Modules ActiveDirectory
    PARAM(
        [Parameter(HelpMessage="Server filter Regular Expression[-ServerRegex '^CN=(SITE1|SITE2)BOX1[0,1].*$']")]
        $ServerRegex=$rgxLocalHubCAS,
        [Parameter(HelpMessage="AD ParentDomain fqdn [-ADParentDomain 'ROOTDOMAIN.DOMAIN.com']")]
        $ADParentDomain=$DomTORParentfqdn
    ) ;
    (Get-ADGroupMember -Identity 'Exchange Servers' -server $DomTORParentfqdn | 
        ? { $_.distinguishedname -match $ServerRegex }).name | 
            get-random | write-output ;
}

#*------^ Get-ExchServerFromExServersGroup.ps1 ^------