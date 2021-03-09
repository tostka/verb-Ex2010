#*------v Get-ExchangeServerInSite.ps1 v------
Function Get-ExchangeServerInSite {
    <#
    .SYNOPSIS
    Get-ExchangeServerInSite - Returns the name of an Exchange server in the local AD site.
    .NOTES
    Author: Mike Pfeiffer
    Website:	http://mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
    REVISIONS   :
    * 12:08 PM 5/15/2020 fixed vpn issue: Try/Catch'd around recently failing $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName qry
    * 11:22 AM 3/13/2020 Get-ExchangeServerInSite added a ping-test, to only return matches that are pingable, added -NoPing param, to permit (faster) untested bypass
    * 6:59 PM 1/15/2020 cleanup
    # 10:03 AM 11/16/2018 Get-ExchangeServerInSite:can't do AD-related functions when not AD authentictaed (home, pre-vpn connect). Added if/then test on status and abort balance when false.
    # 12:10 PM 8/1/2017 updated example code at bottom, to accommodate lyn & adl|spb
    # 11:28 AM 3/31/2016 validated that latest round of updates are still functional
    #1:58 PM 9/3/2015 - added pshelp and some docs
    #April 12, 2010 - web version
    .PARAMETER  Site
    Optional: Ex Servers from which Site (defaults to AD lookup against local computer's Site)
    .DESCRIPTION
    Get-ExchangeServerInSite - Returns the name of an Exchange server in the local AD site.
    Uses an ADSI DirectorySearcher to search the current Active Directory site for Exchange 2010 servers.
    Returned object includes the post-filterable Role property which reflects the following
    installed-roles on the discovered server
	    Mailbox Role - 2
        Client Access Role - 4
        Unified Messaging Role - 16
        Hub Transport Role - 32
        Edge Transport Role - 64
        Add the above up to combine roles:
        HubCAS = 32 + 4 = 36
        HubCASMbx = 32+4+2 = 38
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns the name of an Exchange server in the local AD site.
    .EXAMPLE
    .\Get-ExchangeServerInSite
    .EXAMPLE
    get-exchangeserverinsite |?{$_.roles -match "(4|32|36)"}
    Return Hub,CAS,or Hub+CAS servers
    .EXAMPLE
    If(!($ExchangeServer)){$ExchangeServer=(Get-ExchangeServerInSite |?{($_.roles -eq 36) -AND ($_.FQDN -match "SITECODE.*")} | Get-Random ).FQDN }
    Return a random HubCas Role server with a name beginning LYN
    .EXAMPLE
    $regex = '(' + [regex]($ADSiteCodeUK,$ADSiteCodeAU -join "|") + ')'
    switch -regex ($($env:computername).substring(0,3)){
       "$($ADSiteCodeUS)" {$tExRole=36 } ;
       "$($regex)" {$tExRole= 38 }  default { write-error "$((get-date).ToString('HH:mm:ss')):UNRECOGNIZED SERVER PREFIX!."; } ;
    } ;
    $exhubcas = (Get-ExchangeServerInSite |?{($_.roles -eq $tExRole) -AND ($_.FQDN -match "$($env:computername.substring(0,3)).*")} | Get-Random ).FQDN ;
    Return a random HubCas Role server with a name matching the $ENV:COMPUTERNAME
    .LINK
    http://mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
    #>
    [CmdletBinding()]
    PARAM(
        [Parameter(HelpMessage="Switch to suppress default 'pingable' test (e.g. returns all matches, no testing)[-NoPing]")]
        [switch] $NoPing
    ) ;
    $Verbose = ($VerbosePreference -eq 'Continue') ; 
    # 9:53 AM 11/16/2018 from vpn/home, $ADSite doesn't populate prior to domain logon (via vpn)
    # 9:41 AM 5/15/2020 issue: vpn/home, $siteDN suddenly doesn't populate, no longer dyn locs an Ex box, implemented try/catch workaround
    if ($ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]) {
        TRY {$siteDN = $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName}
        CATCH {
            $siteDN =$Ex10siteDN # [infra] returns DN to : cn=[SITENAME],cn=sites,cn=configuration,dc=ad,dc=[DOMAIN],dc=com
            write-warning "$((get-date).ToString('HH:mm:ss')):`$siteDN lookup FAILED, deferring to hardcoded `$Ex10siteDN string in infra file!" ;
        } ; 
        TRY {$configNC = ([ADSI]"LDAP://RootDse").configurationNamingContext}
        CATCH {
            $configNC =$Ex10configNC #  [infra] returns: "CN=Configuration,DC=ad,DC=[DOMAIN],DC=com"
            write-warning "$((get-date).ToString('HH:mm:ss')):`$configNC lookup FAILED, deferring to hardcoded `$Ex10configNC string in infra file!" ;
        } ; 
        if($siteDN -AND $configNC){
            $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC") ;
            $objectClass = "objectClass=msExchExchangeServer" ;
            $version = "versionNumber>=1937801568" ;
            $site = "msExchServerSite=$siteDN" ;
            $search.Filter = "(&($objectClass)($version)($site))" ;
            $search.PageSize = 1000 ;
            [void] $search.PropertiesToLoad.Add("name") ;
            [void] $search.PropertiesToLoad.Add("msexchcurrentserverroles") ;
            [void] $search.PropertiesToLoad.Add("networkaddress") ;
            $search.FindAll() | % {
                $matched = New-Object PSObject -Property @{
                    Name  = $_.Properties.name[0] ;
                    FQDN  = $_.Properties.networkaddress |
                        % { if ($_ -match "ncacn_ip_tcp") { $_.split(":")[1] } } ;
                    Roles = $_.Properties.msexchcurrentserverroles[0] ;
                } ;
                if($NoPing){
                    $matched | write-output ; 
                } else { 
                    $matched | %{If(test-connection $_.FQDN -count 1 -ea 0) {$_} else {} } | 
                        write-output ; 
                } ; 
            } ;
        }else {
            write-warning  "$((get-date).ToString('HH:mm:ss')):MISSING `$siteDN:($($siteDN)) `nOR `$configNC:($($configNC)) values`nABORTING!" ;
            $false | write-output ;
        } ;
    }else {
        write-warning -verbose:$true  "$((get-date).ToString('HH:mm:ss')):`$ADSite blank, not authenticated to a domain! ABORTING!" ;
        $false | write-output ;
    } ;
}

#*------^ Get-ExchangeServerInSite.ps1 ^------