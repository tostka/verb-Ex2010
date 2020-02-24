#*----------------v Function Get-ExchangeServerInSite v----------------
Function Get-ExchangeServerInSite {
  <#
    .SYNOPSIS
    Get-ExchangeServerInSite - Returns the name of an Exchange server in the local AD site.
    .NOTES
    Author: Mike Pfeiffer
    Website:	http://mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
    REVISIONS   :
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
    If(!($ExchangeServer)){$ExchangeServer=(Get-ExchangeServerInSite |?{($_.roles -eq 36) -AND ($_.FQDN -match "LYN.*")} | Get-Random ).FQDN }
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
  # 9:53 AM 11/16/2018 from vpn/home, $ADSite doesn't populate prior to domain logon (via vpn)
  if ($ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]) {
    $siteDN = $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName ;
    # if conn'd rets: cn=lyndale,cn=sites,cn=configuration,dc=ad,dc=toro,dc=com
    $configNC = ([ADSI]"LDAP://RootDse").configurationNamingContext ; # returns: "CN=Configuration,DC=ad,DC=toro,DC=com"
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
      New-Object PSObject -Property @{
        Name  = $_.Properties.name[0] ;
        FQDN  = $_.Properties.networkaddress |
        % { if ($_ -match "ncacn_ip_tcp") { $_.split(":")[1] } } ;
        Roles = $_.Properties.msexchcurrentserverroles[0] ;
      } ;
    } ;
  }
  else {
    write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):`$ADSite blank, not authenticated to a domain! ABORTING!" ;
    $false | write-output ;
  } ;
} #*----------------^ END Function Get-ExchangeServerInSite ^---------------- ;