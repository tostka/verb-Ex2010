#*------v Get-ExchangeServerInSite.ps1 v------
Function Get-ExchangeServerInSite {
    <#
    .SYNOPSIS
    Get-ExchangeServerInSite - Returns a summary of all Exchange servers in the local AD site.
    .NOTES
    Version     : 0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 20150903
    FileName    : Get-ExchangeServerInSite.ps1
    License     : (none-asserted)
    Copyright   : (none-asserted)
    Github      : https://github.com/tostka/verb-Ex2010
    Tags        : Powershell
    AddedCredit : Mike Pfeiffer
    AddedWebsite: mikepfeiffer.net
    AddedTwitter: URL
    AddedCredit : Sammy Krosoft 
    AddedWebsite: http://aka.ms/sammy
    AddedTwitter: URL
    REVISIONS
    * 10:31 AM 4/7/2023 added CBH expl of postfilter/sorting to draw predictable pattern 
    * 4:36 PM 4/6/2023 validated Psv51 & Psv20 and Ex10 & 16; added -Roles & -RoleNames params, to perform role filtering within the function (rather than as an external post-filter step). 
    For backward-compat retain historical output field 'Roles' as the msexchcurrentserverroles summary integer; 
    use RoleNames as the text role array; 
     updated for psv2 compat: flipped hash key lookups into properties, found capizliation differences, (psv2 2was all lower case, wouldn't match); 
    flipped the [pscustomobject] with new... psobj, still psv2 doesn't index the hash keys ; updated for Ex13+: Added  16  "UM"; 20  "CAS, UM"; 54  "MBX" Ex13+ ; 16385 "CAS" Ex13+ ; 16439 "CAS, HUB, MBX" Ex13+
    Also hybrided in some good ideas from SammyKrosoft's Get-SKExchangeServers.psm1 
    (emits Version, Site, low lvl Roles # array, and an array of Roles, for post-filtering); 
    # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
    * 12:08 PM 5/15/2020 fixed vpn issue: Try/Catch'd around recently failing $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName qry
    * 11:22 AM 3/13/2020 Get-ExchangeServerInSite added a ping-test, to only return matches that are pingable, added -NoPing param, to permit (faster) untested bypass
    * 6:59 PM 1/15/2020 cleanup
    # 10:03 AM 11/16/2018 Get-ExchangeServerInSite:can't do AD-related functions when not AD authentictaed (home, pre-vpn connect). Added if/then test on status and abort balance when false.
    # 12:10 PM 8/1/2017 updated example code at bottom, to accommodate variant sites
    # 11:28 AM 3/31/2016 validated that latest round of updates are still functional
    #1:58 PM 9/3/2015 - added pshelp and some docs
    #April 12, 2010 - web version
    .DESCRIPTION
    Get-ExchangeServerInSite - Returns a summary of all Exchange servers in the local AD site.

    Uses an ADSI DirectorySearcher to search the current Active Directory site for Exchange on-prem servers.
    Intent is to discover connection points for Powershell, wo the need to preload/pre-connect to Exchange.

    Returned object (in array):
    Site      : {ADSITENAME}
    Roles     : {64}
    Version   : {Version 15.1 (Build 32375.7)}
    Name      : SERVERNAME
    RoleNames : EDGE
    FQDN      : SERVERNAME.DOMAIN.TLD

    ... includes the post-filterable Role property ($_.Role -contains 'CAS') which reflects the following
    installed-roles ('msExchCurrentServerRoles') on the discovered servers
        2   {"MBX"} # Ex10
        4   {"CAS"}
        16  {"UM"}
        20  {"CAS, UM" -split ","} # 
        32  {"HUB"}
        36  {"CAS, HUB" -split ","}
        38  {"CAS, HUB, MBX" -split ","}
        54  {"MBX"} # Ex13+
        64  {"Edge"}
        16385   {"CAS"} # Ex13+
        16439   {"CAS, HUB, MBX"  -split ","} # Ex13+

    .PARAMETER Roles
    Array of msExchCurrentServerRoles 'role' integers to be filtered against (2|4|16|20|32|36|38|54|64|16385|16439)[-Roles @(38,16385)]
    .PARAMETER RoleNames
    Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']    
    .PARAMETER NoPing
    Switch to suppress default 'pingable' test (e.g. returns all matches, no testing)[-NoPing]
    .INPUTS
    None. Does not accepted piped input
    .OUTPUTS
    System.Array of System.Object's
    .EXAMPLE
    PS> If(!($ExchangeServer)){$ExchangeServer = (Get-ExchangeServerInSite| ?{$_.RoleNames -contains 'CAS' -OR $_.RoleNames -contains 'HUB' -AND ($_.FQDN -match "^SITECODE") } | Get-Random ).FQDN
    Return a random Hub Cas Role server in the local Site with a fqdn beginning SITECODE
    .EXAMPLE
    PS> $regex = '(' + [regex]($ADSiteCodeUK,$ADSiteCodeAU -join "|") + ')'
    PS> switch -regex ($($env:computername).substring(0,3)){
    PS>    "$($ADSiteCodeUS)" {$tExRole=36 } ;
    PS>    "$($regex)" {$tExRole= 38 }  default { write-error "$((get-date).ToString('HH:mm:ss')):UNRECOGNIZED SERVER PREFIX!."; } ;
    PS> } ;
    PS> $exhubcas = (Get-ExchangeServerInSite |?{($_.roles -eq $tExRole) -AND ($_.FQDN -match "$($env:computername.substring(0,3)).*")} | Get-Random ).FQDN ;
    Use a switch block to select different role combo targets for a given server fqdn prefix string.
    .EXAMPLE
    PS> $ExchangeServer = Get-ExchangeServerInSite | ?{$_.Roles -match '(4|20|32|36|38|16385|16439)'} | select -expand fqdn | get-random ; 
    Another/Older approach filtering on the Roles integer (targeting combos with Hub or CAS in the mix)
    .EXAMPLE
    PS> $ret = get-exchangeserverinsite -Roles @(4,20,32,36,38,16385,16439) -verbose 
    Demo use of the -Roles param, feeding it an array of Role integer values to be filtered against. In this case, the Role integers that include a CAS or HUB role.
    .EXAMPLE
    PS> $ret = get-exchangeserverinsite -RoleNames 'HUB','CAS' -verbose ;
    Demo use of the -RoleNames param, feeding it the array 'HUB','CAS' Role name strings to be filtered against
    .EXAMPLE
    PS> $ExchangeServer = get-exchangeserverinsite | sort version,roles,name | ?{$_.rolenames -contains 'CAS'}  | select -last 1 | select -expand fqdn ;
    Demo post sorting & filtering, to deliver a rule-based predictable pattern for server selection: 
    Above will always pick the highest Version, 'CAS' RoleName containing, alphabetically last server name (that is pingable). 
    And should stick to that pattern, until the servers installed change, when it will shift to the next predictable box.
    .LINK
    http://mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
    .LINK
    https://github.com/SammyKrosoft/Search-AD-Using-Plain-PowerShell/blob/master/Get-SKExchangeServers.psm1
    .LINK
    https://github.com/tostka/verb-Ex2010
    #>
    [CmdletBinding()]
    PARAM(
        [Parameter(HelpMessage="Array of msExchCurrentServerRoles 'role' integers to be filtered against (2|4|16|20|32|36|38|54|64|16385|16439)[-Roles @(38,16385)]")]
        [ValidateSet(2,4,16,20,32,36,38,54,64,16385,16439)]
        [int[]] $Roles,
        [Parameter(HelpMessage="Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']")]
        [ValidateSet('MBX','CAS','HUB','UM','MBX','EDGE')]
        [string[]] $RoleNames,
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
            [void] $search.PropertiesToLoad.Add("msExchServerSite") ;
            #[void] $search.PropertiesToLoad.Add("objectcategory") ;
            #[void] $search.PropertiesToLoad.Add("objectClass") ;
            #[void] $search.PropertiesToLoad.Add("msExchVersion") ;
            #[void] $search.PropertiesToLoad.Add("msExchMinAdminVersion") ;
            [void] $search.PropertiesToLoad.Add("serialNumber") ;
            $allresults = $search.FindAll() ; 
            $Aggr = @() ; 
            Foreach ($item in $allresults) {
                #if($host.version.major -ge 3){$props=[ordered]@{Dummy = $null ;} }
                #else {$props = New-Object Collections.Specialized.OrderedDictionary} ;
                $props=@{Dummy = $null ;} 
                If($props.Contains("Dummy")){$props.remove("Dummy")} ;
                'Name','FQDN','Version','Site','Roles','RoleNames' | % { $props.add($_,$null) } ;
                # ,'objectcategory','objectclass','msexchversion','msexchminadminversion'
                $props.Name = $item.Properties.name[0]  ; 
                $props.FQDN  = $item.Properties.networkaddress |
                    foreach-object { if ($_ -match "ncacn_ip_tcp") { $_.split(":")[1] } } ;
                #$props.ObjectCategory = $item.Properties.objectcategory ; 
                #$props.ObjectClass = $item.Properties.objectclass ; 
                #$props.Version1 = $item.Properties.msexchversion ; 
                #$props.Version2 = $item.Properties.msexchminadminversion ; 
                $props.Version = $item.Properties.serialnumber ;  
                $props.Site = @("$($item.Properties.msexchserversite -Replace '^CN=|,.*$')") ; 
                $props.Roles = $item.Properties.msexchcurrentserverroles ; 
                $props.RoleNames = switch ($item.Properties.msexchcurrentserverroles){
                    2       {"MBX"} # Ex10
                    4       {"CAS"}
                    16      {"UM"}
                    20      {"CAS;UM".split(';')} 
                    32      {"HUB"}
                    36      {"CAS;HUB".split(';')}
                    38      {"CAS;HUB;MBX".split(';')}
                    54      {"MBX"} # Ex13+
                    64      {"EDGE"}
                    16385   {"CAS"} # Ex13+
                    16439   {"CAS;HUB;MBX".split(';')} # Ex13+
                } ; 
                if($NoPing){
                    <#$props | foreach-object{
                        #$Aggr += [pscustomobject]$_ ; 
                        $Aggr += New-Object -TypeName PsObject -Property $_ ; 
                    } ; 
                    #>
                    $Aggr += New-Object -TypeName PsObject -Property $props ; 
                } else {
                    <#$props | foreach-object{
                         If(test-connection $_.FQDN -count 1 -ea 0) {
                            #$Aggr += [pscustomobject]$_ ; 
                            $Aggr += New-Object -TypeName PsObject -Property $_ ; 
                        } else {} 
                    } ; 
                    #>
                    If(test-connection $props.FQDN -count 1 -ea 0) {
                        #$Aggr += [pscustomobject]$_ ; 
                        $Aggr += New-Object -TypeName PsObject -Property $props ; 
                    } else {} 
                } ;
                # Roles, RoleNames
                
            } ; 
            # use indexed hash to self-limit dupe addition (faster than -contains as well)
            $httmp = @{} ; 
            if($Roles){
                # can match the roles integer w a regex OR'd on the values
                [regex]$rgxRoles = ('(' + (($roles |%{[regex]::escape($_)}) -join '|') + ')') ;
                $matched =  @( $aggr | ?{$_.Roles -match $rgxRoles}) ; 
                foreach($m in $matched){
                    if($httmp[$m.name]){} else { 
                        $httmp[$m.name] = $m ; 
                    } ; 
                } ; 
            } ; 
            if($RoleNames){
                # to do multivalue -contains, you need to -OR the combo ($x -contains 'value' -OR $x -contains 'other'), 
                # or loop the compares and add per pass (using index hash to exclude dupe adds)
                foreach ($RoleName in $RoleNames){
                    $matched = @($Aggr | ?{$_.RoleNames -contains $RoleName} ) ; 
                    foreach($m in $matched){
                        if($httmp[$m.name]){} else { 
                            $httmp[$m.name] = $m ; 
                        } ; 
                    } ; 
                } ; 
            } ; 
            # hashtable always reads 'populated', so check if it has postive count, and then assign back to $aggr.
            if(($httmp.Values| measure).count -gt 0){
                $Aggr  = $httmp.Values ; 
            } ; 
            $Aggr | write-output ; 
        }else {
            write-warning  "$((get-date).ToString('HH:mm:ss')):MISSING `$siteDN:($($siteDN)) `nOR `$configNC:($($configNC)) values`nABORTING!" ;
            $false | write-output ;
        } ;
    }else {
        write-warning "$((get-date).ToString('HH:mm:ss')):`$ADSite blank, not authenticated to a domain! ABORTING!" ;
        $false | write-output ;
    } ;
} ; 
#*------^ Get-ExchangeServerInSite.ps1 ^------