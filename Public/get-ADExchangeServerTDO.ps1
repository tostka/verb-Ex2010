# get-ADExchangeServerTDO.ps1


#region GET_ADEXCHANGESERVERTDO ; #*------v get-ADExchangeServerTDO v------
#if(-not(gi function:get-ADExchangeServerTDO -ea 0)){
    Function get-ADExchangeServerTDO {
        <#
        .SYNOPSIS
        get-ADExchangeServerTDO.ps1 - Returns Exchangeserver summary(s) from AD records
        .NOTES
        Version     : 3.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2015-09-03
        FileName    : get-ADExchangeServerTDO.ps1
        License     : (none-asserted)
        Copyright   : (none-asserted)
        Github      : https://github.com/tostka/verb-Ex2010
        Tags        : Powershell, ActiveDirectory, Exchange, Discovery
        AddedCredit : Mike Pfeiffer
        AddedWebsite: mikepfeiffer.net
        AddedTwitter: URL
        AddedCredit : Sammy Krosoft 
        AddedWebsite: http://aka.ms/sammy
        AddedTwitter: URL
        AddedCredit : Brian Farnsworth
        AddedWebsite: https://codeandkeep.com/
        AddedTwitter: URL
        REVISIONS
        * 10;05 am 4/30/2025 fixed code for Edge role in raw PS, missing evaris for Ex: added discovery from reg & stock file system dirs for version etc.
        * 3:57 PM 11/26/2024 updated simple write-host,write-verbose with full pswlt support;  syncd dbg & vx10 copies.
        * 12:57 PM 6/11/2024 Validated, Ex2010 & Ex2019, hub, mail & edge roles: tested ☑️ on CMW mail role (Curly); and Jumpbox; copied in CBH from repo copy, which has been updated/debugged compat on CMW Edge 
        * 2:05 PM 8/28/2023 REN -> Get-ExchangeServerInSite -> get-ADExchangeServerTDO (aliased orig); to better steer profile-level options - including in cmw org, added -TenOrg, and default Site to constructed vari, targeting new profile $XXX_ADSiteDefault vari; Defaulted -Roles to HUB,CAS as well.
        * 3:42 PM 8/24/2023 spliced together combo of my long-standing, and some of the interesting ideas BF's version had. Functional prod:
            - completely removed ActiveDirectory module dependancies from BF's code, and reimplemented in raw ADSI calls. Makes it fully portable, even into areas like Edge DMZ roles, where ADMS would never be installed.

        * 3:17 PM 8/23/2023 post Edge testing: some logic fixes; add: -Names param to filter on server names; -Site & supporting code, to permit lookup against sites *not* local to the local machine (and bypass lookup on the local machine) ; 
            ren $Ex10siteDN -> $ExOPsiteDN; ren $Ex10configNC -> $ExopconfigNC
        * 1:03 PM 8/22/2023 minor cleanup
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
        * 11/18/18 BF's posted rev
        # 12:10 PM 8/1/2017 updated example code at bottom, to accommodate variant sites
        # 11:28 AM 3/31/2016 validated that latest round of updates are still functional
        #1:58 PM 9/3/2015 - added pshelp and some docs
        #April 12, 2010 - web version
        .DESCRIPTION
        get-ADExchangeServerTDO.ps1 - Returns Exchangeserver summary(s) from AD records

        Hybrided together ideas from Brian Farnsworth's blog post
        [PowerShell - ActiveDirectory and Exchange Servers – CodeAndKeep.Com – Code and keep calm...](https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/)
        ... with much older concepts from  Sammy Krosoft, and much earlier Mike Pfeiffer. 

        - Subbed in MP's use of ADSI for ActiveDirectory Ps mod cmds - it's much more dependancy-free; doesn't require explicit install of the AD ps module
        ADSI support is built into windows.
        - spliced over my addition of Roles, RoleNames, Name & NoTest params, for prefiltering and suppressing testing.


        [briansworth · GitHub](https://github.com/briansworth)

        Uses an ADSI DirectorySearcher to search the current Active Directory site for Exchange on-prem servers.
                Intent is to discover connection points for Powershell, wo the need to preload/pre-connect to Exchange.

                But, as a non-Exchange-Management-Shell-dependant info source on Exchange Server configs, it can be used before connection, with solely AD-available data, to check configuration spes on the subject server(s). 

                For example, this query will return sufficient data under Version to indicate which revision of Exchange is in use:


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
                    64  {"EDGE"}
                    16385   {"CAS"} # Ex13+
                    16439   {"CAS, HUB, MBX"  -split ","} # Ex13+
        .PARAMETER Roles
        Array of msExchCurrentServerRoles 'role' integers to be filtered against (2|4|16|20|32|36|38|54|64|16385|16439)[-Roles @(38,16385)]
        .PARAMETER RoleNames
        Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']
        .PARAMETER Server
        Array of Server name strings to be filtered against[-Server 'SERVER1','SERVER2']
        .PARAMETER SiteName
        Name of specific AD SiteName to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-SiteName 'SITENAME']
        .PARAMETER TenOrg
        Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
        .PARAMETER NoPing
        Switch to suppress default 'pingable' test (e.g. returns all matches, no testing)[-NoPing]
        .INPUTS
        None. Does not accepted piped input.(.NET types, can add description)
        .OUTPUTS
        None. Returns no objects or output (.NET types)
        System.Boolean
        [| get-member the output to see what .NET obj TypeName is returned, to use here]
        System.Array of System.Object's
        .EXAMPLE
        PS> If(!($ExchangeServer)){$ExchangeServer = (get-ADExchangeServerTDO| ?{$_.RoleNames -contains 'CAS' -OR $_.RoleNames -contains 'HUB' -AND ($_.FQDN -match "^SITECODE") } | Get-Random ).FQDN
        Return a random Hub Cas Role server in the local Site with a fqdn beginning SITECODE
        .EXAMPLE
        PS> $localADExchserver = get-ADExchangeServerTDO -Names $env:computername -SiteName ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().name)
        Demo, if run from an Exchange server, return summary details about the local server (-SiteName isn't required, is default imputed from local server's Site, but demos explicit spec for remote sites)
        .EXAMPLE
        PS> $regex = '(' + [regex]($ADSiteCodeUK,$ADSiteCodeAU -join "|") + ')'
        PS> switch -regex ($($env:computername).substring(0,3)){
        PS>    "$($ADSiteCodeUS)" {$tExRole=36 } ;
        PS>    "$($regex)" {$tExRole= 38 }  default { write-error "$((get-date).ToString('HH:mm:ss')):UNRECOGNIZED SERVER PREFIX!."; } ;
        PS> } ;
        PS> $exhubcas = (get-ADExchangeServerTDO |?{($_.roles -eq $tExRole) -AND ($_.FQDN -match "$($env:computername.substring(0,3)).*")} | Get-Random ).FQDN ;
        Use a switch block to select different role combo targets for a given server fqdn prefix string.
        .EXAMPLE
        PS> $ExchangeServer = get-ADExchangeServerTDO | ?{$_.Roles -match '(4|20|32|36|38|16385|16439)'} | select -expand fqdn | get-random ; 
        Another/Older approach filtering on the Roles integer (targeting combos with Hub or CAS in the mix)
        .EXAMPLE
        PS> $ret = get-ADExchangeServerTDO -Roles @(4,20,32,36,38,16385,16439) -verbose 
        Demo use of the -Roles param, feeding it an array of Role integer values to be filtered against. In this case, the Role integers that include a CAS or HUB role.
        .EXAMPLE
        PS> $ret = get-ADExchangeServerTDO -RoleNames 'HUB','CAS' -verbose ;
        Demo use of the -RoleNames param, feeding it the array 'HUB','CAS' Role name strings to be filtered against
        PS> $ret = get-ADExchangeServerTDO -Names 'SERVERName' -verbose ;
        Demo use of the -RoleNames param, feeding it the array 'HUB','CAS' Role name strings to be filtered against
        .EXAMPLE
        PS> $ExchangeServer = get-ADExchangeServerTDO | sort version,roles,name | ?{$_.rolenames -contains 'CAS'}  | select -last 1 | select -expand fqdn ;
        Demo post sorting & filtering, to deliver a rule-based predictable pattern for server selection: 
        Above will always pick the highest Version, 'CAS' RoleName containing, alphabetically last server name (that is pingable). 
        And should stick to that pattern, until the servers installed change, when it will shift to the next predictable box.
        .EXAMPLE
        PS> $ExOPServer = get-ADExchangeServerTDO -Name LYNMS650 -SiteName Lyndale
        PS> if([double]$ExVersNum = [regex]::match($ExOPServer.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
        PS>     switch -regex ([string]$ExVersNum) {
        PS>         '15\.2' { $isEx2019 = $true ; $ExVers = 'Ex2019' }
        PS>         '15\.1' { $isEx2016 = $true ; $ExVers = 'Ex2016'}
        PS>         '15\.0' { $isEx2013 = $true ; $ExVers = 'Ex2013'}
        PS>         '14\..*' { $isEx2010 = $true ; $ExVers = 'Ex2010'}
        PS>         '8\..*' { $isEx2007 = $true ; $ExVers = 'Ex2007'}
        PS>         '6\.5' { $isEx2003 = $true ; $ExVers = 'Ex2003'}
        PS>         '6|6\.0' {$isEx2000 = $true ; $ExVers = 'Ex2000'} ;
        PS>         default {
        PS>             $smsg = "UNRECOGNIZED ExchangeServer.AdminDisplayVersion.Major.Minor string:$($ExOPServer.version)! ABORTING!" ;
        PS>             write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
        PS>         }
        PS>     } ; 
        PS> }else {
        PS>     $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$ExOPServer.version:$($ExOPServer.version)!" ; 
        PS>     write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ; 
        PS>     throw $smsg ; 
        PS>     break ; 
        PS> } ; 
        Demo of parsing the returned Version property, into the proper Exchange Server revision.      
        .LINK
        https://github.com/tostka/verb-XXX
        .LINK
        https://bitbucket.org/tostka/powershell/
        .LINK
        http://mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
        .LINK
        https://github.com/SammyKrosoft/Search-AD-Using-Plain-PowerShell/blob/master/Get-SKExchangeServers.psm1
        .LINK
        https://github.com/tostka/verb-Ex2010
        .LINK
        https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/
        #>
        [CmdletBinding()]
        [Alias('Get-ExchangeServerInSite')]
        PARAM(
            [Parameter(Position=0,HelpMessage="Array of Server name strings to be filtered against[-Server 'SERVER1','SERVER2']")]
                [string[]]$Server,
            [Parameter(Position=1,HelpMessage="Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-SiteName 'SITENAME']")]
                [Alias('Site')]
                [string]$SiteName = (gv -name "$($TenOrg)_ADSiteDefaultName" -ea 0).Value,
            [Parameter(Position=2,HelpMessage="Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']")]
                [ValidateSet('MBX','CAS','HUB','UM','MBX','EDGE')]
                [string[]]$RoleNames = @('HUB','CAS'),
            [Parameter(HelpMessage="Array of msExchCurrentServerRoles 'role' integers to be filtered against (2|4|16|20|32|36|38|54|64|16385|16439)[-Roles @(38,16385)]")]
                [ValidateSet(2,4,16,20,32,36,38,54,64,16385,16439)]
                [int[]]$Roles,
            [Parameter(HelpMessage="Switch to suppress default 'pingable' test (e.g. returns all matches, no testing)[-NoTest]")]
                [Alias('NoPing')]
                [switch]$NoTest,
            [Parameter(HelpMessage="Milliseconds of max timeout to wait during port 80 test (defaults 100)[-SpeedThreshold 500]")]
                [int]$SpeedThreshold=100,
            [Parameter(Mandatory=$FALSE,HelpMessage="Tenant Tag (3-letter abbrebiation - defaults to global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']")]
                [ValidateNotNullOrEmpty()]
                [string]$TenOrg = $global:o365_TenOrgDefault,
            [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials[-Credentials [credential object]]")]
                [System.Management.Automation.PSCredential]$Credential
        ) ;
        BEGIN{
            ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
            $Verbose = ($VerbosePreference -eq 'Continue') ;
            $_sBnr="#*======v $(${CmdletName}): v======" ;
            $smsg = $_sBnr ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 }
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
        PROCESS{
            TRY{
                $configNC = ([ADSI]"LDAP://RootDse").configurationNamingContext ;
                $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC") ;
                $bLocalEdge = $false ; 
                if($Sitename -eq $env:COMPUTERNAME){
                    $smsg = "`$SiteName -eq `$env:COMPUTERNAME:$($SiteName):$($env:COMPUTERNAME)" ; 
                    $smsg += "`nThis computer appears to be an EdgeRole system (non-ADConnected)" ; 
                    $smsg += "`n(Blanking `$sitename and continuing discovery)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    #$bLocalEdge = $true ; 
                    $SiteName = $null ; 
                
                } ; 
                If($siteName){
                    $smsg = "Getting Site: $siteName" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    $objectClass = "objectClass=site" ;
                    $objectName = "name=$siteName" ;
                    $search.Filter = "(&($objectClass)($objectName))" ;
                    $site = ($search.Findall()) ;
                    $siteDN = ($site | select -expand properties).distinguishedname  ;
                } else {
                    $smsg = "(No -Site specified, resolving site from local machine domain-connection...)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                    else{ write-host -foregroundcolor green "$($smsg)" } ;
                    TRY{$siteDN = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().GetDirectoryEntry().distinguishedName}
                    CATCH [System.Management.Automation.MethodInvocationException]{
                        $ErrTrapd=$Error[0] ;
                        if(($ErrTrapd.Exception -match 'The computer is not in a site.') -AND $env:ExchangeInstallPath){
                            $smsg = "$($env:computername) is non-ADdomain-connected" ;
                            $smsg += "`nand has `$env:ExchangeInstalled populated: Likely Edge Server" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                            else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $vers = (get-item "$($env:ExchangeInstallPath)\Bin\Setup.exe").VersionInfo.FileVersionRaw ; 
                            $props = @{
                                Name=$env:computername;
                                FQDN = ([System.Net.Dns]::gethostentry($env:computername)).hostname;
                                Version = "Version $($vers.major).$($vers.minor) (Build $($vers.Build).$($vers.Revision))" ; 
                                #"$($vers.major).$($vers.minor)" ; 
                                #$exServer.serialNumber[0];
                                Roles = [System.Object[]]64 ;
                                RoleNames = @('EDGE');
                                DistinguishedName =  "CN=$($env:computername),CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=First Organization,CN=Microsoft Exchange,CN=Services,CN=Configuration,CN={nnnnnnnn-FAKE-GUID-nnnn-nnnnnnnnnnnn}" ;
                                Site = [System.Object[]]'NOSITE'
                                ResponseTime = if($rsp){$rsp.ResponseTime} else { 0} ;
                                NOTE = "This summary object, returned for a non-AD-connected EDGE server, *approximates* what would be returned on an AD-connected server" ;
                            } ;
                        
                            $smsg = "(-NoTest:Defaulting Fast:`$true)" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $props.add('Fast',$true) ;
                        
                            return (New-Object -TypeName PsObject -Property $props) ;
                        }elseif(-not $env:ExchangeInstallPath){
                            $smsg = "Non-Domain Joined machine, with NO ExchangeInstallPath e-vari: `nExchange is not installed locally: local computer resolution fails:`nPlease specify an explicit -Server, or -SiteName" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $false | write-output ;
                        } else {
                            $smsg = "$($env:computername) is both NON-Domain-joined -AND lacks an Exchange install (NO ExchangeInstallPath e-vari)`nPlease specify an explicit -Server, or -SiteName" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $false | write-output ;
                        };
                    } CATCH {
                        $siteDN =$ExOPsiteDN ;
                        write-warning "`$siteDN lookup FAILED, deferring to hardcoded `$ExOPsiteDN string in infra file!" ;
                    } ;
                } ;
                $smsg = "Getting Exservers in Site:$($siteDN)" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
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
                [void] $search.PropertiesToLoad.Add("serialNumber") ;
                [void] $search.PropertiesToLoad.Add("DistinguishedName") ;
                $exchServers = $search.FindAll() ;
                $Aggr = @() ;
                foreach($exServer in $exchServers){
                    $fqdn = ($exServer.Properties.networkaddress |
                        Where-Object{$_ -match '^ncacn_ip_tcp:'}).split(':')[1] ;
                    if($NoTest){} else {
                        $rsp = test-connection $fqdn -count 1 -ea 0 ;
                    } ;
                    $props = @{
                        Name = $exServer.Properties.name[0]
                        FQDN=$fqdn;
                        Version = $exServer.Properties.serialnumber
                        Roles = $exserver.Properties.msexchcurrentserverroles
                        RoleNames = $null ;
                        DistinguishedName = $exserver.Properties.distinguishedname;
                        Site = @("$($exserver.Properties.msexchserversite -Replace '^CN=|,.*$')") ;
                        ResponseTime = if($rsp){$rsp.ResponseTime} else { 0} ;
                    } ;
                    $props.RoleNames = switch ($exserver.Properties.msexchcurrentserverroles){
                        2       {"MBX"}
                        4       {"CAS"}
                        16      {"UM"}
                        20      {"CAS;UM".split(';')}
                        32      {"HUB"}
                        36      {"CAS;HUB".split(';')}
                        38      {"CAS;HUB;MBX".split(';')}
                        54      {"MBX"}
                        64      {"EDGE"}
                        16385   {"CAS"}
                        16439   {"CAS;HUB;MBX".split(';')}
                    }
                    if($NoTest){
                        $smsg = "(-NoTest:Defaulting Fast:`$true)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $props.add('Fast',$true) ;
                    }else {
                        $props.add('Fast',[boolean]($rsp.ResponseTime -le $SpeedThreshold)) ;
                    };
                    $Aggr += New-Object -TypeName PsObject -Property $props ;
                } ;
                $httmp = @{} ;
                if($Roles){
                    [regex]$rgxRoles = ('(' + (($roles |%{[regex]::escape($_)}) -join '|') + ')') ;
                    $matched =  @( $aggr | ?{$_.Roles -match $rgxRoles}) ;
                    foreach($m in $matched){
                        if($httmp[$m.name]){} else {
                            $httmp[$m.name] = $m ;
                        } ;
                    } ;
                } ;
                if($RoleNames){
                    foreach ($RoleName in $RoleNames){
                        $matched = @($Aggr | ?{$_.RoleNames -contains $RoleName} ) ;
                        foreach($m in $matched){
                            if($httmp[$m.name]){} else {
                                $httmp[$m.name] = $m ;
                            } ;
                        } ;
                    } ;
                } ;
                if($Server){
                    foreach ($Name in $Server){
                        $matched = @($Aggr | ?{$_.Name -eq $Name} ) ;
                        foreach($m in $matched){
                            if($httmp[$m.name]){} else {
                                $httmp[$m.name] = $m ;
                            } ;
                        } ;
                    } ;
                } ;
                if(($httmp.Values| measure).count -gt 0){
                    $Aggr  = $httmp.Values ;
                } ;
                $smsg = "Returning $((($Aggr|measure).count|out-string).trim()) match summaries to pipeline..." ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                $Aggr | write-output ;
            }CATCH{
                Write-Error $_ ;
            } ;
        } ;
        END{
            $smsg = "$($_sBnr.replace('=v','=^').replace('v=','^='))" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 }
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
    } ;
#}
#endregion GET_ADEXCHANGESERVERTDO ;#*------^ END Function get-ADExchangeServerTDO ^------ ;