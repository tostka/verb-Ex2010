#*------v Connect-Ex2010.ps1 v------
Function Connect-Ex2010 {
  <#
    .SYNOPSIS
    Connect-Ex2010 - Setup Remote ExchOnPrem Mgmt Shell connection (validated functional Exch2010 - Exch2016)
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    AddedCredit : Inspired by concept code by ExactMike Perficient, Global Knowl... (Partner)
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Version     : 1.1.0
    CreatedDate : 2020-02-24
    FileName    : Connect-Ex2010()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell
    REVISIONS   :
    * 1:30 PM 9/5/2024 added  update-SecurityProtocolTDO() SB to begin
    * 3:11 PM 7/15/2024 needed to change CHKPREREQ to check for presence of prop, not that it had a value (which fails as $false); hadn't cleared $MetaProps = ...,'DOESNTEXIST' ; confirmed cxo working non-based
    * 10:47 AM 7/11/2024 cleared debugging NoSuch etc meta tests
    * 1:34 PM 6/21/2024 ren $Global:E10Sess -> $Global:EXOPSess ; add: prereq checks, and $isBased support, to devert into most connect-exchangeServerTDO, get-ADExchangeServerTDO 100% generic fall back support (including buffering in the pair of funcs)
    # 9:43 AM 7/27/2021 revised -PSTitleBar to support suffix EMS[ctl]
    # 1:31 PM 7/21/2021 revised Add-PSTitleBar $sTitleBarTag with TenOrg spec (for prompt designators)
    # 3:18 PM 5/18/2021 somehow lost $credOpTORSID, so flipped lost default $credOPTor -> $credTORSID
    * 11:40 AM 5/14/2021 added -ea 0 to the gv tests (suppresses not-found error when called without logging config)
    * 11:22 AM 4/21/2021 coded around recent 'verbose the heck out of everything', yanked 99% of the verbose support - this seldom fails in a way that you need verbose, and when it's on, every cmdlet in the modules get echo'd, spams the heck out of console & logging. One key change (not sure if source) was to switch from inline import-pss & import-mod, into 2 steps with varis.
    * 10:02 AM 4/12/2021 add alias connect-ExOP (eventually rename verb-ex2010 to verb-exOnPrem)
    * 12:06 PM 4/2/2021 added alias cxOP ; added explicit echo on import-session|module, removed redundant catch block; added trycatch around import-sess|mod ; added recStatus support
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods ; renamed-standardized splat names (EMSSplat ->pltNSess ; ) ; flipped prefix into splat add ;
    * 2:36 PM 3/23/2021 getting away from dyn, random from array in $XXXMeta.Ex10Server, doesn't rely on AD lookups for referrals
    * 10:14 AM 3/23/2021 flipped default $Cred spec, pointed at an OP cred (matching reconnect-ex2010())
    * 11:36 AM 3/5/2021 updated colorcode, subed wv -verbose with just write-verbose, added cred.uname echo
    * 1:15 PM 3/1/2021 added org-level color-coded console
    * 3:28 PM 2/17/2021 updated to support cross-org, leverages new $XXXMeta.ExRevision, ExViewForest
    * 5:16 PM 10/22/2020 switched to no-loop meta lookup; debugged, fixed
    * 7:13 AM 7/22/2020 replaced codeblock w get-TenantTag(), flipped ExAdmin fr switch to un-typed
    * 5:11 PM 7/21/2020 added VEN support
    * 12:20 PM 5/27/2020 moved aliases: Add-EMSRemote,cx10 win func
    * 10:13 AM 5/15/2020 with vpn AD Ex lookup issue, patched in backup pass of get-ExchangeServerFromExGroup, in case of fail ; added failthrough to updated get-ExchangeServerFromExGroup, and finally to profile $smtpserver
    * 10:19 AM 2/24/2020 Connect-Ex2010/-OBS v1.1.0: updated cx10 to reflect infra file cred name change: cred####SID -> cred###SID, debugged, working, updated output banner to draw from global session, rather than imported module (was blank output). Ren'ing this one to the primary vers, and the prior to -OBS. Changed attribution, other than function names & concept, none of the code really sources back to Mike's original any more.
    * 6:59 PM 1/15/2020 cleanup
    * 7:51 AM 12/5/2019 Connect-Ex2010:retooled $ExAdmin variant webpool support - now has detect in the server-pick logic, and on failure, it retries to the stock pool.
    * 8:55 AM 11/27/2019 expanded $Credential support to switch to torolab & - potentiall/uncfg'd - CMW mail infra. Fw seems to block torolab access (wtf)
    * # 7:54 AM 11/1/2017 add titlebar tag & updated example to test for pres of Add-PSTitleBar
    * 12:09 PM 12/9/2016 implented and debugged as part of verb-Ex2010 set
    * 2:37 PM 12/6/2016 ported to local EMSRemote
    * 2/10/14 posted version
    $Credential can leverage a global: $Credential = $global:SIDcred
    .DESCRIPTION
    Connect-Ex2010 - Setup Remote Exch2010 Mgmt Shell connection
    This supports Non-Restricted IIS custom pools, which are created via create-EMSOpenRemotePool.ps1
    .PARAMETER  ExchangeServer
    Exch server to Remote to
    .PARAMETER  ExAdmin
    Use exadmin IIS WebPool for remote EMS[-ExAdmin]
    .PARAMETER  Credential
    Credential object
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    # -----------
    try{
        $reqMods="Connect-Ex2010;Reconnect-Ex2010;Disconnect-Ex2010;Get-ExchangeServerInSite;Disconnect-PssBroken;Add-PSTitleBar".split(";") ;
        $reqMods | % {if( !(test-path function:$_ ) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing $($_) function. EXITING." } } ;
        Reconnect-Ex2010 ;
    } CATCH {
        Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
        Exit #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ;
    } ;

    # -----------
    .EXAMPLE
    # -----------
    $rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" ;
    $rgxRemsPssName="^(Exchange2010|Session\sfor\simplicit\sremoting\smodule\sat\s.*)" ;
    $rgxSnapPssname="^Session\d{1}$" ;
    $rgxEx2010SnapinName="^Microsoft\.Exchange\.Management\.PowerShell\.E2010$";
    $Ex2010SnapinName="Microsoft.Exchange.Management.PowerShell.E2010" ;
    $Error.Clear() ;
    TRY {
    if(($host.version.major -lt 3) -AND (get-service MSExchangeADTopology -ea SilentlyContinue)){
        if (!(Get-PSSnapin | where {$_.Name -match $rgxEx2010SnapinName})) {Add-PSSnapin $Ex2010SnapinName -ea Stop} ;
            write-verbose -verbose:$bshowVerbose  "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Using Local Server EMS10 Snapin" ;
            $Global:E10IsDehydrated=$false ;
        } else {
            $reqMods="Connect-Ex2010;Reconnect-Ex2010;Disconnect-Ex2010;Get-ExchangeServerInSite;Disconnect-PssBroken;Cleanup;Add-PSTitleBar;Remove-PSTitleBar".split(";") ;
            $reqMods | % {if( !(test-path function:$_ ) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing $($_) function. EXITING." } } ;
            if(!(Get-PSSession |?{$_.ComputerName -match "^(adl|spb|lyn|bcc)ms\d{3}\.global\.ad\.toro\.com$" -AND $_.ConfigurationName -eq "Microsoft.Exchange" -AND $_.Name -eq "Exchange2010" -AND $_.State -eq "Opened" -AND $_.Availability -eq "Available"})){
    reconnect-Ex2010 ;
            $Global:E10IsDehydrated=$true ;
        } else {
          write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Functional REMS connection found. " ;
        } ;
    } ;
    get-exchangeserver | out-null ;
    # -----------
    More detailed REMS & server-EMS snapin coexistince version.
    .EXAMPLE
    # -----------
    if(!(Get-PSSnapin | where {$_.Name -match $rgxEx2010SnapinName})){
        Do {
            write-host "." -NoNewLine;Start-Sleep -m (1000 * 5)
            if( !(Get-PSSession|?{$_.Name -match $rgxRemsPssName -AND $_.ComputerName -match $rgxProdEx2010ServersFqdn -AND $_.State -eq 'Opened' -AND $_.Availability -eq 'Available'}) ){
                    Reconnect-Ex2010 ;
            } ;
        } Until ((Get-PSSession|?{($_.Name -match $rgxRemsPssName -AND $_.ComputerName -match $rgxProdEx2010ServersFqdn) -AND ($_.State -eq 'Opened') -AND ($_.Availability -eq 'Available')}))
    } ;
    # -----------
    Looping reconnect test example ; defers to existing Snapin (which should be self-maintaining)
    .LINK
    https://github.com/tostka/verb-Ex2010/
    #>
    #[CmdletBinding()] # 10:03 AM 4/21/2021 disable, see if it kills verbose
    [Alias('Add-EMSRemote','cx10','cxOP','connect-ExOP')]
    Param(
        [Parameter(Position = 0, HelpMessage = "Exch server to Remote to")]
            [string]$ExchangeServer,
        [Parameter(HelpMessage = 'Use exadmin IIS WebPool for remote EMS[-ExAdmin]')]
            $ExAdmin,
        [Parameter(HelpMessage = 'Credential object')][System.Management.Automation.PSCredential]
            $Credential = $credTORSID
    )  ;
    BEGIN{
        #$verbose = ($VerbosePreference -eq "Continue") ;
		$CurrentVersionTlsLabel = [Net.ServicePointManager]::SecurityProtocol ; # Tls, Tls11, Tls12 ('Tls' == TLS1.0)  ;
        write-verbose "PRE: `$CurrentVersionTlsLabel : $($CurrentVersionTlsLabel )" ;
        # psv6+ already covers, test via the SslProtocol parameter presense
        if ('SslProtocol' -notin (Get-Command Invoke-RestMethod).Parameters.Keys) {
            $currentMaxTlsValue = [Math]::Max([Net.ServicePointManager]::SecurityProtocol.value__,[Net.SecurityProtocolType]::Tls.value__) ;
            write-verbose "`$currentMaxTlsValue : $($currentMaxTlsValue )" ;
            $newerTlsTypeEnums = [enum]::GetValues('Net.SecurityProtocolType') | Where-Object { $_ -gt $currentMaxTlsValue }
            if($newerTlsTypeEnums){
                write-verbose "Appending upgraded/missing TLS `$enums:`n$(($newerTlsTypeEnums -join ','|out-string).trim())" ;
            } else {
                write-verbose "Current TLS `$enums are up to date with max rev available on this machine" ;
            };
            $newerTlsTypeEnums | ForEach-Object {
                [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $_
            } ;
        } ;
        
        #region CHKPREREQ ; #*------v CHKPREREQ v------
        # critical dependancy Meta variables
        $MetaNames = ,'TOR','CMW','TOL' #,'NOSUCH' ; 
        # critical dependancy Meta variable properties
        $MetaProps = 'Ex10Server','Ex10WebPoolVariant','ExRevision','ExViewForest','ExOPAccessFromToro','legacyDomain' #,'DOESNTEXIST' ; 
        # critical dependancy parameters
        $gvNames = 'Credential' 
        $isBased = $true ; $gvMiss = @() ; $ppMiss = @() ; 
        foreach($met in $metanames){
            write-verbose "chk:`$$($met)Meta" ; 
            if(-not (gv -name "$($met)Meta" -ea 0)){$isBased = $false; $gvMiss += "$($met)Meta" } ; 
            if($MetaProps){
                foreach($mp in $MetaProps){
                    write-verbose "chk:`$$($met)Meta.$($mp)" ; 
                    #if(-not (gv -name "$($met)Meta" -ea 0).value[$mp]){
                    if(-not (gv -name "$($met)Meta" -ea 0).value.keys -contains $mp){
                        $isBased = $false; $ppMiss += "$($met)Meta.$($mp)" 
                    } ; 
                } ; 
            } ; 
        } ; 
        if($gvNames){
            foreach($gvN in $gvNames){
                write-verbose "chk:`$$($gvN)" ; 
                if(-not (gv -name "$($gvN)" -ea 0)){$isBased = $false; $gvMiss += "$($gvN)" } ; 
            } ; 
        } ; 
        if($gvmiss){write-warning "Missing Dependant Meta variables:`n$(($gvMiss |%{"`$$($_)" }) -join ',')" } ; 
        if($ppMiss){write-warning "Missing Dependant Meta vari properties:`n$(($ppMiss |%{"`$$($_)" }) -join ',')" } ; 
        if(-not $isBased){ write-warning  "missing critical dependancy profile config!" } ; 
        #endregion CHKPREREQ ; #*------^ END CHKPREREQ ^------


        $sWebPoolVariant = "exadmin" ;
        $CommandPrefix = $null ;
        # use credential domain to determine target org
        $rgxLegacyLogon = '\w*\\\w*' ;

        #region CONNEXOPTDO ; #*------v  v------
        #*------v Function Connect-ExchangeServerTDO v------
        #if(-not(get-command Connect-ExchangeServerTDO -ea SilentlyContinue)){
            Function Connect-ExchangeServerTDO {
                <#
                .SYNOPSIS
                Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
                will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellREmote (REMS) connect to each server, 
                stopping at the first successful connection.
                .NOTES
                Version     : 3.0.3
                Author      : Todd Kadrie
                Website     : http://www.toddomation.com
                Twitter     : @tostka / http://twitter.com/tostka
                CreatedDate : 2024-05-30
                FileName    : Connect-ExchangeServerTDO.ps1
                License     : (none-asserted)
                Copyright   : (none-asserted)
                Github      : https://github.com/tostka/verb-Ex2010
                Tags        : Powershell, ActiveDirectory, Exchange, Discovery
                AddedCredit : Brian Farnsworth
                AddedWebsite: https://codeandkeep.com/
                AddedTwitter: URL
                AddedCredit : David Paulson
                AddedWebsite: https://techcommunity.microsoft.com/t5/exchange-team-blog/exchange-health-checker-has-a-new-home/ba-p/2306671
                AddedTwitter: URL
                REVISIONS
                * 12:49 PM 6/21/2024 flipped PSS Name to Exchange$($ExchVers[dd])
                * 12:57 PM 6/11/2024 Validated, Ex2010 & Ex2019, hub, mail & edge roles: tested ☑️ on CMW mail role (Curly); and Jumpbox; 
                    copied in CBH from repo copy, which has been updated/debugged compat on CMW Edge 
                    includes local snapin detect & load for edge role (simplest EMS load option for Edge role, from David Paulson's original code; no longer published with Ex2010 compat)
                * 11:28 AM 5/30/2024 fixed failure to recognize existing functional PSSession; Made substantial update in logic, validate works fine with other orgs, and in our local orgs.
                * 4:02 PM 8/28/2023 debuged, updated CBH, renamed connect-ExchangeSErver -> Connect-ExchangeServerTDO (avoid name clashes, pretty common verb-noun combo).
                * 12:36 PM 8/24/2023 init

                .DESCRIPTION
                Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
                will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellRemote (REMS) connect to each server, 
                stopping at the first successful connection.

                Relies upon/requires get-ADExchangeServerTDO(), to return a descriptive summary of the Exchange server(s) revision etc, for connectivity logic.
                Supports Exchange 2010 through 2019, as implemented.
            
                Intent, as contrasted with verb-EXOP/Ex2010 is to have no local module dependancies, when running EXOP into other connected orgs, where syncing profile & supporting modules code can be problematic. 
                This uses native ADSI calls, which are supported by Windows itself, without need for external ActiveDirectory module etc.

                The particular approach inspired by BF's demo func that accompanied his take on get-adExchangeServer(), which I hybrided with my own existing code for cred-less connectivity. 
                I added get-OrganizationConfig testing, for connection pre/post confirmation, along with Exchange Server revision code for continutional handling of new-pssession remote powershell EMS connections.
                Also shifted connection code into _connect-EXOP() internal func.
                As this doesn't rely on local module presnece, it doesn't have to do the usual local remote/local invocation detection you'd do for non-dehydrated on-server EMS (more consistent this way, anyway; 
                there are only a few cmdlet outputs I'm aware of, that have fundementally broken returns dehydrated, and require local non-remote EMS use to function.

                My core usage would be to paste the function into the BEGIN{} block for a given remote org process, to function as a stricly local ad-hoc function.
                .PARAMETER name
                FQDN of a specific Exchange server[-Name EXSERVER.DOMAIN.COM]
                .PARAMETER discover
                Boolean paraameter that drives auto-discovery of target Exchange servers for connection (defaults `$true)[-discover:`$false]
                .PARAMETER credential
                Use specific Credentials[-Credentials [credential object]
                    .PARAMETER Site
                Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-Site 'SITENAME']
                .PARAMETER RoleNames
                Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']
                .PARAMETER TenOrg
                Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
                .INPUTS
                None. Does not accepted piped input.(.NET types, can add description)
                .OUTPUTS
                [system.object] Returns a system object containing a successful PSSession
                System.Boolean
                [| get-member the output to see what .NET obj TypeName is returned, to use here]
                System.Array of System.Object's
                .EXAMPLE
                PS> $PSSession = Connect-ExchangeServerTDO -siteName SITENAME -RoleNames @('HUB','CAS') -verbose 
                Demo's connecting to a functional Hub or CAS server in the SITENAME site with verbose outputs, the `PSSession variable will contain information about the successful connection. Makes automatic Exchangeserver discovery calls into AD (using ADSI) leveraging the separate get-ADExchangeServerTDO()
                .EXAMPLE
                PS> TRY{$Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name}CATCH{$Site=$env:COMPUTERNAME} ;
                PS> $PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
                Demo including support for EdgeRole, which is detected on it's lack of AD Site specification (which gets fed through to call, by setting the Site to the machine itself).
                .LINK
                https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/
                .LINK
                https://github.com/Lucifer1993/PLtools/blob/main/HealthChecker.ps1
                .LINK
                https://microsoft.github.io/CSS-Exchange/Diagnostics/HealthChecker/
                .LINK
                https://bitbucket.org/tostka/powershell/
                .LINK
                https://github.com/tostka/verb-Ex2010
                #>        
                [CmdletBinding(DefaultParameterSetName='discover')]
                PARAM(
                    [Parameter(Position=0,Mandatory=$true,ParameterSetName='name',HelpMessage="FQDN of a specific Exchange server[-Name EXSERVER.DOMAIN.COM]")]
                        [String]$name,
                    [Parameter(Position=0,ParameterSetName='discover',HelpMessage="Boolean paraameter that drives auto-discovery of target Exchange servers for connection (defaults `$true)[-discover:`$false]")]
                        [bool]$discover=$true,
                    [Parameter(Position=1,HelpMessage = "Use specific Credentials[-Credentials [credential object]")]
                        [Management.Automation.PSCredential]$credential,
                    [Parameter(Position=1,HelpMessage="Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-Site 'SITENAME']")]
                        [Alias('Site')]
                        [string]$SiteName = (gv -name "$($TenOrg)_ADSiteDefaultName" -ea 0).Value,
                    [Parameter(Position=2,HelpMessage="Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']")]
                        [ValidateSet('MBX','CAS','HUB','UM','MBX','EDGE')]
                        [string[]]$RoleNames = @('HUB','CAS'),
                    [Parameter(Mandatory=$FALSE,HelpMessage="Tenant Tag (3-letter abbrebiation - defaults to global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']")]
                        [ValidateNotNullOrEmpty()]
                        [string]$TenOrg = $global:o365_TenOrgDefault
                ) ;
                BEGIN{
                    $Verbose = ($VerbosePreference -eq 'Continue') ;
                    write-verbose "#*------v Function _connect-ExOP v------" ;
                    function _connect-ExOP{
                        [CmdletBinding()]
                        PARAM(
                            [Parameter(Position=0,Mandatory=$true,HelpMessage="Exchange server AD Summary system object[-Server EXSERVER.DOMAIN.COM]")]
                                [system.object]$Server,
                            [Parameter(Position=1,HelpMessage = "Use specific Credentials[-Credentials [credential object]")]
                                [Management.Automation.PSCredential]$credential
                        );
                        $verbose = $($VerbosePreference -eq "Continue") ;
                        if([double]$ExVersNum = [regex]::match($Server.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
                            switch -regex ([string]$ExVersNum) {
                                '15.2' { $isEx2019 = $true ; $ExVers = 'Ex2019' }
                                '15.1' { $isEx2016 = $true ; $ExVers = 'Ex2016'}
                                '15.0' { $isEx2013 = $true ; $ExVers = 'Ex2013'}
                                '14.*' { $isEx2010 = $true ; $ExVers = 'Ex2010'}
                                '8.*' { $isEx2007 = $true ; $ExVers = 'Ex2007'}
                                '6.5' { $isEx2003 = $true ; $ExVers = 'Ex2003'}
                                '6' {$isEx2000 = $true ; $ExVers = 'Ex2000'} ;
                                default {
                                    $smsg = "UNRECOGNIZED ExVersNum.Major.Minor string:$($ExVersNum)! ABORTING!" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    THROW $SMSG ;
                                    BREAK ;
                                }
                            } ;
                        }else {
                            $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$Server.version:$($Server.version)!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            throw $smsg ;
                            break ;
                        } ;
                        if($Server.RoleNames -eq 'EDGE'){
                            if(($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup')) -or
                                ($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup')) -or
                                $ByPassLocalExchangeServerTest)
                            {
                                if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole') -or
                                     (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'))
                                {
                                    write-verbose ("We are on Exchange Edge Transport Server")
                                    $IsEdgeTransport = $true
                                }
                                TRY {
                                    Get-ExchangeServer -ErrorAction Stop | Out-Null
                                    write-verbose "Exchange PowerShell Module already loaded."
                                    $passed = $true 
                                }CATCH {
                                    write-verbose ("Failed to run Get-ExchangeServer")
                                    if($isLocalExchangeServer){
                                        write-host  "Loading Exchange PowerShell Module..."
                                        TRY{
                                            if($IsEdgeTransport){
                                                # implement local snapins access on edge role: Only way to get access to EMS commands.
                                                [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exshell.psc1" -ErrorAction Stop
                                                ForEach($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn){
                                                    write-verbose ("Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name)
                                                    Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop
                                                } ; 
                                                Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop ; 
                                                $passed = $true #We are just going to assume this passed.
                                            }else{
                                                Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                                                Connect-ExchangeServer -Auto -ClientApplication:ManagementShell
                                                $passed = $true #We are just going to assume this passed.
                                            } 
                                        }CATCH {
                                            write-host ("Failed to Load Exchange PowerShell Module...")
                                        }                               
                                    } ;
                                } FINALLY {
                                    if($LoadExchangeVariables -and $passed -and $isLocalExchangeServer){
                                        if($ExInstall -eq $null -or $ExBin -eq $null){
                                            if(Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup'){
                                                $Global:ExInstall = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath
                                            }else{
                                                $Global:ExInstall = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath
                                            }
            
                                            $Global:ExBin = $Global:ExInstall + "\Bin"
            
                                            write-verbose ("Set ExInstall: {0}" -f $Global:ExInstall)
                                            write-verbose ("Set ExBin: {0}" -f $Global:ExBin)
                                        }
                                    }
                                }
                            } else  {
                                write-verbose ("Does not appear to be an Exchange 2010 or newer server.")
                            }
                            if(get-command -Name Get-OrganizationConfig -ea 0){
                                $smsg = "Running in connected/Native EMS" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                Return $true ; 
                            } else { 
                                TRY{
                                    $smsg = "Initiating Edge EMS local session (exshell.psc1 & exchange.ps1)" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                    # 5;36 PM 5/30/2024 didn't work, went off to nowhere for a long time, and exited the script
                                    #& (gcm powershell.exe).path -PSConsoleFile "$($env:ExchangeInstallPath)bin\exshell.psc1" -noexit -command ". '$($env:ExchangeInstallPath)bin\Exchange.ps1'"
                                    <# [Adding the Transport Server to Exchange - Mark Lewis Blog](https://marklewis.blog/2020/11/19/adding-the-transport-server-to-exchange/)
                                    To access the management console on the transport server, I opened PowerShell then ran
                                    exshell.psc1
                                    Followed by
                                    exchange.ps1
                                    At this point, I was able to create a new subscription using he following PowerShel
                                    #>
                                    invoke-command exshell.psc1 ; 
                                    invoke-command exchange.ps1
                                    if(get-command -Name Get-OrganizationConfig -ea 0){
                                        $smsg = "Running in connected/Native EMS" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                        Return $true ;
                                    } else { return $false };  
                                }CATCH{
                                    Write-Error $_ ;
                                } ;
                            } ; 
                        } else {
                            $pltNPSS=@{ConnectionURI="http://$($Server.FQDN)/powershell"; ConfigurationName='Microsoft.Exchange' ; name="Exchange$($ExVersNum.tostring())"} ;
                            # use ExVersUnm dd instead of hardcoded (Exchange2010)
                            if($ExVersNum -ge 15){
                                write-verbose "EXOP.15+:Adding -Authentication Kerberos" ;
                                $pltNPSS.add('Authentication',"Kerberos") ;
                                $pltNPSS.name = $ExVers ;
                            } ;
                            $smsg = "Adding EMS (connecting to $($Server.FQDN))..." ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $smsg = "New-PSSession w`n$(($pltNPSS|out-string).trim())" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $ExPSS = New-PSSession @pltNPSS  ;
                            $ExIPSS = Import-PSSession $ExPSS -allowclobber ;
                            $ExPSS | write-output ;
                            $ExPSS= $ExIPSS = $null ;
                        } ; 
                    } ;
                    write-verbose "#*------^ END Function _connect-ExOP ^------" ;
                    $pltGADX=@{
                        ErrorAction='Stop';
                    } ;
                } ;
                PROCESS{
                    if($PSBoundParameters.ContainsKey('credential')){
                        $pltGADX.Add('credential',$credential) ;
                    }
                    if($SiteName){
                        $pltGADX.Add('siteName',$siteName) ;
                    } ;
                    if($RoleNames){
                        $pltGADX.Add('RoleNames',$RoleNames) ;
                    } ;
                    TRY{
                        if($discover){
                            $smsg = "Getting list of Exchange Servers" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $exchServers=get-ADExchangeServerTDO @pltGADX | sort responsetime ;
                        }else{
                            $exchServers=get-ADExchangeServerTDO @pltGADX | sort responsetime ;
                        } ;
                        $pltTW=@{
                            'ErrorAction'='Stop';
                        } ;
                        $pltCXOP = @{
                            verbose = $($VerbosePreference -eq "Continue") ;
                        } ;
                        if($pltGADX.credential){
                            $pltCXOP.Add('Credential',$pltCXOP.Credential) ;
                        } ;
                        $prpPSS = 'Id','Name','ComputerName','ComputerType','State','ConfigurationName','Availability' ; 
                        foreach($exServer in $exchServers){
                            write-verbose "testing conn to:$($exServer.name.tostring())..." ; 
                            if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig' -ea SilentlyContinue){
                                if($pssEXOP = Get-PSSession |  where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ( $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND ( $_.runspace.ConnectionInfo.Port -eq '80') }){
                                    if($pssEXOP.State -ne "Opened" -OR $pssEXOP.Availability -ne "Available"){
                                        $pssEXOP | remove-pssession ; $pssEXOP = $null ;
                                    } ;
                                } ; 
                            } else {
                                $smsg = "(mangled ExOP conn: disconnect/reconnect...)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                if($pssEXOP = Get-PSSession |  where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ( $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND ( $_.runspace.ConnectionInfo.Port -eq '80') }){
                                    if($pssEXOP.State -ne "Opened" -OR $pssEXOP.Availability -ne "Available"){
                                        $pssEXOP | remove-pssession ; $pssEXOP = $null ;
                                    } ;
                                } ; 
                            } ;
                            if(-not $pssEXOP){
                                $smsg = "Connecting to: $($exServer.FQDN)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                if($NoTest){
                                    $ExPSS =$ExPSS = _connect-ExOP @pltCXOP -Server $exServer
                               } else {
                                    TRY{
                                        $smsg = "Testing Connection: $($exServer.FQDN)" ;
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                        If(test-connection $exServer.FQDN -count 1 -ea 0) {
                                            $smsg = "confirmed pingable..." ;
                                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                        } else {
                                            $smsg = "Unable to Ping $($exServer.FQDN)" ; ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        } ;
                                        $smsg = "Testing WinRm: $($exServer.FQDN)" ;
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                        $winrm=Test-WSMan @pltTW -ComputerName $exServer.FQDN ;
                                        if($winrm){
                                            $ExPSS = _connect-ExOP @pltCXOP -Server $exServer;
                                        } else {
                                            $smsg = "Unable to Test-WSMan $($exServer.FQDN) (skipping)" ; ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        } ;
                                    }CATCH{
                                        $errMsg="Server: $($exServer.FQDN)] $($_.Exception.Message)" ;
                                        Write-Error -Message $errMsg ;
                                        continue ;
                                    } ;
                                };
                            } else {
                                $smsg = "$((get-date).ToString('HH:mm:ss')):Accepting first valid connection w`n$(($pssEXOP | ft -a $prpPSS|out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $ExPSS = $pssEXOP ; 
                                break ; 
                            }  ;
                        } ;
                    }CATCH{
                        Write-Error $_ ;
                    } ;
                } ;
                END{
                    if(-not $ExPSS){
                        $smsg = "NO SUCCESSFUL CONNECTION WAS MADE, WITH THE SPECIFIED INPUTS!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $smsg = "(returning `$false to the pipeline...)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        return $false
                    } else{
                        if($ExPSS.State -eq "Opened" -AND $ExPSS.Availability -eq "Available"){
                            if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                                $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ;
                                throw $smsg ;
                                $smsg | write-warning  ;
                            } else {
                                $smsg = "(connected to EXOP.Org:$($orgName))" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                            return $ExPSS
                        } ;
                    } ; 
                } ;
            } ;
        #} ; 
        #*------^ END Function Connect-ExchangeServerTDO ^------
        #endregion CONNEXOPTDO ; #*------^ END CONNEXOPTDO ^------
    
        #region GADEXSERVERTDO ; #*------v  v------
        #*------v Function get-ADExchangeServerTDO v------
        #if(-not(get-command get-ADExchangeServerTDO -ea SilentlyContinue)){
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
                            $smsg = "WVGetting Site: $siteName" ;
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
        #} ; 
        #*------^ END Function get-ADExchangeServerTDO ^------ ;
        #endregion GADEXSERVERTDO ; #*------^ END GADEXSERVERTDO ^------

        if($isBased){
            $TenOrg = get-TenantTag -Credential $Credential ;
        } else { 

        } ; 
        <#
        if($Credential.username -match $rgxLegacyLogon){
            $credDom =$Credential.username.split('\')[0] ;
        } elseif ($Credential.username.contains('@')){
            $credDom = ($Credential.username.split("@"))[1] ;
        } else {
            write-warning "$((get-date).ToString('HH:mm:ss')):UNRECOGNIZED CREDENTIAL!:$($Credential.Username)`nUNABLE TO RESOLVE DEFAULT EX10SERVER FOR CONNECTION!" ;
        } ;
        #>
    } ;  # BEG-E
    PROCESS{
        if($isBased){
            $ExchangeServer=$null ;
            # flip from dyn lookup to array in Ex10Server, and always use get-random to pick between. Returns a value, even when only a single value
            $ExchangeServer = (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server|get-random ;
            $ExAdmin = (Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant ;
            $ExVers = (Get-Variable  -name "$($TenOrg)Meta").value.ExRevision ;
            $ExVwForest = (Get-Variable  -name "$($TenOrg)Meta").value.ExViewForest ;
            $ExOPAccessFromToro = (Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro
            # force unresolved to dyn
            if(!$ExchangeServer){
                #$ExchangeServer = 'dynamic' ;
                # getting away from dyn, random from array in Ex10Server
                throw "Undefined `$ExchangeServer for $($TenOrg) org, and `$$($TenOrg)Meta.Ex10Server property" ;
                Exit ;
            } ;
            if($ExchangeServer -eq 'dynamic'){
                if( $ExchangeServer = (Get-ExchangeServerInSite | Where-Object { ($_.roles -eq 36) } | Get-Random ).FQDN){}
                else {
                    write-warning "$((get-date).ToString('HH:mm:ss')):Get-ExchangeServerInSite *FAILED*,`ndeferring to Get-ExchServerFromExServersGroup" ;
                    if(!($ExchangeServer = Get-ExchServerFromExServersGroup)){
                        write-warning "$((get-date).ToString('HH:mm:ss')):Get-ExchServerFromExServersGroup *FAILED*,`n deferring to profile `$smtpserver:$($smtpserver))"  ;
                        $ExchangeServer = $smtpserver ;
                    };
                } ;
            } ;

            write-host -foregroundcolor darkgray "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Adding EMS (connecting to $($ExchangeServer))..." ;
            # splat to open a session - # stock 'PSLanguageMode=Restricted' powershell IIS Webpool
            $pltNSess = @{ConnectionURI = "http://$ExchangeServer/powershell"; ConfigurationName = 'Microsoft.Exchange' ; name = "Exchange$($ExVers)" } ;
            if($env:USERDOMAIN -ne (Get-Variable  -name "$($TenOrg)Meta").value.legacyDomain){
                # if not in the $TenOrg legacy domain - running cross-org -  add auth:Kerberos
                <#suppresses: The WinRM client cannot process the request. It cannot determine the content type of the HTTP response f rom the destination computer. The content type is absent or invalid
                #>
                $pltNSess.add('Authentication','Kerberos') ;
            } ;
            if ($ExAdmin) {
              # use variant IIS Webpool
              $pltNSess.ConnectionURI = $pltNSess.ConnectionURI.replace("/powershell", "/$($sWebPoolVariant)") ;
            }
            if ($Credential) {
                 $pltNSess.Add("Credential", $Credential)
                 write-verbose "(using cred:$($credential.username))" ;
            } ;

            # -Authentication Basic only if specif needed: for Ex configured to connect via IP vs hostname)
            # try catch against and retry into stock if fails
            $error.clear() ;
            TRY {
                $Global:EXOPSess = New-PSSession @pltNSess -ea STOP  ;
            } CATCH {
                $ErrTrapd = $_ ;
                write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                #-=-record a STATUSWARN=-=-=-=-=-=-=
                $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                if(Get-Variable passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                if(Get-Variable -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                #-=-=-=-=-=-=-=-=
                if ($ExAdmin) {
                    # switch to stock pool and retry
                    $pltNSess.ConnectionURI = $pltNSess.ConnectionURI.replace("/$($sWebPoolVariant)", "/powershell") ;
                    write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TARGETING EXADMIN POOL`nRETRY W STOCK POOL: New-PSSession w`n$(($pltNSess|out-string).trim())" ;
                    $Global:EXOPSess = New-PSSession @pltNSess -ea STOP  ;
                } else {
                    BREAK ;
                } ;
            } ;

            write-verbose "$((get-date).ToString('HH:mm:ss')):Importing Exchange 2010 Module" ;
            #$pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ; verbose=$false} ; # force verbose off, suppress spam in console
            # tear verbose out
            $pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ;} ;
            #$pltISess = [ordered]@{Session = $Global:EXOPSess ; DisableNameChecking = $true  ; AllowClobber = $true ; ErrorAction = 'Stop' ; Verbose = $false ;} ;
            $pltISess = [ordered]@{Session = $Global:EXOPSess ; DisableNameChecking = $true  ; AllowClobber = $true ; ErrorAction = 'Stop' ; } ;
            if($CommandPrefix){
                $pltIMod.add('Prefix',$CommandPrefix) ;
                $pltISess.add('Prefix',$CommandPrefix) ;
            } ;
            $smsg = "$((get-date).ToString('HH:mm:ss')):Import-PSSession  w`n$(($pltISess|out-string).trim())`nImport-Module w`n$(($pltIMod|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $error.clear() ;
            TRY {
                # 9:57 AM 4/21/2021 coming through full verbose, suppress the pref
                if($VerbosePreference -eq "Continue"){
                    $VerbosePrefPrior = $VerbosePreference ;
                    $VerbosePreference = "SilentlyContinue" ;
                    $verbose = ($VerbosePreference -eq "Continue") ;
                } ;
                #$Global:E10Mod = Import-Module (Import-PSSession @pltISess) @pltIMod   ;
                # try 2-stopping (suppress verbose)
                $xIPS = Import-PSSession @pltISess ;
                $Global:E10Mod = Import-Module $xIPS @pltIMod ;
                if($ExVwForest){
                    write-host "Setting EMS Session: Set-AdServerSettings -ViewEntireForest `$True" ;
                    Set-AdServerSettings -ViewEntireForest $True ;
                } ;
                # reenable VerbosePreference:Continue, if set, during mod loads
                if($VerbosePrefPrior -eq "Continue"){
                    $VerbosePreference = $VerbosePrefPrior ;
                    $verbose = ($VerbosePreference -eq "Continue") ;
                } ;
            } CATCH {
                $ErrTrapd = $_ ;
                write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                #-=-record a STATUSERROR=-=-=-=-=-=-=
                $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                if(Get-Variable passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                if(Get-Variable -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                #-=-=-=-=-=-=-=-=
            } ;
            # 7:54 AM 11/1/2017 add titlebar tag
            #Add-PSTitleBar 'EMS' ;
            # 1:31 PM 7/21/2021 build with TenOrg spec
            # 9:00 AM 7/27/2021 revise to support EMS[tlc] single-letter onprem conne designator (already in infra file OrgSvcs list).
            if($TenOrg){
                <# can't just use last char, lab varies from others
                if($TenOrg -ne 'TOL' ){
                    $sTitleBarTag = @("EMS$($TenOrg.substring(0,1).tolower())") ; # 1st char
                }else{
                    $sTitleBarTag = @("EMS$($TenOrg.substring(2,1).tolower())") ; # last char
                } ; 
                #>
                switch -regex ($TenOrg){
                    '^(CMW|TOR)$'{
                        $sTitleBarTag = @("EMS$($TenOrg.substring(0,1).tolower())") ; # 1st char
                    }
                    '^TOL$'{
                        $sTitleBarTag = @("EMS$($TenOrg.substring(2,1).tolower())") ; # last char
                    } ; 
                    default{
                        throw "$($TenOrg):unsupported `$TenOrg!" ; 
                        break ; 
                    }
                } ; 
            } else { 
                $sTitleBarTag = @("EMS") ;
            } ; 
            write-verbose "`$sTitleBarTag:$($sTitleBarTag)" ; 
        
            #$sTitleBarTag += $TenOrg ;
            Add-PSTitleBar $sTitleBarTag -verbose:$($VerbosePreference -eq "Continue")  ;
            # tag E10IsDehydrated
            $Global:E10IsDehydrated = $true ;
            write-host -foregroundcolor darkgray "`n$(($Global:EXOPSess | select ComputerName,Availability,State,ConfigurationName | format-table -auto |out-string).trim())" ;

        } else { 
            
            #region SERVICE_CONNECTIONS_DEPENDANCYLESS #*======v SERVICE_CONNECTIONS_DEPENDANCYLESS v======
            # DON'T RUN THIS AND THE USEXOP= BLOCK TOGETHER!
            # SIMPLE DEP-LESS VARIANT FOR EXOP-ONLY, NO DEPS ON VERB-*, OTHER THAN REQS: load-ADMs() & get-ADExchangeSErverTDO() (both should be local function includes)
            # PRETUNE STEERING separately *before* pasting in balance of region
            #*------v STEERING VARIS v------
            # CAN USE THIS BLOCK TO FORCE UPPER SERVICE_CONNECTIONS DRIVERS FALSE BEFORE USE HERE
            #$UseOP=$false ; 
            #$UseExOP=$false ;
            #$useForestWide = $false ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
            #$UseOPAD = $false ; 
            # ---
            $UseOPDYN=$true ; 
            $UseExOPDYN=$true ;
            $UseEXOPInSite=$true ;
            $useForestWideDYN = $false ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
            $UseOPADDYN = $true ; 
            $UseOPDYN = [boolean]($UseOPDYN -OR $UseExOPDYN -OR $UseOPADDYN) ;
            if($UseOPDYN -AND $UseOP){ write-warning "BOTH `$UseOPDYN -AND `$UseOP ARE TRUE!`nUSE ONE OR THE OTHER!" } ;
            if($UseExOPDYN -AND $UseExOP){ write-warning "BOTH `$UseExOPDYN -AND `$UseExOP ARE TRUE!`nUSE ONE OR THE OTHER!" } ;
            if($UseOPADDYN -AND $UseOPAD){ write-warning "BOTH `$UseOPADDYN -AND `$UseOPAD ARE TRUE!`nUSE ONE OR THE OTHER!" } ;
            if($useForestWideDYN -AND $useForestWide){ write-warning "BOTH `$useForestWideDYN -AND `$useForestWide ARE TRUE!`nUSE ONE OR THE OTHER!" } ;
            #*------^ END STEERING VARIS ^------
            #region GENERIC_EXOP_SRVR_CONN #*------v GENERIC_EXOP_SRVR_CONN BP v------
            if($UseOPDYN){
                #*------v GENERIC EXOP SRVR CONN BP v------
                # connect to ExOP 
                if($UseExOPDYN){
                    'get-ADExchangeSErverTDO','Connect-ExchangeServerTDO'  | foreach-object{
                        if(get-command $_ -ErrorAction STOP){} else { 
                            $smsg = "MISSING DEP FUNC:$($_)! must be either local include, or pre-loaded for EXOP connectivity to work!" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            throw $smsg ; 
                            break ; 
                        } ; 
                    } ; 
                    if($UseEXOPInSite){
                        TRY{
                            $Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name 
                        }CATCH{$Site=$env:COMPUTERNAME} ;
                        #$HubServers = get-ADExchangeSErverTDO -RoleNames 'HUB' -verbose
                        $PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
                    }else{
                        $PSSession = Connect-ExchangeServerTDO -RoleNames @('HUB','CAS') -verbose ; 
                    } ; 
                    # from get-ADExchangeSErverTDO return 
                    #if([double]$ExVersNum = [regex]::match($ExOPServer.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
                    # from Connect-ExchangeServerTDO pssession return
                    if([double]$ExVersNum = [regex]::match($PsSession.applicationprivatedata.supportedversions,'(\d+\.\d+)\.(\d+\.\d+)').groups[1].value){
                        switch -regex ([string]$ExVersNum) {
                            '15.2' { $isEx2019 = $true ; $ExVers = 'Ex2019' }
                            '15.1' { $isEx2016 = $true ; $ExVers = 'Ex2016'}
                            '15.0' { $isEx2013 = $true ; $ExVers = 'Ex2013'}
                            '14.*' { $isEx2010 = $true ; $ExVers = 'Ex2010'}
                            '8.*' { $isEx2007 = $true ; $ExVers = 'Ex2007'}
                            '6.5' { $isEx2003 = $true ; $ExVers = 'Ex2003'}
                            '6' {$isEx2000 = $true ; $ExVers = 'Ex2000'} ;
                            default {
                                $smsg = "UNRECOGNIZED ExVersNum.Major.Minor string:$($ExVersNum)! ABORTING!" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                THROW $SMSG ;
                                BREAK ;
                            }
                        } ;
                        $smsg = "`$ExVersNum: $($PsSession.applicationprivatedata.supportedversions)`n$((gv isex*| %{"`n`$$($_.name): `$$($_.value)"}|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    }else {
                        #$smsg = "UNABLE TO RESOLVE `$ExVersNum from `$ExOPServer.version:$($ExOPServer.version)!" ;
                        $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$PsSession.applicationprivatedata.supportedversions:$($PsSession.applicationprivatedata.supportedversions)!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        throw $smsg ;
                        break ;
                    } ; 
                    TRY{
                        if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig' -ea SilentlyContinue){
                            if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                                $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ; 
                                throw $smsg ; 
                                $smsg | write-warning  ; 
                            }else{
                                $smsg = "Connected to Orgname: $($OrgName)" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            } ; 
                        }else{
                            $smsg = "Missing 'tmp_*' module with 'get-OrganizationConfig'!" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ; 
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = $ErrTrapd ;
                        $smsg += "`n";
                        $smsg += $ErrTrapd.Exception.Message ;
                        if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        CONTINUE ;
                    } ;
                } ; 
                if($useForestWideDYN){
                    #region  ; #*------v USEFORESTWIDEDYN OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT v------
                    $smsg = "(`$useForestWideDYN:$($useForestWideDYN)):Enabling EXoP Forestwide)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Set-AdServerSettings -ViewEntireForest $True ;
                    #endregion  ; #*------^ END USEFORESTWIDEDYN OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT ^------
                } ;
            } else {
                $smsg = "(`$UseOPDYN:$($UseOPDYN))" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            }  ;  # if-E $UseOPDYN
            #endregion GENERIC_EXOP_SRVR_CONN #*------^ GENERIC_EXOP_SRVR_CONN BP ^------
            #region UseOPDYN #*------v UseOPDYN v------
            if($UseOPDYN -OR $UseOPADDYN){
                #region GENERIC_ADMS_CONN_&_XO #*------v GENERIC ADMS CONN & XO  v------
                if($UseOPADDYN){
                    'load-ADMs' | foreach-object{
                        if(get-command $_ -ErrorAction STOP){} else { 
                            $smsg = "MISSING DEP FUNC:$($_)! must be either local include, or pre-loaded for EXOP connectivity to work!" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            throw $smsg ; 
                            break ; 
                        } ; 
                    } ; 
                } ; 
                $smsg = "(loading ADMS...)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                # always capture load-adms return, it outputs a $true to pipeline on success
                $ADMTLoaded = load-ADMS -Verbose:$FALSE ;
                # 9:32 AM 4/20/2023 trimmed disabled/fw-borked cross-org code
                TRY {
                    if(-not(Get-ADDomain  -ea STOP).DNSRoot){
                        $smsg = "Missing AD Connection! (no (Get-ADDomain).DNSRoot returned)" ; 
                        throw $smsg ; 
                        $smsg | write-warning  ; 
                    } ; 
                    $objforest = get-adforest -ea STOP ; 
                    # Default new UPNSuffix to the UPNSuffix that matches last 2 elements of the forestname.
                    if($forestdom = $UPNSuffixDefault = $objforest.UPNSuffixes | ?{$_ -eq (($objforest.name.split('.'))[-2..-1] -join '.')}){

                    } elseif( $objforest.RootDomain -eq 'cmw.internal'){
                        # cmw doesn't use cmw.internal (forestname), as a UPN suffix, to try to match
                        # they have no default, tho' if we build up in the route, they'd have charlesmachineworks.com
                        $forestdom = $objforest.RootDomain; 
                        $UPNSuffixDefault = 'charlesmachine.works'
                    } else {
                         $smsg = "Unsupported `$objforest.RootDomain ($objforest.RootDomain), aborting!" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        throw $smsg ; 
                        break ; 
                    }
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = $ErrTrapd ;
                    $smsg += "`n";
                    $smsg += $ErrTrapd.Exception.Message ;
                    if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    CONTINUE ;
                } ;        
                #endregion GENERIC_ADMS_CONN_&_XO #*------^ END GENERIC ADMS CONN & XO ^------
            } else {
                $smsg = "(`$UseOPDYN:$($UseOPDYN))" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            }  ;
            #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
            #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller = get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
            # use new get-GCFastXO cross-org dc finde
            # default to Op_ExADRoot forest from $TenOrg Meta
            #if($UseOPDYN -AND -not $domaincontroller){
            if($UseOPDYN -AND -not (get-variable domaincontroller -ea 0)){
                #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((get-variable -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};
                # need to debug the above, credential issue?
                # just get it done
                #$domaincontroller = get-GCFast
                # AD – Return one GC in local site: (uses ADMS) Completely dynamic, no installed verb-xxx dependancies, 
                $domaincontroller = (Get-ADDomainController -Filter  {isGlobalCatalog -eq $true -AND Site -eq "$((get-adreplicationsite).name)"}).name| Get-Random ; 
            }  else { 
                # have to defer to get-azuread, or use EXO's native cmds to poll grp members
                # TODO 1/15/2021
                $useEXOforGroups = $true ; 
                $smsg = "$($TenOrg):HAS NO ON-PREM ACTIVEDIRECTORY, DEFERRING ALL GROUP ACCESS & MGMT TO NATIVE EXO CMDS!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
            if($useForestWideDYN -AND -not $GcFwide){
                #region GCFWIDE ; #*------v GCFWIDE OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT: v------
                $smsg = "`$GcFwide = Get-ADDomainController -Discover -Service GlobalCatalog" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;
                $smsg = "Discovered `$GcFwide:$($GcFwide)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #endregion GCFWIDE ; #*------^ END GCFWIDE OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT ^------
            } ;
            #endregion UseOPDYN #*------^ END UseOPDYN ^------
            #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
            #endregion SERVICE_CONNECTIONS_DEPENDANCYLESS #*======^ END SERVICE_CONNECTIONS_DEPENDANCYLESS ^======

        } ;  # if-E $isBased
    } ;  # PROC-E
    END {
        <# borked by psreadline v1/v2 breaking changes
        if(($PSFgColor = (Get-Variable  -name "$($TenOrg)Meta").value.PSFgColor) -AND ($PSBgColor = (Get-Variable  -name "$($TenOrg)Meta").value.PSBgColor)){
            write-verbose "(setting console colors:$($TenOrg)Meta.PSFgColor:$($PSFgColor),PSBgColor:$($PSBgColor))" ;
            $Host.UI.RawUI.BackgroundColor = $PSBgColor
            $Host.UI.RawUI.ForegroundColor = $PSFgColor ;
        } ;
        #>
    }
}

#*------^ Connect-Ex2010.ps1 ^------