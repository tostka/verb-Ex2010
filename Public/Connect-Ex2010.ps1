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
        [Parameter(Position = 0, HelpMessage = "Exch server to Remote to")][string]$ExchangeServer,
        [Parameter(HelpMessage = 'Use exadmin IIS WebPool for remote EMS[-ExAdmin]')]$ExAdmin,
        [Parameter(HelpMessage = 'Credential object')][System.Management.Automation.PSCredential]$Credential = $credOpTORSID
    )  ;
    BEGIN{
        #$verbose = ($VerbosePreference -eq "Continue") ;
        $sWebPoolVariant = "exadmin" ;
        $CommandPrefix = $null ;
        # use credential domain to determine target org
        $rgxLegacyLogon = '\w*\\\w*' ;
        $TenOrg = get-TenantTag -Credential $Credential ;
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
            $Global:E10Sess = New-PSSession @pltNSess -ea STOP  ;
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
                $Global:E10Sess = New-PSSession @pltNSess -ea STOP  ;
            } else {
                BREAK ;
            } ;
        } ;

        write-verbose "$((get-date).ToString('HH:mm:ss')):Importing Exchange 2010 Module" ;
        #$pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ; verbose=$false} ; # force verbose off, suppress spam in console
        # tear verbose out
        $pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ;} ;
        #$pltISess = [ordered]@{Session = $Global:E10Sess ; DisableNameChecking = $true  ; AllowClobber = $true ; ErrorAction = 'Stop' ; Verbose = $false ;} ;
        $pltISess = [ordered]@{Session = $Global:E10Sess ; DisableNameChecking = $true  ; AllowClobber = $true ; ErrorAction = 'Stop' ; } ;
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
        Add-PSTitleBar 'EMS' ;
        # tag E10IsDehydrated
        $Global:E10IsDehydrated = $true ;
        write-host -foregroundcolor darkgray "`n$(($Global:E10Sess | select ComputerName,Availability,State,ConfigurationName | format-table -auto |out-string).trim())" ;
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