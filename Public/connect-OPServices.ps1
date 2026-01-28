# connect-OPServices.ps1

#region CONNECT_OPSERVICES ; #*======v connect-OPServices v======
#if(-not (get-childitem function:connect-OPServices -ea 0)){
    function connect-OPServices {
        <#
        .SYNOPSIS
        connect-OPServices - logic wrapper for my histortical scriptblock that resolves creds, svc avail and relevent status, to connect to range of Services (in OnPrem)
        .NOTES
        Version     : 0.0.
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2024-06-07
        FileName    : connect-OPServices
        License     : MIT License
        Copyright   : (c) 2024 Todd Kadrie
        Github      : https://github.com/tostka/verb-AAD
        Tags        : Powershell,AzureAD,Authentication,Test
        AddedCredit :
        AddedWebsite:
        AddedTwitter:
        REVISIONS
        * 1:02 PM 1/27/2026 latest dbg'd version
        * 1:07 PM 1/20/2026 pulled the defer, should *never* defer the function source copy
        * 9:00 AM 6/3/2025 revised CBH demo, properly handle cross-org conn attempts, incl forestwide spec recovery
        * 4:36 PM 6/2/2025 updated CBH demo to cover cross org fails, wo breaking cloud run (against MGDomain updates .ps1s)
        * 2:56 PM 5/19/2025 updated cross-org access fail, to rnot say missing creds ; rem'd $prefVaris dump (blank values, throws errors)
        3:35 PM 5/16/2025 spliced over local dep internal_funcs (out of the main paramt block) ;  dbgd, few minor fixes; but substantially working
        * 8:16 AM 5/15/2025 init
        .DESCRIPTION
        connect-OPServices - logic wrapper for my histortical scriptblock that resolves creds, svc avail and relevent status, to connect to range of Services (in OnPrem)
        .PARAMETER EnvSummary
        Pre-resolved local environrment summary (product of output of verb-io\resolve-EnvironmentTDO())[-EnvSummary `$rvEnv]
        .PARAMETER NetSummary
        Pre-resolved local network summary (product of output of verb-network\resolve-NetworkLocalTDO())[-NetSummary `$netsettings]
        .PARAMETER XoPSummary
        Pre-resolved local ExchangeServer summary (product of output of verb-ex2010\test-LocalExchangeInfoTDOO())[-XoPSummary `$lclExOP]
        .PARAMETER UseExOP
        Connect to OnPrem ExchangeManagementShell(Remote (Local,Edge))[-UseExOP]
        .PARAMETER useExopNoDep
        Connect to OnPrem ExchangeManagementShell using No Dependancy options)[-useEXO]
        .PARAMETER ExopVers
        Connect to OnPrem ExchangeServer version (Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000). An array represents a min/max range of all between; null indicates all versions returned by verb-Ex2010\get-ADExchangeServerTDO())[-useEXO]
        XOP Switch to set ForestWide Exchange EMS scope(e.g. Set-AdServerSettings -ViewEntireForest `$True)[-useForestWide]
        .PARAMETER UseOPAD
        Connect to OnPrem ActiveDirectory powershell module)[-UseOPAD]
        .PARAMETER useExOPVers
        String array to indicate target OnPrem Exchange Server version for connections. If an array is specified, will be assumed to reflect a span of versions to include, connections will aways be to a random server of the latest version specified (Ex2000|Ex2003|Ex2007|Ex2010|Ex2000|Ex2003|Ex2007|Ex2010|Ex2016|Ex2019), used with verb-Ex2010\get-ADExchangeServerTDO() dyn location via ActiveDirectory.[-useExOPVers @('Ex2010','Ex2016')]
        .PARAMETER TenOrg
        Tenant Tag (3-letter abbrebiation)[-TenOrg 'XYZ']
        .PARAMETER Credential
        Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]
        .PARAMETER UserRole
        Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]
        .PARAMETER useExOPVers
        String array to indicate target OnPrem Exchange Server version to target with connections, if an array, will be assumed to reflect a span of versions to include, connections will aways be to a random server of the latest version specified (Ex2000|Ex2003|Ex2007|Ex2010|Ex2000|Ex2003|Ex2007|Ex2010|Ex2016|Ex2019), used with verb-Ex2010\get-ADExchangeServerTDO() dyn location via ActiveDirectory.[-useExOPVers @('Ex2010','Ex2016')]
        .PARAMETER Silent
        Silent output (suppress status echos)[-silent]
        .INPUTS
        Does not accept piped input
        .OUTPUTS
        None (records transcript file)
        .EXAMPLE
        PS> $PermsRqd = connect-OPServices -path D:\scripts\new-MGDomainRegTDO.ps1 ;
        Typical pass script pass, using the -path param
        .EXAMPLE
        PS> $PermsRqd = connect-OPServices -scriptblock (gcm -name connect-OPServices).definition ;
        Typical function pass, using get-command to return the definition/scriptblock for the subject function.
        .EXAMPLE
        PS> #region CALL_CONNECT_OPSERVICES ; #*======v CALL_CONNECT_OPSERVICES v======
        PS> #$useOP = $false ; 
        PS> if($useOP){
        PS>     $pltCcOPSvcs=[ordered]@{
        PS>         # environment parameters:
        PS>         EnvSummary = $rvEnv ;
        PS>         NetSummary = $netsettings ;
        PS>         XoPSummary = $lclExOP ;
        PS>         # service choices
        PS>         UseExOP = $true ;
        PS>         useForestWide = $true ;
        PS>         useExopNoDep = $false ;
        PS>         ExopVers = 'Ex2010' ;
        PS>         UseOPAD = $true ;
        PS>         useExOPVers = $useExOPVers; # 'Ex2010' ;
        PS>         # Service Connection parameters
        PS>         TenOrg = $TenOrg ; # $global:o365_TenOrgDefault ;
        PS>         Credential = $Credential ;
        PS>         #[ValidateSet("SID","ESVC","LSVC")]
        PS>         UserRole = $UserRole ; # @('SID','ESVC') ;
        PS>         # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
        PS>         silent = $silent ;
        PS>     } ;
        PS>     
        PS>     write-verbose "(Purge no value keys from splat)" ;
        PS>     $mts = $pltCcOPSvcs.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltCcOPSvcs.remove($_.Name)} ; rv mts -ea 0 ;
        PS>     if((get-command connect-OPServices -EA STOP).parameters.ContainsKey('whatif')){
        PS>         $pltCcOPSvcsnDSR.add('whatif',$($whatif))
        PS>     } ;
        PS>     $smsg = "connect-OPServices w`n$(($pltCcOPSvcs|out-string).trim())" ;
        PS>     if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        PS>     $ret_CcOPSvcs = connect-OPServices @pltCcOPSvcs ; 
        PS>     
        PS>     # #region CONFIRM_CCOPRETURN ; #*------v CONFIRM_CCOPRETURN v------
        PS>     # matches each: $plt.useXXX:$true to matching returned $ret.hasXXX:$true
        PS>     $vplt = $pltCcOPSvcs ; $vret = 'ret_CcOPSvcs' ;  ; $ACtionCommand = 'connect-OPServices' ; 
        PS>     $vplt.GetEnumerator() |?{$_.key -match '^use' -ANd $_.value -match $true} | foreach-object{
        PS>         $pltkey = $_ ;
        PS>         $smsg = "$(($pltkey | ft -HideTableHeaders name,value|out-string).trim())" ; 
        PS>         if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        PS>         else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        PS>         $vtests = @() ;  $vFailMsgs = @()  ; 
        PS>         $tprop = $pltkey.name -replace '^use','has';
        PS>         if($rProp = (gv $vret).Value.psobject.properties | ?{$_.name -match $tprop}){
        PS>             $smsg = "$(($rprop | ft -HideTableHeaders name,value |out-string).trim())" ; 
        PS>             if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        PS>             else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        PS>             if($rprop.Value -eq $pltkey.value){
        PS>                 $vtests += $true ; 
        PS>                 $smsg = "Validated: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
        PS>                 if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
        PS>                 else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        PS>                 #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
        PS>             } else {
        PS>                 $smsg = "NOT VALIDATED: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
        PS>                 $vtests += $false ; 
        PS>                 $vFailMsgs += "`n$($smsg)" ; 
        PS>                 if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        PS>                 else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        PS>             };
        PS>         } else{
        PS>             $smsg = "Unable to locate: $($pltKey.name):$($pltKey.value) to any matching $($rprop.name)!)" ;
        PS>             $smsg = "" ; 
        PS>             if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        PS>             else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        PS>         } ; 
        PS>     } ; 
        PS>     if($useOP -AND $vtests -notcontains $false){
        PS>         $smsg = "==> $($ACtionCommand): confirmed specified connections *all* successful " ; 
        PS>         if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
        PS>         else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        PS>         #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
        PS>     }elseif($vtests -contains $false -AND (get-variable ret_CcOPSvcs) -AND (gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper() -ne $env:userdomain){
        PS>         $smsg = "==> $($ACtionCommand): FAILED SOME SPECIFIED CONNECTIONS" ; 
        PS>         $smsg += "`nCROSS-ORG ONPREM CONNECTION: ATTEMPTING TO CONNECT TO ONPREM '$((gv -name "$($tenorg)meta").value.o365_Prefix)' $((gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper()) domain, FROM $($env:userdomain)!" ;
        PS>         $smsg += "`nEXPECTED ERROR, SKIPPING ONPREM ACCESS STEPS (force `$useOP:$false)" ;
        PS>         if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        PS>         else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        PS>         $useOP = $false ; 
        PS>     }elseif(-not $useOP -AND -not (get-variable ret_CcOPSvcs)){
        PS>         $smsg = "-useOP: $($useOP), skipped connect-OPServices" ; 
        PS>         if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        PS>         else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        PS>     } else {
        PS>         $smsg = "==> $($ACtionCommand): FAILED SOME SPECIFIED CONNECTIONS" ; 
        PS>         $smsg += "`n`$ret_CcOPSvcs:`n$(($ret_CcOPSvcs|out-string).trim())" ; 
        PS>         if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        PS>         else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        PS>         $sdEmail.SMTPSubj = "FAIL Rpt:$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"
        PS>         $sdEmail.SmtpBody = "`n===Processing Summary:" ;
        PS>         if($vFailMsgs){
        PS>             $sdEmail.SmtpBody += "`n$(($vFailMsgs|out-string).trim())" ; 
        PS>         } ; 
        PS>         $sdEmail.SmtpBody += "`n" ;
        PS>         if($SmtpAttachment){
        PS>             $sdEmail.SmtpAttachment = $SmtpAttachment
        PS>             $sdEmail.smtpBody +="`n(Logs Attached)" ;
        PS>         };
        PS>         $sdEmail.SmtpBody += "Pass Completed $([System.DateTime]::Now)" ;
        PS>         $smsg = "Send-EmailNotif w`n$(($sdEmail|out-string).trim())" ;
        PS>         if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        PS>         else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        PS>         Send-EmailNotif @sdEmail ;
        PS>         throw $smsg ; 
        PS>         BREAK ; 
        PS>     } ; 
        PS>     #endregion CONFIRM_CCOPRETURN ; #*------^ END CONFIRM_CCOPRETURN ^------
        PS>     #region CONFIRM_OPFORESTWIDE ; #*------v CONFIRM_OPFORESTWIDE v------    
        PS>     if($useOP -AND $pltCcOPSvcs.useForestWide -AND $ret_CcOPSvcs.hasForestWide -AND $ret_CcOPSvcs.AdGcFwide){
        PS>         $smsg = "==> $($ACtionCommand): confirmed has BOTH .hasForestWide & .AdGcFwide ($($ret_CcOPSvcs.AdGcFwide))" ; 
        PS>         if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
        PS>         else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        PS>         #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success        
        PS>     }elseif($pltCcOPSvcs.useForestWide -AND (get-variable ret_CcOPSvcs) -AND (gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper() -ne $env:userdomain){
        PS>         $smsg = "`nCROSS-ORG ONPREM CONNECTION: ATTEMPTING TO CONNECT TO ONPREM '$((gv -name "$($tenorg)meta").value.o365_Prefix)' $((gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper()) domain, FROM $($env:userdomain)!" ;
        PS>         $smsg += "`nEXPECTED ERROR, SKIPPING ONPREM FORESTWIDE SPEC" ;
        PS>         if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        PS>         else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        PS>         $useOP = $false ; 
        PS>     }elseif($useOP -AND $pltCcOPSvcs.useForestWide -AND -NOT $ret_CcOPSvcs.hasForestWide){
        PS>         $smsg = "==> $($ACtionCommand): MISSING CRITICAL FORESTWIDE SUPPORT COMPONENT:" ; 
        PS>         if(-not $ret_CcOPSvcs.hasForestWide){
        PS>             $smsg += "`n----->$($ACtionCommand): MISSING .hasForestWide (Set-AdServerSettings -ViewEntireForest `$True) " ; 
        PS>         } ; 
        PS>         if(-not $ret_CcOPSvcs.AdGcFwide){
        PS>             $smsg += "`n----->$($ACtionCommand): MISSING .AdGcFwide GC!:`n((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):326) " ; 
        PS>         } ; 
        PS>         if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        PS>         else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        PS>         $smsg = "MISSING SOME KEY CONNECTIONS. DO YOU WANT TO IGNORE, AND CONTINUE WITH CONNECTED SERVICES?" ;
        PS>         if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        PS>         else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        PS>         $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ;
        PS>         if ($bRet.ToUpper() -eq "YYY") {
        PS>             $smsg = "(Moving on), WITH THE FOLLOW PARTIAL CONNECTION STATUS" ;
        PS>             $smsg += "`n`n$(($ret_CcOPSvcs|out-string).trim())" ; 
        PS>             write-host -foregroundcolor green $smsg  ;
        PS>         } else {
        PS>             throw $smsg ; 
        PS>             break ; #exit 1
        PS>         } ;         
        PS>     }; 
        PS>     #endregion CONFIRM_OPFORESTWIDE ; #*------^ END CONFIRM_OPFORESTWIDE ^------
        PS> } ; 
        PS> #endregion CALL_CONNECT_OPSERVICES ; #*======^ END CALL_CONNECT_OPSERVICES ^======             
        Demo leveraging resolve-environmentTDO outputs
        .LINK
        https://bitbucket.org/tostka/verb-dev/
        #>
        ##Requires -Modules AzureAD, verb-AAD
        [CmdletBinding()]
        ## PSV3+ whatif support:[CmdletBinding(SupportsShouldProcess)]
        ###[Alias('Alias','Alias2')]
        PARAM(
            # environment parameters:
            [Parameter(Mandatory=$true,HelpMessage="Pre-resolved local environrment summary (product of output of verb-io\resolve-EnvironmentTDO())[-EnvSummary `$rvEnv]")]
                $EnvSummary, # $rvEnv
            [Parameter(Mandatory=$true,HelpMessage="Pre-resolved local network summary (product of output of verb-network\resolve-NetworkLocalTDO())[-NetSummary `$netsettings]")]
                $NetSummary, # $netsettings
            [Parameter(Mandatory=$true,HelpMessage="Pre-resolved local ExchangeServer summary (product of output of verb-ex2010\test-LocalExchangeInfoTDOO())[-XoPSummary `$lclExOP]")]
                $XoPSummary, # $lclExOP = test-LocalExchangeInfoTDO ;
            # service choices
            # OP switches
            #[Parameter(HelpMessage="Connect to OnPrem ExchangeManagementShell(Remote (Local,Edge))[-UseOP]")]
            #    [switch]$UseOP, # interpolate from below
            [Parameter(HelpMessage="Connect to OnPrem ExchangeManagementShell(Remote (Local,Edge))[-UseExOP]")]
                [switch]$UseExOP,
            [Parameter(HelpMessage="Connect to OnPrem ExchangeManagementShell using No Dependancy options)[-useEXO]")]
                [switch]$useExopNoDep,
            [Parameter(HelpMessage="Connect to OnPrem ExchangeServer version (Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000). An array represents a min/max range of all between; null indicates all versions returned by verb-Ex2010\get-ADExchangeServerTDO())[-useEXO]")]
                [AllowNull()]
                [ValidateSet('Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000')]
                [string[]]$ExopVers, # = 'Ex2010' # 'Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000', Null for All versions
                #if($Version){
                #   $ExopVers = $Version ; #defer to local script $version if set
                #} ;
            [Parameter(HelpMessage="XOP Switch to set ForestWide Exchange EMS scope(e.g. Set-AdServerSettings -ViewEntireForest `$True)[-useForestWide]")]
                [switch]$useForestWide,
            [Parameter(HelpMessage="Connect to OnPrem ActiveDirectory powershell module)[-UseOPAD]")]
                [switch]$UseOPAD,
            #
            # Service Connection parameters
            [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
                [ValidateNotNullOrEmpty()]
                #[ValidatePattern("^\w{3}$")]
                [string]$TenOrg = $global:o365_TenOrgDefault,
            [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]")]
                [System.Management.Automation.PSCredential]$Credential,
            [Parameter(Mandatory = $false, HelpMessage = "Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]")]
                # sourced from get-admincred():#182: $targetRoles = 'SID', 'CSID', 'ESVC','CSVC','UID','ESvcCBA','CSvcCBA','SIDCBA' ;
                #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
                # pulling the pattern from global vari w friendly err
                [ValidateScript({
                    if(-not $rgxPermittedUserRoles){$rgxPermittedUserRoles = '(SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)'} ;
                    if(-not ($_ -match $rgxPermittedUserRoles)){throw "'$($_)' doesn't match `$rgxPermittedUserRoles:`n$($rgxPermittedUserRoles.tostring())" ; } ;
                    return $true ;
                })]
                [string[]]$UserRole = @('SID','ESVC'),
                # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
            [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
                [switch] $silent,
            [Parameter(Mandatory=$FALSE,HelpMessage="String array to indicate target OnPrem Exchange Server version for connections. If an array is specified, will be assumed to reflect a span of versions to include, connections will aways be to a random server of the latest version specified (Ex2000|Ex2003|Ex2007|Ex2010|Ex2000|Ex2003|Ex2007|Ex2010|Ex2016|Ex2019), used with verb-Ex2010\get-ADExchangeServerTDO() dyn location via ActiveDirectory.[-useExOPVers @('Ex2010','Ex2016')]")]
                [AllowNull()]
                [ValidateSet('Ex2000','Ex2003','Ex2007','Ex2010','Ex2000','Ex2003','Ex2007','Ex2010','Ex2016','Ex2019')]
                [string[]]$useExOPVers = 'Ex2010'
        );
        BEGIN {
            # for scripts wo support, can use regions to fake BEGIN;PROCESS;END: (tho' can use the real deal in scripts as well as adv funcs, as long as all code is inside the blocks)
            # ps1 faked:#region BEGIN ; #*------v BEGIN v------
            # 8:59 PM 4/23/2025 with issues in CMW - funcs unrecog'd unless loaded before any code use - had to move the entire FUNCTIONS block to the top of BEGIN{}

            #region FUNCTIONS_INTERNAL ; #*======v FUNCTIONS_INTERNAL v======
            # Pull the CUser mod dir out of psmodpaths:
            #$CUModPath = $env:psmodulepath.split(';')|?{$_ -like '*\Users\*'} ;

            #region CONNECT_EXCHANGESERVERTDO ; #*------v Connect-ExchangeServerTDO v------
            if(-not(gi function:Connect-ExchangeServerTDO -ea 0)){
                Function Connect-ExchangeServerTDO {
                    <#
                    .SYNOPSIS
                    Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
                    will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellREmote (REMS) connect to each server, 
                    stopping at the first successful connection.
                    .NOTES
                    REVISIONS
                    * 3:58 PM 5/14/2025 restored prior dropped earlier rev history (routinely trim for psparamt inclu)
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
                    .PARAMETER Version
                    Specific Exchange Server Version to connect to('Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000')[-Version 'Ex2016']
                    .PARAMETER TenOrg
                    Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
                    .INPUTS
                    None. Does not accepted piped input.(.NET types, can add description)
                    .OUTPUTS
                    [system.object] Returns a system object containing a successful PSSession
                    System.Boolean
                    .EXAMPLE
                    PS> $PSSession = Connect-ExchangeServerTDO -siteName SITENAME -RoleNames @('HUB','CAS') -verbose 
                    Demo's connecting to a functional Hub or CAS server in the SITENAME site with verbose outputs, the `PSSession variable will contain information about the successful connection. Makes automatic Exchangeserver discovery calls into AD (using ADSI) leveraging the separate get-ADExchangeServerTDO()
                    .EXAMPLE
                    PS> TRY{$Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name}CATCH{$Site=$env:COMPUTERNAME} ;
                    PS> $PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
                    Demo including support for EdgeRole, which is detected on it's lack of AD Site specification (which gets fed through to call, by setting the Site to the machine itself).
                    .EXAMPLE
                    PS> $PSSession = Connect-ExchangeServerTDO -siteName SITENAME -RoleNames @('HUB','CAS') -Version Ex2016 -verbose 
                    Demo's connecting to a functional Hub or CAS server Version Ex2016 in the SITENAME site with verbose outputs, the `PSSession variable will contain information about the successful connection. Makes automatic Exchangeserver discovery calls into AD (using ADSI) leveraging the separate get-ADExchangeServerTDO()
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
                        [Parameter(Position=2,HelpMessage="Specific Exchange Server Version to connect to('Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000')[-Version 'Ex2016']")]
                            [ValidateSet('Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000')]
                            [string[]]$Version = 'Ex2010',
                        [Parameter(Mandatory=$FALSE,HelpMessage="Tenant Tag (3-letter abbrebiation - defaults to global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']")]
                            [ValidateNotNullOrEmpty()]
                            [string]$TenOrg = $global:o365_TenOrgDefault
                    ) ;
                    BEGIN{
                        $Verbose = ($VerbosePreference -eq 'Continue') ;
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
                
                        # 5:15 PM 4/22/2025 on CMW, have to patch version to Ex2016

                        #*------v Function _connect-ExOP v------
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
                                            $smsg = "We are on Exchange Edge Transport Server"
                                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                            $IsEdgeTransport = $true
                                        }
                                        TRY {
                                            Get-ExchangeServer -ErrorAction Stop | Out-Null
                                            $smsg = "Exchange PowerShell Module already loaded."
                                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                            $passed = $true 
                                        }CATCH {
                                            $smsg = "Failed to run Get-ExchangeServer"
                                            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
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
                                                    $smsg = "Failed to Load Exchange PowerShell Module..." ; 
                                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
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
    
                                                    $smsg = ("Set ExInstall: {0}" -f $Global:ExInstall)
                                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                                    $smsg = ("Set ExBin: {0}" -f $Global:ExBin)
                                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                                } ; 
                                            } ; 
                                        } ; 
                                    } else  {
                                        $smsg = "Does not appear to be an Exchange 2010 or newer server." ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                                    }
                                    if(get-command -Name Get-OrganizationConfig -ea 0){
                                        $smsg = "Running in connected/Native EMS" ; 
                                        if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
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
                                    $pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ;} ;
                                    # use ExVersUnm dd instead of hardcoded (Exchange2010)
                                    if($ExVersNum -ge 15){
                                        $smsg = "EXOP.15+:Adding -Authentication Kerberos" ;
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
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
                                    # 3:59 PM 1/9/2025 appears credprompting is due to it's missing the import-module $ExIPSS ! 
                                    $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                    $Global:E10Mod = Import-Module $ExIPSS @pltIMod ;
                                    $ExPSS | write-output ;
                                    $ExPSS= $ExIPSS = $null ;
                                } ; 
                            } ;
                        #*------^ END Function _connect-ExOP ^------
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
                                $pltCXOP.Add('Credential',$pltGADX.credential) ;
                            } ;
                            if($Version){
                                switch ($Version){
                                  'Ex2000'{$rgxExVersNum = '6' } 
                                  'Ex2003'{$rgxExVersNum = '6.5' } 
                                  'Ex2007'{$rgxExVersNum = '8.*' } 
                                  'Ex2010'{$rgxExVersNum = '14.*'} 
                                  'Ex2013'{$rgxExVersNum = '15.0' } 
                                  'Ex2016'{$rgxExVersNum = '15.1'} 
                                  'Ex2019'{$rgxExVersNum = '15.2' } 
                                } ; 
                                $exchServers  = $exchServers | ?{ [double]([regex]::match( $_.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value) -match $rgxExVersNum } ; 

                            } else {
                                write-verbose "no -Version: Sorting Newest first, then names, descending" ; 
                                $exchServers  = $exchServers | sort version,name -desc
                            } ; 
                            $prpPSS = 'Id','Name','ComputerName','ComputerType','State','ConfigurationName','Availability' ; 
                            foreach($exServer in $exchServers){
                                $smsg = "testing conn to:$($exServer.name.tostring())..." ; 
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                #if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig' -ea SilentlyContinue){
                                if($tmod = (get-module |?{$_.name -like 'tmp_*'}).name){
                                    if(get-command -module $tmod.name -name 'get-OrganizationConfig' -ea SilentlyContinue){
                                    # above throws an error: get-command : The term 'get-OrganizationConfig' is not recognized as the name of a cmdlet, function, script file, or operable program. Check the spelling of the name, or if a path was included, verify that the path is correct and try again.
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
                                }else{
                                    $smsg = "UNABLE TO:`$tmod = (get-module |?{$_.name -like 'tmp_*'}).name ~" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                    THROW $SMSG ; 
                                    BREAK ; 
                                } 
                                if(-not $pssEXOP){
                                    $smsg = "Connecting to: $($exServer.FQDN)" ;
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                    $smsg = "_connect-ExOP w`n$(($pltCXOP|out-string).trim())" ;
                                    $smsg += "`nServer $($exServer.FQDN)" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
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
            } ; 
            #endregion CONNECT_EXCHANGESERVERTDO ; #*------^ END Connect-ExchangeServerTDO ^------

            #region GET_ADEXCHANGESERVERTDO ; #*------v get-ADExchangeServerTDO v------
            if(-not(gi function:get-ADExchangeServerTDO -ea 0)){
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
            }
            #endregion GET_ADEXCHANGESERVERTDO ;#*------^ END Function get-ADExchangeServerTDO ^------ ;

            #region load_ADMS  ; #*------v load-ADMS v------
            if(-not(gi function:load-ADMS -ea 0)){
                function load-ADMS {
                    <#
                    .NOTES
                    REVISIONS   :
                    * 4:08 PM 5/14/2025 added alias import-adms
                    .INPUTS
                    None.
                    .OUTPUTS
                    Outputs $True/False load-status
                    .EXAMPLE
                    PS> $ADMTLoaded = load-ADMS ; Write-Debug "`$ADMTLoaded: $ADMTLoaded" ;
                    .EXAMPLE
                    PS> $ADMTLoaded = load-ADMS -Cmdlet get-aduser,get-adcomputer ; Write-Debug "`$ADMTLoaded: $ADMTLoaded" ;
                    Load solely the specified cmdlets from ADMS
                    .EXAMPLE
                    # load ADMS
                    PS> $reqMods+="load-ADMS".split(";") ;
                    PS> if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
                    PS> write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):(loading ADMS...)" ;
                    PS> load-ADMS | out-null ;
                    #load-ADMS -cmdlet get-aduser,Set-ADUser,Get-ADGroupMember,Get-ADDomainController,Get-ADObject,get-adforest | out-null ;
                    Demo a load from the verb-ADMS.ps1 module, with opt specific -Cmdlet set
                    .EXAMPLE
                    PS> if(connect-ad){write-host 'connected'}else {write-warning 'unable to connect'}  ;
                    Variant capturing & testing returned (returns true|false), using the alias name (if don't cap|eat return, you'll get a 'True' in console
                    #>
                    [CmdletBinding()]
                    [Alias('connect-AD')]
                    PARAM(
                        [Parameter(HelpMessage="Specifies an array of cmdlets that this cmdlet imports from the module into the current session. Wildcard characters are permitted[-Cmdlet get-aduser]")]
                        [ValidateNotNullOrEmpty()]$Cmdlet
                    ) ;
                    $Verbose = ($VerbosePreference -eq 'Continue') ;
                    # focus specific cmdlet loads to SPEED them UP!
                    $tMod = "ActiveDirectory" ;
                    $ModsReg=Get-Module -Name $tMod -ListAvailable ;
                    $ModsLoad=Get-Module -name $tMod ;
                    $pltAD=@{Name=$tMod ; ErrorAction="Stop"; Verbose = ($VerbosePreference -eq 'Continue') } ;
                    if($Cmdlet){$pltAD.add('Cmdlet',$Cmdlet) } ;
                    if ($ModsReg) {
                        if (!($ModsLoad)) {
                            $env:ADPS_LoadDefaultDrive = 0 ;
                            import-module @pltAD;
                            if(get-command Add-PSTitleBar -ea 0){
                                Add-PSTitleBar 'ADMS' -verbose:$($VerbosePreference -eq "Continue") ;
                            } ; 
                            return $TRUE;
                        } else {
                            return $TRUE;
                        } # if-E ;
                    } else {
                        Write-Error {"$((get-date).ToString('HH:mm:ss')):($env:computername) does not have AD Mgmt Tools installed!";};
                        return $FALSE
                    } # if-E ;
                } ;
            } ; 
            #endregion load_ADMS ; #*----------^END Function load-ADMS ^---------- 

            #region GET_GCFAST ; #*------v get-GCFast v------
            if(-not(gi function:get-GCFast -ea 0)){
                function get-GCFast {
                    <#
                    .NOTES
                    REVISIONS   :
                    * 2:39 PM 1/23/2025 added -exclude (exclude array of dcs by name), -ServerPrefix (exclude on leading prefix of name) params, added expanded try/catch, swapped out w-h etc for wlt calls
                    .PARAMETER  Domain
                    Which AD Domain [Domain fqdn]
                    .PARAMETER  Site
                    DCs from which Site name (defaults to AD lookup against local computer's Site)
                    .PARAMETER Exclude
                    Array of Domain controller names in target site/domain to exclude from returns (work around temp access issues)
                    .PARAMETER ServerPrefix
                    Prefix string to filter for, in returns (e.g. 'ABC' would only return DCs with name starting 'ABC')
                    .PARAMETER SpeedThreshold
                    Threshold in ms, for AD Server response time(defaults to 100ms)
                    .INPUTS
                    None. Does not accepted piped input.
                    .OUTPUTS
                    Returns one DC object, .Name is name pointer
                    .EXAMPLE
                    PS> get-gcfast -domain dom.for.domain.com -site Site
                    Lookup a Global domain gc, with Site specified (whether in Site or not, will return remote site dc's)
                    .EXAMPLE
                    PS> get-gcfast -domain dom.for.domain.com
                    Lookup a Global domain gc, default to Site lookup from local server's perspective
                    .EXAMPLE    
                    PS> if($domaincontroller = get-gcfast -Exclude ServerBad -Verbose){
                    PS>     write-warning "Changing DomainControler: Waiting 20seconds, for RelSync..." ;
                    PS>     start-sleep -Seconds 20 ;
                    PS> } ; 
                    Demo acquireing a new DC, excluding a caught bad DC, and waiting before moving on, to permit ADRerplication from prior dc to attempt to ensure full sync of changes. 
                    PS> get-gcfast -ServerPrefix ABC -verbose
                    Demo use of -ServerPrefix to only return DCs with servernames that begin with the string 'ABC'
                    .EXAMPLE
                    PS> $adu=$null ;
                    PS> $Exit = 0 ;
                    PS> Do {
                    PS>     TRY {
                    PS>         $adu = get-aduser -id $rmbx.DistinguishedName -server $domainController -Properties $adprops -ea 0| select $adprops ;
                    PS>         $Exit = $DoRetries ;
                    PS>     }CATCH [System.Management.Automation.RuntimeException] {
                    PS>         if ($_.Exception.Message -like "*ResourceUnavailable*") {
                    PS>             $ErrorTrapped=$Error[0] ;
                    PS>             $smsg = "Failed to exec cmd because: $($ErrorTrapped.Exception.Message )" ;
                    PS>             if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    PS>             else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    PS>             # re-quire a new DC
                    PS>             $badDC = $domaincontroller ; 
                    PS>             $smsg = "PROBLEM CONTACTING $(domaincontroller)!:Resource unavailable: $($ErrorTrapped.Exception.Message)" ; 
                    PS>             $smsg += "get-GCFast() an alterate DC" ; 
                    PS>             if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    PS>             else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    PS>             if($domaincontroller = get-gcfast -Exclude $$badDC -Verbose){
                    PS>                 write-warning "Changing DomainController:($($badDC)->$($domaincontroller)):Waiting 20seconds, for ReplSync..." ;
                    PS>                 start-sleep -Seconds 20 ;
                    PS>             } ;                             
                    PS>         }else {
                    PS>             throw $Error[0] ;
                    PS>         } ; 
                    PS>     } CATCH {
                    PS>         $ErrorTrapped=$Error[0] ;
                    PS>         Start-Sleep -Seconds $RetrySleep ;
                    PS>         $Exit ++ ;
                    PS>         $smsg = "Failed to exec cmd because: $($ErrorTrapped)" ;
                    PS>         $smsg += "`nTry #: $Exit" ;
                    PS>         if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    PS>         else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    PS>         If ($Exit -eq $DoRetries) {
                    PS>             $smsg =  "Unable to exec cmd!" ;
                    PS>             if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    PS>             else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    PS>         } ;
                    PS>         Continue ;
                    PS>     }  ;
                    PS> } Until ($Exit -eq $DoRetries) ;
                    Retry demo that includes aquisition of a new DC, excluding a caught bad DC, and waiting before moving on, to permit ADRerplication from prior dc to attempt to ensure full sync of changes. 
                    #>
                    [CmdletBinding()]
                    PARAM(
                        [Parameter(Position = 0, Mandatory = $False, HelpMessage = "Optional: DCs from what Site name? (default=Discover)")]
                            [string]$Site,
                        [Parameter(HelpMessage = 'Target AD Domain')]
                            [string]$Domain,
                        [Parameter(HelpMessage = 'Array of Domain controller names in target site/domain to exclude from returns (work around temp access issues)')]
                            [string[]]$Exclude,
                        [Parameter(HelpMessage = "Prefix string to filter for, in returns (e.g. 'ABC' would only return DCs with name starting 'ABC')")]
                            [string]$ServerPrefix,
                        [Parameter(HelpMessage = 'Threshold in ms, for AD Server response time(defaults to 100ms)')]
                            $SpeedThreshold = 100
                    ) ;
                    $Verbose = $($PSBoundParameters['Verbose'] -eq $true)
                    $SpeedThreshold = 100 ;
                    $rgxSpbDCRgx = 'CN=EDCMS'
                    $ErrorActionPreference = 'SilentlyContinue' ; # Set so we don't see errors for the connectivity test
                    $env:ADPS_LoadDefaultDrive = 0 ; 
                    $sName = "ActiveDirectory"; 
                    TRY{
                        if ( -not(Get-Module | Where-Object { $_.Name -eq $sName }) ) {
                            $smsg = "Adding ActiveDirectory Module (`$script:ADPSS)" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            $script:AdPSS = Import-Module $sName -PassThru -ea Stop ;
                        } ;
                        if (-not $Domain) {
                            $Domain = (get-addomain -ea Stop).DNSRoot ; # use local domain
                            $smsg = "Defaulting domain: $Domain";
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        }
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 

                    # Get all the local domain controllers
                    if ((-not $Site)) {
                        # if no site, look the computer's Site Up in AD
                        TRY{
                            $Site = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name ;
                            $smsg = "Using local machine Site: $($Site)";
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } CATCH {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ; 
                    } ;

                    # gc filter
                    #$LocalDCs = Get-ADDomainController -filter { (isglobalcatalog -eq $true) -and (Site -eq $Site) } ;
                    # ISSUE: ==3:26 pm 3/7/2024: NO LOCAL SITE DC'S IN SPB
                    # os: LOGONSERVER=\\EDCMS8100
                    TRY{
                        $LocalDCs = Get-ADDomainController -filter { (isglobalcatalog -eq $true) -and (Site -eq $Site) -and (Domain -eq $Domain) } -ErrorAction STOP
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 
                    if( $LocalDCs){
                        $smsg = "`Discovered `$LocalDCs:`n$(($LocalDCs|out-string).trim())" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    } elseif($Site -eq 'Spellbrook'){
                        $smsg = "Get-ADDomainController -filter { (isglobalcatalog -eq `$true) -and (Site -eq $($Site)) -and (Domain -eq $($Domain)}"
                        $smsg += "`nFAILED to return DCs, and `$Site -eq Spellbrook:" 
                        $smsg += "`ndiverting to $($rgxSpbDCRgx) dcs in entire Domain:" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        TRY{
                            $LocalDCs = Get-ADDomainController -filter { (isglobalcatalog -eq $true) -and (Domain -eq $Domain) } -EA STOP | 
                                ?{$_.ComputerObjectDN -match $rgxSpbDCRgx } 
                        } CATCH {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ; 
                    } ; 

                    # any dc filter
                    #$LocalDCs = Get-ADDomainController -filter {(Site -eq $Site)} ;

                    $PotentialDCs = @() ;
                    # Check connectivity to each DC against $SpeedThreshold
                    if ($LocalDCs) {
                        foreach ($LocalDC in $LocalDCs) {
                            $TCPClient = New-Object System.Net.Sockets.TCPClient ;
                            $Connect = $TCPClient.BeginConnect($LocalDC.Name, 389, $null, $null) ;
                            $Wait = $Connect.AsyncWaitHandle.WaitOne($SpeedThreshold, $False) ;
                            if ($TCPClient.Connected) {
                                $PotentialDCs += $LocalDC.Name ;
                                $Null = $TCPClient.Close() ;
                            } # if-E
                        } ;
                        if($Exclude){
                            $smsg = "-Exclude specified:`n$((($exclude -join ',')|out-string).trim())" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            foreach($excl in $Exclude){
                                $PotentialDCs = $PotentialDCs |?{$_ -ne $excl} ; 
                            } ; 
                        } ; 
                        if($ServerPrefix){
                            $smsg = "-ServerPrefix specified: $($ServerPrefix)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            $PotentialDCs = $PotentialDCs |?{$_ -match "^$($ServerPrefix)" } ; 
        
                        }
                        write-host -foregroundcolor yellow  
                        $smsg = "`$PotentialDCs: $PotentialDCs";
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $DC = $PotentialDCs | Get-Random ;

                        $smsg = "(returning random domaincontroller from result to pipeline:$($DC)" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $DC | write-output  ;
                    } else {
                        write-host -foregroundcolor yellow  "NO DCS RETURNED BY GET-GCFAST()!";
                        write-output $false ;
                    } ;
                }  ; 
            } ; 
            #endregion GET_GCFAST ; #*------^ END get-GCFast ^------

            #endregion FUNCTIONS_INTERNAL ; #*======^ END FUNCTIONS_INTERNAL ^======

            #region CONSTANTS_AND_ENVIRO ; #*======v CONSTANTS_AND_ENVIRO v======
            #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
            <#
            $Verbose = [boolean]($VerbosePreference -eq 'Continue') ;
            $rPSCmdlet = $PSCmdlet ; # an object that represents the cmdlet or advanced function that's being run. Available on functions w CmdletBinding (& $args will not be available). (Blank on non-CmdletBinding/Non-Adv funcs).
            $rPSScriptRoot = $PSScriptRoot ; # the full path of the executing script's parent directory., PS2: valid only in script modules (.psm1). PS3+:it's valid in all scripts. (Funcs: ParentDir of the file that hosts the func)
            $rPSCommandPath = $PSCommandPath ; # the full path and filename of the script that's being run, or file hosting the funct. Valid in all scripts.
            $rMyInvocation = $MyInvocation ; # populated only for scripts, function, and script blocks.
            # - $MyInvocation.MyCommand.Name returns name of a function, to identify the current command,  name of the current script (pop'd w func name, on Advfuncs)
            # - Ps3+:$MyInvocation.PSScriptRoot : full path to the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
            # - Ps3+:$MyInvocation.PSCommandPath : full path and filename of the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
            #     ** note: above pair contain information about the _invoker or calling script_, not the current script
            $rPSBoundParameters = $PSBoundParameters ;
            #>
            #region PREF_VARI_DUMP ; #*------v PREF_VARI_DUMP v------
            <#$script:prefVaris = @{
                whatifIsPresent = $whatif.IsPresent
                whatifPSBoundParametersContains = $rPSBoundParameters.ContainsKey('WhatIf') ;
                whatifPSBoundParameters = $rPSBoundParameters['WhatIf'] ;
                WhatIfPreferenceIsPresent = $WhatIfPreference.IsPresent ; # -eq $true
                WhatIfPreferenceValue = $WhatIfPreference;
                WhatIfPreferenceParentScopeValue = (Get-Variable WhatIfPreference -Scope 1).Value ;
                ConfirmPSBoundParametersContains = $rPSBoundParameters.ContainsKey('Confirm') ;
                ConfirmPSBoundParameters = $rPSBoundParameters['Confirm'];
                ConfirmPreferenceIsPresent = $ConfirmPreference.IsPresent ; # -eq $true
                ConfirmPreferenceValue = $ConfirmPreference ;
                ConfirmPreferenceParentScopeValue = (Get-Variable ConfirmPreference -Scope 1).Value ;
                VerbosePSBoundParametersContains = $rPSBoundParameters.ContainsKey('Confirm') ;
                VerbosePSBoundParameters = $rPSBoundParameters['Verbose'] ;
                VerbosePreferenceIsPresent = $VerbosePreference.IsPresent ; # -eq $true
                VerbosePreferenceValue = $VerbosePreference ;
                VerbosePreferenceParentScopeValue = (Get-Variable VerbosePreference -Scope 1).Value;
                VerboseMyInvContains = '-Verbose' -in $rPSBoundParameters.UnboundArguments ;
                VerbosePSBoundParametersUnboundArgumentContains = '-Verbose' -in $rPSBoundParameters.UnboundArguments
            } ;
            write-verbose "`n$(($script:prefVaris.GetEnumerator() | Sort-Object Key | Format-Table Key,Value -AutoSize|out-string).trim())`n" ;
            #>
            #endregion PREF_VARI_DUMP ; #*------^ END PREF_VARI_DUMP ^------
            #region RV_ENVIRO ; #*------v RV_ENVIRO v------
            <#
            $pltRvEnv=[ordered]@{
                PSCmdletproxy = $rPSCmdlet ;
                PSScriptRootproxy = $rPSScriptRoot ;
                PSCommandPathproxy = $rPSCommandPath ;
                MyInvocationproxy = $rMyInvocation ;
                PSBoundParametersproxy = $rPSBoundParameters
                verbose = [boolean]($PSBoundParameters['Verbose'] -eq $true) ;
            } ;
            write-verbose "(Purge no value keys from splat)" ;
            $mts = $pltRVEnv.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltRVEnv.remove($_.Name)} ; rv mts -ea 0 -whatif:$false -confirm:$false;
            $smsg = "resolve-EnvironmentTDO w`n$(($pltRVEnv|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $rvEnv = resolve-EnvironmentTDO @pltRVEnv ;
            $smsg = "`$rvEnv returned:`n$(($rvEnv |out-string).trim())" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            #>
            #endregion RV_ENVIRO ; #*------^ END RV_ENVIRO ^------
            #region NETWORK_INFO ; #*======v NETWORK_INFO v======
            #$NetSummary = resolve-NetworkLocalTDO ;
            if($env:Userdomain){
                switch($env:Userdomain){
                    'CMW'{
                        #$logon_SID = $CMW_logon_SID
                    }
                    'TORO'{
                        #$o365_SIDUpn = $o365_Toroco_SIDUpn ;
                        #$logon_SID = $TOR_logon_SID ;
                    }
                    $env:COMPUTERNAME{
                        $smsg = "%USERDOMAIN% -EQ %COMPUTERNAME%: $($env:computername) => non-domain-connected, likely edge role Ex server!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        if($NetSummary.Workgroup){
                            $smsg = "WorkgroupName:$($NetSummary.Workgroup)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ;
                    } ;
                    default{
                        $smsg = "$($env:userdomain):UNRECOGIZED/UNCONFIGURED USER DOMAIN STRING!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        THROW $SMSG
                        BREAK ;
                    }
                } ;
            } ;  # $env:Userdomain-E
            #endregion NETWORK_INFO ; #*======^ END NETWORK_INFO ^======
            #region TEST_EXOPLOCAL ; #*------v TEST_EXOPLOCAL v------
            #
            #$XoPSummary = test-LocalExchangeInfoTDO ;
            write-verbose "Expand returned NoteProperty properties into matching local variables" ;
            if($host.version.major -gt 2){
                $XoPSummary.PsObject.Properties | ?{$_.membertype -eq 'NoteProperty'} | foreach-object{set-variable -name $_.name -value $_.value -verbose -whatif:$false -Confirm:$false ;} ;
            }else{
                write-verbose "Psv2 lacks the above expansion capability; just create simpler variable set" ;
                $ExVers = $XoPSummary.ExVers ; $isLocalExchangeServer = $XoPSummary.isLocalExchangeServer ; $IsEdgeTransport = $XoPSummary.IsEdgeTransport ;
            } ;
            #endregion TEST_EXOPLOCAL ; #*------^ END TEST_EXOPLOCAL ^------
            #

            #endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------
            #region TLS_LATEST_FORCE ; #*------v TLS_LATEST_FORCE v------
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
            #endregion TLS_LATEST_FORCE ; #*------^ END TLS_LATEST_FORCE ^------

            #region COMMON_CONSTANTS ; #*------v COMMON_CONSTANTS v------

            if(-not $DoRetries){$DoRetries = 4 } ;    # # times to repeat retry attempts
            if(-not $RetrySleep){$RetrySleep = 10 } ; # wait time between retries
            if(-not $RetrySleep){$DawdleWait = 30 } ; # wait time (secs) between dawdle checks
            if(-not $DirSyncInterval){$DirSyncInterval = 30 } ; # AADConnect dirsync interval
            if(-not $ThrottleMs){$ThrottleMs = 50 ;}
            if(-not $rgxDriveBanChars){$rgxDriveBanChars = '[;~/\\\.:]' ; } ; # ;~/\.:,
            if(-not $rgxCertThumbprint){$rgxCertThumbprint = '[0-9a-fA-F]{40}' } ; # if it's a 40char hex string -> cert thumbprint
            if(-not $rgxSmtpAddr){$rgxSmtpAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ; } ; # email addr/UPN
            if(-not $rgxDomainLogon){$rgxDomainLogon = '^[a-zA-Z][a-zA-Z0-9\-\.]{0,61}[a-zA-Z]\\\w[\w\.\- ]+$' } ; # DOMAIN\samaccountname
            if(-not $exoMbxGraceDays){$exoMbxGraceDays = 30} ;
            if(-not $XOConnectionUri ){$XOConnectionUri = 'https://outlook.office365.com'} ;
            if(-not $SCConnectionUri){$SCConnectionUri = 'https://ps.compliance.protection.outlook.com'} ;
            #$rgxADDistNameGAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 1 ) -join ',')"
            #$rgxADDistNameAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 2 ) -join ',')"

            write-verbose "Coerce configured but blank Resultsize to Unlimited" ;
            if(get-variable -name resultsize -ea 0){
                if( ($null -eq $ResultSize) -OR ('' -eq $ResultSize) ){$ResultSize = 'unlimited' }
                elseif($Resultsize -is [int]){} else {throw "Resultsize must be an integer or the string 'unlimited' (or blank)"} ;
            } ;
            #$ComputerName = $env:COMPUTERNAME ;
            #$NoProf = [bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
            # XXXMeta derived constants:
            # - AADU Licensing group checks
            # calc the rgxLicGrpName fr the existing $xxxmeta.rgxLicGrpDN: (get-variable tormeta).value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
            #$rgxLicGrpName = (get-variable -name "$($tenorg)meta").value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
            # use the dn vers LicGrouppDN = $null ; # | ?{$_ -match $tormeta.rgxLicGrpDN}
            #$rgxLicGrpDN = (get-variable -name "$($tenorg)meta").value.rgxLicGrpDN
            # email trigger vari, it will be semi-delimd list of mail-triggering events
            $script:PassStatus = $null ;
            # TenOrg or other looped-specific PassStatus (auto supported by 7pswlt)
            #New-Variable -Name PassStatus_$($tenorg) -scope Script -Value $null ;
            [array]$SmtpAttachment = $null ;
            #write-verbose "start-Timer:Master" ;
            $swM = [Diagnostics.Stopwatch]::StartNew() ;
            # $ByPassLocalExchangeServerTest = $true # rough in, code exists below for exempting service/regkey testing on this variable status. Not yet implemented beyond the exemption code, ported in from orig source.
            #endregion COMMON_CONSTANTS ; #*------^ END COMMON_CONSTANTS ^------

            #region LOCAL_CONSTANTS ; #*------v LOCAL_CONSTANTS v------
            # BELOW TRIGGERS/DRIVES TEST_MODS: array of: "[modname];[modDLUrl,or pscmdline install]"
            <#$tDepModules = @("Microsoft.Graph.Authentication;https://www.powershellgallery.com/packages/Microsoft.Graph/",
            "ExchangeOnlineManagement;https://www.powershellgallery.com/packages/ExchangeOnlineManagement/",
            "ActiveDirectory;get-windowscapability -name RSAT* -Online | ?{$_.name -match 'Rsat\.ActiveDirectory'} | %{Add-WindowsCapability -online -name $_.name}"
            #,"AzureAD;https://www.powershellgallery.com/packages/AzureAD"
            ) ;
            #>
            $tDepModules = @() ; 
            if($useEXO){$tDepModules += @("ExchangeOnlineManagement;https://www.powershellgallery.com/packages/ExchangeOnlineManagement/")} ; 
            if($UseMSOL){$tDepModules += @("MSOnline;https://www.powershellgallery.com/packages/MSOnline/")} ; 
            if($UseAAD){$tDepModules += @("AzureAD;https://www.powershellgallery.com/packages/AzureAD/")} ; 
            if($useEXO){$tDepModules += @("ExchangeOnlineManagement;https://www.powershellgallery.com/packages/ExchangeOnlineManagement/")} ; 
            if($UseMG){$tDepModules += @("Microsoft.Graph.Authentication;https://www.powershellgallery.com/packages/Microsoft.Graph/")} ; 
            if($UseOPAD){$tDepModules += @("ActiveDirectory;get-windowscapability -name RSAT* -Online | ?{$_.name -match 'Rsat\.ActiveDirectory'} | %{Add-WindowsCapability -online -name $_.name}")} ; 
       
            #region ENCODED_CONTANTS ; #*------v ENCODED_CONTANTS v------
            # ENCODED CONsTANTS & SUPPORT FUNCTIONS:
            #region 2B4 ; #*------v 2B4 v------
            if(-not (get-command 2b4 -ea 0)){function 2b4{[CmdletBinding()][Alias('convertTo-Base64String')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ; PROCESS{$str|%{[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_))}  };} ; } ;
            #endregion 2B4 ; #*------^ END 2B4 ^------
            #region 2B4C ; #*------v 2B4C v------
            # comma-quoted return
            if(-not (get-command 2b4c -ea 0)){function 2b4c{ [CmdletBinding()][Alias('convertto-Base64StringCommaQuoted')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ;BEGIN{$outs = @()} PROCESS{[array]$outs += $str | %{[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_))} ; } END {'"' + $(($outs) -join '","') + '"' | out-string | set-clipboard } ; } ; } ;
            #endregion 2B4C ; #*------^ END 2B4C ^------
            #region FB4 ; #*------v FB4 v------
            # DEMO: $SitesNameList = 'THluZGFsZQ==','U3BlbGxicm9vaw==','QWRlbGFpZGU=' | fb4 ;
            if(-not (get-command fb4 -ea 0)){function fb4{[CmdletBinding()][Alias('convertFrom-Base64String')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ; PROCESS{$str | %{ [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($_)) }; } ; } ; };
            #endregion FB4 ; #*------^ END FB4 ^------
            # FOLLOWING CONSTANTS ARE USED FOR DEPENDANCY-LESS CONNECTIONS
            if(-not $CMW_logon_SID){$CMW_logon_SID = 'Q01XXGQtdG9kZC5rYWRyaWU=' | fb4 } ;
            if(-not $o365_Toroco_SIDUpn){$o365_Toroco_SIDUpn = 'cy10b2RkLmthZHJpZUB0b3JvLmNvbQ==' | fb4 } ;
            if(-not $TOR_logon_SID){$TOR_logon_SID = 'VE9ST1xrYWRyaXRzcw==' | fb4 } ;

            #endregion ENCODED_CONTANTS ; #*------^ END ENCODED_CONTANTS ^------

            #endregion CONSTANTS_AND_ENVIRO ; #*======^ CONSTANTS_AND_ENVIRO ^======

            #region SUBMAIN ; #*======v SUB MAIN v======

            #region TEST_MODS ; #*------v TEST_MODS v------
            if($tDepModules){
                foreach($tmod in $tDepModules){
                    $tmodName,$tmodURL = $tmod.split(';') ;
                    if (-not(Get-Module $tmodName -ListAvailable)) {
                        $smsg = "This script requires a recent version of the $($tmodName) PowerShell module. Download it here:`n$($tmodURL )";
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        return
                    } else {
                        write-verbose "$tModName confirmed available" ;
                    } ;
                } ;
            } ;
            #endregion TEST_MODS ; #*------^ END TEST_MODS ^------

            # return status obj
            <#
            $ret_ccOPs = [ordered]@{
                CredentialOP = $null ; 
                # OP switches
                hasExOP = $false ;
                xOPPssession = $null ; 
                ExopVers = $null ;
                hasForestWide = $false ;
                AdGcFwide = $null ; 
                AdDomainController = $null ; 
                hasOPAD = $false ;
                ADForestRoot = $null ; 
                ADforestDom = $null ; 
                ADUPNSuffixDefault = $null ; 
            } ; 
            #>
            if($host.version.major -ge 3){$ret_ccOPs=[ordered]@{Dummy = $null ;} }
            else {$ret_ccOPs = $ret_ccOPs = @{Dummy = $null ;} } ;
            if($ret_ccOPs.keys -contains 'dummy'){$ret_ccOPs.remove('Dummy') };
            $fieldsBoolean = 'hasExOP','hasForestWide','hasOPAD' | select -unique | sort ; $fieldsBoolean | % { $ret_ccOPs.add($_,$false) } ;
            $fieldsnull = 'CredentialOP','ExopVers','AdGcFwide','AdDomainController','ADForestRoot','ADforestDom','ADUPNSuffixDefault' | select -unique | sort ; $fieldsnull | % { $ret_ccOPs.add($_,$null) } ;
            
            # PRETUNE STEERING separately *before* pasting in balance of region
            # THIS BLOCK DEPS ON VERB-* FANCY CRED/AUTH HANDLING MODULES THAT *MUST* BE INSTALLED LOCALLY TO FUNCTION
            # NOTE: *DOES* INCLUDE *PARTIAL* DEP-LESS $useExopNoDep=$true OPT THAT LEVERAGES Connect-ExchangeServerTDO, VS connect-ex2010 & CREDS ARE ASSUMED INHERENT TO THE ACCOUNT)
            # Connect-ExchangeServerTDO HAS SUBSTANTIAL BENEFIT, OF WORKING SEAMLESSLY ON EDGE SERVER AND RANGE OF DOMAIN-=CONNECTED EXOP ROLES
            <#
            $useO365 = $true ;
            $useEXO = $true ;
            $UseOP=$true ;
            $UseExOP=$true ;
            $useExopNoDep = $true ; # switch to use Connect-ExchangeServerTDO, vs connect-ex2010 (creds are assumed inherent to the account)
            $ExopVers = 'Ex2010' # 'Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000', Null for All versions
            if($Version){
                $ExopVers = $Version ; #defer to local script $version if set
            } ;
            $useForestWide = $true ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
            $UseOPAD = $true ;
            $UseMSOL = $false ; # should be hard disabled now in o365
            $UseAAD = $false  ;
            $UseMG = $true ;
            #>
            
            if($env:userdomain -eq $env:computername){
                $isNonDomainServer = $true ;
                $UseOPAD = $false ;
            }
            if($IsEdgeTransport){
                $UseExOP = $true ;
                if($IsEdgeTransport -AND $psise){
                    $smsg = "powershell_ISE UNDER Exchange Edge Transport role!"
                    $smsg += "`nThis script is likely to fail the get-messagetrackingLog calls with Access Denied errors"
                    $smsg += "`nif run with this combo."
                    $smsg += "`nEXIT POWERSHELL ISE, AND RUN THIS DIRECTLY UNDER EMS FOR EDGE USE";
                    $smsg += "`n(bug appears to be a conflict in Remote EMS v EMS access permissions, not resolved yet)" ;
                    write-warning $msgs ;
                } ;
            } ;
            $UseOP = [boolean]($UseOP -OR $UseExOP -OR $UseOPAD) ;
            #*------^ END STEERING VARIS ^------
            # assert Org from Credential specs (if not param'd)
            # 1:36 PM 7/7/2023 and revised again -  revised the -AND, for both, logic wasn't working
            if($TenOrg){
                $smsg = "Confirmed populated `$TenOrg" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            } elseif(-not($tenOrg) -and $Credential){
                $smsg = "(unconfigured `$TenOrg: asserting from credential)" ;
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                if((get-command get-TenantTag).Parameters.keys -contains 'silent'){
                    $TenOrg = get-TenantTag -Credential $Credential -silent ;;
                }else {
                    $TenOrg = get-TenantTag -Credential $Credential ;
                }
            } else {
                # if not using Credentials or a TargetTenants/TenOrg loop, default the $TenOrg on the $env:USERDOMAIN
                $smsg = "(unconfigured `$TenOrg & *NO* `$Credential: fallback asserting from `$env:USERDOMAIN)" ;
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                switch -regex ($env:USERDOMAIN){
                    ([regex]('(' + (( @($TORMeta.legacyDomain,$CMWMeta.legacyDomain)  |foreach-object{[regex]::escape($_)}) -join '|') + ')')).tostring() {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
                    $TOLMeta.legacyDomain {$TenOrg = 'TOL' }
                    $env:COMPUTERNAME {
                        # non-domain-joined, no domain, but the $NetSummary.fqdn has a dns suffix that can be steered.
                        if($NetSummary.fqdn){
                            switch -regex (($NetSummary.fqdn.split('.') | select -last 2 ) -join '.'){
                              'toro\.com$' {$tenorg = 'TOR' ; } ;
                              '(charlesmachineworks\.com|cmw\.internal)$' { $TenOrg = 'CMW'} ;
                              '(torolab\.com|snowthrower\.com)$'  { $TenOrg = 'TOL'} ;
                              default {throw "UNRECOGNIZED DNS SUFFIX!:$(($NetSummary.fqdn.split('.') | select -last 2 ) -join '.')" ; break ; } ;
                            } ;
                        }else{
                            throw "NIC.ip $($NetSummary.ipaddress) does not PTR resolve to a DNS A with a full fqdn!" ;
                        } ;
                    } ;
                    default {throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ; exit ; } ;
                } ;
            } ;
            
            #region GENERIC_EXOP_CREDS_N_SRVR_CONN #*------v GENERIC EXOP CREDS N SRVR CONN BP v------
            # steer all onprem code on $XXXMeta.ExOPAccessFromToro & Ex10Server values
            #$UseOP=$true ; 
            #$UseExOP=$true ;
            #$useExopNoDep = $true # switch to use Connect-ExchangeServerTDO, vs connect-ex2010 (creds are assumed inherent to the account) 
            #$useForestWide = $true ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
            <# no onprem dep
            if((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro -AND (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server){
                $UseOP = $UseExOP = $true ;
                $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`ENABLING use of OnPrem Ex system this pass." ;
                if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            } else {
                $UseOP = $UseExOP = $false ;
                $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`nDISABLING use of OnPrem Ex system this pass." ;
                if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            } ;
            #>
            if($UseOP){
                if($env:userdomain -eq (get-variable "$($Tenorg)Meta" -ea 0).value.legacydomain){
                    $smsg = "(confirmed alignment: `$env:userdomain -eq $($Tenorg)Meta.legacydomain$((get-variable "$($Tenorg)Meta" -ea 0).value.legacydomain))" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                }else{
                    $smsg = "-TenOrg:$($TenOrg) specified, but LOCAL LOGON IS IN $($env:USERDOMAIN)!" ; 
                    $smsg += "`nTHERE _CAN BE NO_ ONPREM EXCHANGE CONNECTION FROM THE LOGON DOMAIN TO THE ONPREM $($TenOrg) ONPREM DOMAIN! (fw blocked)" ;
                    $SMSG += "`nSETTING `$useOP:`$false & `$UseOPAD: `$false!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    $UseOP = $false ; $UseOPAD = $false ; 
                }
            }
            if($UseOP){
                <#if($useExopNoDep){
                    # Connect-ExchangeServerTDO use: creds are implied from the PSSession creds; assumed to have EXOP perms
                    # 3:14 PM 1/9/2025 no they aren't, it still wants explicit creds to connect - I've just been doing rx10 and pre-initiating
                } else {
                #>
                # useExopNoDep: at this point creds are *not* implied from the PS context creds. So have to explicitly pass in $creds on the new-Pssession etc, 
                # so we always need the EXOP creds block, or at worst an explicit get-credential prompt to gather when can't find in enviro or profile. 
                #*------v GENERIC EXOP CREDS N SRVR CONN BP v------
                if($TenOrg -ne 'CMW'){
                    if(get-item function:get-HybridOPCredentials -ea STOP){
                        # do the OP creds too
                        $OPCred=$null ;
                        # default to the onprem svc acct
                        # userrole='ESVC','SID'
                        #$pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC','SID'; verbose=$($verbose)} ;
                        # userrole='SID','ESVC'
                        #$pltGHOpCred=@{TenOrg=$TenOrg ;userrole='SID','ESVC'; verbose=$($verbose)} ;
                        # defer to param
                        $pltGHOpCred=@{TenOrg=$TenOrg ;userrole=$userRole ; verbose=$($verbose)} ;
                        $smsg = "get-HybridOPCredentials w`n$(($pltGHOpCred|out-string).trim())" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                            # make it script scope, so we don't have to predetect & purge before using new-variable
                            if(get-Variable -Name "cred$($tenorg)OP" -scope Script -ea 0 ){ remove-Variable -Name "cred$($tenorg)OP" -scope Script } ;
                            New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred -whatif:$false -confirm:$false; ;
                            $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } else {
                            $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                            $script:PassStatus += $statusdelta ;
                            set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                            $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                            Break ;
                        } ;
                        $smsg= "Using OnPrem/EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        <### CALLS ARE IN FORM: (cred$($tenorg))
                        $pltRX10 = @{
                            Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                            #verbose = $($verbose) ;
                            Verbose = $FALSE ; 
                        } ;
                        $1stConn = $false ; # below uses silent suppr for both x10 & xo!
                        if($1stConn){
                            $pltRX10.silent = $pltRXO.silent = $false ;
                        } else {
                            $pltRX10.silent = $pltRXO.silent =$true ;
                        } ;
                        if($pltRX10){ReConnect-Ex2010 @pltRX10 }
                        else {ReConnect-Ex2010 }
                        #$pltRx10 creds & .username can also be used for local ADMS connections
                        ###>
                        $pltRX10 = @{
                            Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                            #verbose = $($verbose) ;
                            Verbose = $FALSE ; 
                        } ;
                        if($silent -AND ((get-command Reconnect-Ex2010).Parameters.keys -contains 'silent')){
                            $pltRX10.add('Silent',[boolean]$silent) ;
                        } ;
                        # defer cx10/rx10, until just before get-recipients qry
                        # connect to ExOP X10
                    } elseif((get-variable "$($Tenorg)Meta" -ea 0).value.OP_SIDAcct){
                        $smsg = "Unable to resolve stock creds: Input suitable creds for OnPrem $($TenOrg): (defaulting to discovered OP_SIDAcct)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        $pltRX10 = @{
                            Credential = Get-Credential -Credential (get-variable "$($Tenorg)Meta" -ea 0).value.OP_SIDAcct ; ;
                            #verbose = $($verbose) ;
                            Verbose = $FALSE ; 
                        } ;
                        if($silent -AND ((get-command Reconnect-Ex2010).Parameters.keys -contains 'silent')){
                            $pltRX10.add('Silent',[boolean]$silent) ;
                        } ;
                    }else{
                        $smsg = "Unable to resolve stock creds: Input suitable creds for OnPrem $($TenOrg):" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $pltRX10 = @{
                            Credential = Get-Credential ; 
                            #verbose = $($verbose) ;
                            Verbose = $FALSE ; 
                        } ;
                        if($silent -AND ((get-command Reconnect-Ex2010).Parameters.keys -contains 'silent')){
                            $pltRX10.add('Silent',[boolean]$silent) ;
                        } ;
                    }; 
                } ; # skip above on CMW, the mods aren't installed
                
            } else {
                $smsg = "(`$useOP:$($UseOP))" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            }  ;  # if-E $UseOP
            #endregion GENERIC_EXOP_CREDS_N_SRVR_CONN #*------^ END GENERIC EXOP CREDS N SRVR CONN BP ^------

            if($pltRX10.Credential){
                $ret_ccOPs.CredentialOP = $pltRX10.Credential ; 
            }else{
                if($UseOP -eq $false){
                    $smsg = "-UseOP:$($useOP): Disabled OnPrem (cross-org?): Expected blank credential set fOR Connect-ExchangeServerTDO!" ; 
                    $smsg += "`n$(($pltRX10|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } else {
                    $smsg = "UNABLE TO RESOLVE A CREDENTIAL SET FOR Connect-ExchangeServerTDO!" ; 
                    $smsg += "`n$(($pltRX10|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    throw $smsg ; 
                    BREAK ; 
                } ; 
            } ; 

        } ; # BEG-E
        PROCESS {

            #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======

            #region USEOP ; #*------v USEOP v------
            if($UseOP){
                #region USEEXOP ; #*------v USEEXOP v------
                if($useEXOP){
                    if($useExopNoDep){ 
                        $smsg = "(Using ExOP:Connect-ExchangeServerTDO(), connect to local ComputerSite)" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;           
                        TRY{
                            $Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name 
                        }CATCH{$Site=$env:COMPUTERNAME} ;
                        $pltCcX10=[ordered]@{
                            siteName = $Site ;
                            RoleNames = @('HUB','CAS') ;
                            verbose  = $($rPSBoundParameters['Verbose'] -eq $true)
                            Credential = $pltRX10.Credential ; 
                        } ;
                        if($ExopVers){
                            $pltCcX10.add('Version',$ExopVers) ; 
                            write-verbose "(Adding specified -Version:$($ExopVers) to `$pltCcX10)"
                        } ; 
                        # 5:15 PM 4/22/2025 on CMW, have to patch version to Ex2016
                        #if($env:userdomain -eq 'CMW'){
                        if($TenOrg -eq 'CMW'){
                            if($pltCcX10.keys -contains 'Version'){
                                $pltCcX10.version = 'Ex2016' ; 
                            } else { $pltCcX10.add('version','Ex2016') } ;
                        } ; 
                        $smsg = "Connect-ExchangeServerTDO w`n$(($pltCcX10|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #$PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
                        $PSSession = Connect-ExchangeServerTDO @pltCcX10 ; 
                    } else {
                        if($pltRX10){
                            #ReConnect-Ex2010XO @pltRX10 ;
                            ReConnect-Ex2010 @pltRX10 ;
                        } else { Reconnect-Ex2010 ; } ;
                        #Add-PSSnapin -Name 'Microsoft.Exchange.Management.PowerShell.SnapIn'
                        #TK: add: test Exch & AD functional connections
                        TRY{
                            if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig'){} else {
                                $smsg = "(mangled Ex10 conn: dx10,rx10...)" ; 
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                disconnect-ex2010 ; reconnect-ex2010 ; 
                            } ; 
                            if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                                $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ; 
                                throw $smsg ; 
                                $smsg | write-warning  ; 
                            } else{
                                $ret_ccOPs.hasExOP = $true ;
                                $ret_ccOPs.xOPPssession = $PSSession ; 

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
                    }
                } ; 
                #endregion USEEXOP ; #*------^ END USEEXOP ^------
                #region USEFORESTWIDE ; #*------v USEFORESTWIDE v------
                if($useForestWide){
                    #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT v------
                    $smsg = "(`$useForestWide:$($useForestWide)):Enabling EXoP Forestwide)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Set-AdServerSettings -ViewEntireForest $True
                    if((get-AdServerSettings).viewentireforest -eq $true){ ;
                        $ret_ccOPs.hasForestWide = $true ;
                    }else{
                        $ret_ccOPs.hasForestWide = $false ;
                    } ;
                    #endregion  ; #*------^ END OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT ^------
                } ;
                #endregion USEFORESTWIDE ; #*------^ END USEFORESTWIDE ^------
            } else {
                $smsg = "(`$useOP:$($UseOP))" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            }  ;  # if-E $UseOP
            #endregion USEOP ; #*------^ END USEOP ^------

            #region UseOPAD #*------v UseOPAD v------
            if($UseOP -OR $UseOPAD){
                if($isNonDomainServer){
                    $smsg = "(non-Domain-connected server:Skipping GENERIC ADMS CONN) "  
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                }else {
                    #region GENERIC_ADMS_CONN_&_XO #*------v GENERIC ADMS CONN & XO  v------
                    $smsg = "(loading ADMS...)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    # always capture load-adms return, it outputs a $true to pipeline on success
                    if($ADMTLoaded = load-ADMS -Verbose:$FALSE){
                        $ret_ccOPs.hasOPAD = $true ; 
                    }
                    # 9:32 AM 4/20/2023 trimmed disabled/fw-borked cross-org code
                    TRY {
                        if(-not($ADForestRoot = Get-ADDomain  -ea STOP).DNSRoot){
                            $smsg = "Missing AD Connection! (no (Get-ADDomain).DNSRoot returned)" ; 
                            throw $smsg ; 
                            $smsg | write-warning  ; 
                            $ret_ccOPs.hasOPAD = $false ; 
                            $ret_ccOPs.ADForestRoot = $null ; 
                        } else{
                            $ret_ccOPs.hasOPAD = $true ; 
                            $ret_ccOPs.ADForestRoot = $ADForestRoot ;  
                        } ; 
                        $objforest = get-adforest -ea STOP ; 
                        # Default new UPNSuffix to the UPNSuffix that matches last 2 elements of the forestname.
                        $forestdom = $UPNSuffixDefault = $objforest.UPNSuffixes | ?{$_ -eq (($objforest.name.split('.'))[-2..-1] -join '.')} ; 
                        $ret_ccOPs.ADforestDom = $forestdom ; 
                        $ret_ccOPs.ADUpnSuffixDefault = $forestdom  ; 
                        if($useForestWide){
                            #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE AD GC QRY SUPPORT v------
                            $smsg = "(`$useForestWide:$($useForestWide)):Enabling AD Forestwide)" ; 
                            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #TK 9:44 AM 10/6/2022 need org wide for rolegrps in parent dom (only for onprem RBAC, not EXO)
                            if($GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268"){
                                $ret_ccOPs.AdGcFwide = $GcFwide ; 
                            }else{
                                $smsg = "UNABLE TO RESOLVE A ForestWide GlobalDomainController (port 3268)!" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            }
                            #endregion  ; #*------^ END  OPTIONAL CODE TO ENABLE FOREST-WIDE AD GC QRY SUPPORT  ^------
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
                    #endregion GENERIC_ADMS_CONN_&_XO #*------^ END GENERIC ADMS CONN & XO ^------
                } ; 
            } else {
                $smsg = "(`$UseOPAD:$($UseOPAD))" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            }  ;
            #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
            #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller = get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
            # use new get-GCFastXO cross-org dc finde
            # default to Op_ExADRoot forest from $TenOrg Meta
            #if($UseOP -AND -not $domaincontroller){
            if($UseOP -AND -not $isNonDomainServer -AND -not (get-variable domaincontroller -ea 0)){
                #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((get-variable -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};
                # need to debug the above, credential issue?
                # just get it done
                $domaincontroller = get-GCFast
            }elseif($isNonDomainServer){
                $smsg = "(non-ADDomain-connected, skipping divert to EXO group resolution)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            }  else { 
                # have to defer to get-azuread, or use EXO's native cmds to poll grp members
                # TODO 1/15/2021
                $useEXOforGroups = $true ; 
                $smsg = "$($TenOrg):HAS NO ON-PREM ACTIVEDIRECTORY, DEFERRING ALL GROUP ACCESS & MGMT TO NATIVE EXO CMDS!" ; 
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
            if(-not $isNonDomainServer -AND ($UseOPAD -OR $UseOP) -AND $useForestWide -AND -not $GcFwide){
                #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT: v------
                $smsg = "`$GcFwide = Get-ADDomainController -Discover -Service GlobalCatalog" ;
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;
                $smsg = "Discovered `$GcFwide:$($GcFwide)" ; 
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #endregion  ; #*------^ END OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT ^------
            } ;
            if($domaincontroller){
                $ret_ccOPs.AdDomainController = $domaincontroller ; 
            } ; 
            #endregion UseOPAD #*------^ END UseOPAD ^------

            <#
            if($VerbosePreference = "Continue"){
                $VerbosePrefPrior = $VerbosePreference ;
                $VerbosePreference = "SilentlyContinue" ;
                $verbose = ($VerbosePreference -eq "Continue") ;
            } ;
            disconnect-exo ;
            if ($script:useEXOv2) { reconnect-eXO2 @pltRXOC }
            else { reconnect-EXO @pltRXOC } ;
            # reenable VerbosePreference:Continue, if set, during mod loads
            if($VerbosePrefPrior -eq "Continue"){
                $VerbosePreference = $VerbosePrefPrior ;
                $verbose = ($VerbosePreference -eq "Continue") ;
            } ;
            #>
            
            #endregion SERVICE_CONNECTIONS #*======^ END SERVICE_CONNECTIONS ^======

        } ; # PROC-E
        END {
            $swM.Stop() ;
            $smsg = ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $swM.Elapsed) ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            <#
            # return status obj
            $ret_ccOPs = [ordered]@{
                CredentialOP = $null ; 
                # OP switches
                hasExOP = $false ;
                ExopVers = $null ;
                hasForestWide = $false ;
                hasOPAD = $false ;
                #
            } ; 
            #>
            $smsg = "Returning connection summary to pipeline:`n$(($ret_ccOPs|out-string).trim())`n" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            [pscustomobject]$ret_ccOPs | write-output ;
        } ; # END-E
    } ;
#} ;
#endregion CONNECT_OPSERVICES ; #*======^ END connect-OPServices ^======
