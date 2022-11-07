#*------v Connect-Ex2010XO.ps1 v------
Function Connect-Ex2010XO {
    <#
    .SYNOPSIS
    Connect-Ex2010XO - Establish PSS to Ex2010, with multi-org support & validation
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-10-15
    FileName    :
    License     :
    Copyright   :
    Github      : https://github.com/tostka
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    AddedCredit : ExactMike Perficient
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    AddedTwitter:
    AddedCredit2 : Jeremy Bradshaw
    AddedWebsite2:	https://github.com/JeremyTBradshaw
    AddedTwitter2:
    REVISIONS   :
    # 3:18 PM 5/18/2021 somehow lost $credOpTORSID, so flipped lost default $credOPTor -> $credTORSID
    # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods; flipped import-psess & import-mod to splats (cleaner) ; line-wrapped longer post-filters for legib
    * 8:30 AM 10/22/2020 ren'd $TentantTag -> $TenOrg, swapped looping meta resolve with 1-liner approach ; added AcceptedDom caching to the middle status test (suppress one more get-exoaccepteddomain call if possible), replaced all $Meta.value with the $TenOrg version
    * 12:56 PM 10/15/2020 converted connect-exo to Ex2010, adding onprem validation
    .DESCRIPTION
    Connect-Ex2010XO - Establish PSS to Ex2010, with multi-org support & validation
    .PARAMETER  ExchangeServer
    On Prem Exch server to Remote to
    .PARAMETER  ExAdmin
    Use exadmin IIS WebPool for remote EMS[-ExAdmin]
    .PARAMETER  Credential
    Credential object
    .PARAMETER  showDebug
    Debugging Flag [-showDebug]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Connect-Ex2010XO
    Connect using defaults, and leverage any pre-set $global:credo365TORSID variable
    .EXAMPLE
    connect-exo -CommandPrefix exo -credential (Get-Credential -credential s-todd.kadrie@torolab.com)  ;
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .EXAMPLE
    $cred = get-credential -credential $o365_Torolab_SIDUpn ;
    connect-exo -credential $cred ;
    Pass in a prefab credential object (useful for auto-shifting to MFA - the function will autoresolve MFA reqs based on the cred domain)
    .LINK
    .LINK
    https://github.com/verb-Exch2010
    #>
    [CmdletBinding()]
    [Alias('cxoxo')]
    Param(
        [Parameter(Position = 0, HelpMessage = "Exch server to Remote to")]
        [string]$ExchangeServer,
        [Parameter(HelpMessage = 'Use variant IIS WebPool for remote EMS[-ExAdmin]')]
        $ExAdmin,
        [Parameter(HelpMessage = 'Credential object')]
        [System.Management.Automation.PSCredential]$Credential = $credTORSID,
        [Parameter(HelpMessage = "Debugging Flag [-showDebug]")]
        [switch] $showDebug
    ) ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ;
        #if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        # $rgxEx10HostName : ^(lyn|bcc|adl|spb)ms6[4,5][0,1].global.ad.toro.com$
        # we'd need to define all possible hostnames to cover potential span. Should probably build dynamically from $XXXMeta vari
        # can build from $TorMeta.OP_ExADRoot:global.ad.toro.com
        <# on curly, from Ps into EMS:
        get-pssession | fl computername,computertype,state,configurationname,availability,name
        ComputerName      : curlyhoward.cmw.internal
        ComputerType      : RemoteMachine
        State             : Opened
        ConfigurationName : Microsoft.Exchange
        Availability      : Available
        Name              : Session1

        ComputerName      : lynms650.global.ad.toro.com
        ComputerType      : RemoteMachine
        State             : Broken
        ConfigurationName : Microsoft.Exchange
        Availability      : None
        Name              : Exchange2010

        "^\w*\.$($CMWMeta.OP_ExADRoot)$"
        => ^\w*\.cmw.internal$
        #>

        #$sTitleBarTag = "EMS" ;
        $CommandPrefix = $null ;

        $TenOrg=get-TenantTag -Credential $Credential ;
        <#if($TenOrg -ne 'TOR'){
            # explicitly leave this tenant (default) untagged
            $sTitleBarTag += $TenOrg ;
        } ;
        #>
        if($TenOrg){
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

        <#
        $credDom = ($Credential.username.split("\"))[0] ;
        $Metas=(get-variable *meta|Where-Object{$_.name -match '^\w{3}Meta$'}) ;
        foreach ($Meta in $Metas){
            if( ($credDom -eq $Meta.value.legacyDomain) -OR ($credDom -eq $Meta.value.o365_TenantDomain) -OR ($credDom -eq $Meta.value.o365_OPDomain)){
                if($Meta.value.OP_ExADRoot){
                    if(!$Meta.value.OP_rgxEMSComputerName){
                        write-verbose "(adding XXXMeta.OP_rgxEMSComputerName value)"
                        # build vari that will match curlyhoward.cmw.internal|lynms650.global.ad.toro.com etc
                        set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'OP_rgxEMSComputerName' = "^\w*\.$([Regex]::Escape($Meta.value.OP_ExADRoot))$"} ) ;
                    } ;
                } else {
                    throw "Missing `$$($Meta.value.o365_Prefix).OP_ExADRoot value.`nProfile hasn't loaded proper tor-incl-infrastrings file)!"
                } ;
            } ; # if-E $credDom
        } ; # loop-E
        #>
        # non-looping vers:
        #$TenOrg = get-TenantTag -Credential $Credential ;
        #.OP_ExADRoot
        if( (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName ){

        } else {
            #.OP_rgxEMSComputerName
            if((Get-Variable  -name "$($TenOrg)Meta").value.OP_ExADRoot){
                set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'OP_rgxEMSComputerName' = "^\w*\.$([Regex]::Escape((Get-Variable  -name "$($TenOrg)Meta").value.OP_ExADRoot))$"} )
            } else {
                $smsg = "Missing `$$((Get-Variable  -name "$($TenOrg)Meta").value.o365_Prefix).OP_ExADRoot value.`nProfile hasn't loaded proper tor-incl-infrastrings file)!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;
        } ;

    } ;  # BEG-E
    PROCESS{
        # if we're using ems-style BasicAuth, clear incompatible existing Rems PSS's
        # ComputerName      : curlyhoward.cmw.internal ;  ComputerType      : RemoteMachine ;  State             : Opened ;  ConfigurationName : Microsoft.Exchange ;  Availability      : Available ;  Name              : Session1 ;   ;
        $rgxRemsPSSName = "^(Session\d|Exchange\d{4})$" ;
        $Rems2Good = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') } ;
        # Computername wrong fqdn suffix
        #$Rems2WrongOrg = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND (-not($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName)) -AND ($_.Availability -eq 'Available') } ;
        # above is seeing outlook EXO conns as wrong org, exempt them too: .ComputerName -match $rgxExoPsHostName
        $Rems2WrongOrg = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND (
            $_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND (
            ( -not($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) ) -AND (
            -not($_.ComputerName -match $rgxExoPsHostName)) ) -AND ($_.Availability -eq 'Available')
        } ;
        $Rems2Broken = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND (
            $_.Name -match $rgxRemsPSSName) -AND ($_.State -like "*Broken*") } ;
        $Rems2Closed = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND (
            $_.Name -match $rgxRemsPSSName) -AND ($_.State -like "*Closed*") } ;

        if ($Rems2Broken.count -gt 0){ for ($index = 0 ;$index -lt $Rems2Broken.count ;$index++){Remove-PSSession -session $Rems2Broken[$index]}  };
        if ($Rems2Closed.count -gt 0){for ($index = 0 ;$index -lt $Rems2Closed.count ; $index++){Remove-PSSession -session $Rems2Closed[$index] } } ;
        if ($Rems2WrongOrg.count -gt 0){for ($index = 0 ;$index -lt $Rems2WrongOrg.count ; $index++){Remove-PSSession -session $Rems2WrongOrg[$index] } } ;
        # preclear until proven *up*
        $bExistingREms = $false ;

        if( Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') } ){
            $bExistingREms= $true ;

        } ;
        if($bExistingREms -eq $false){
            #$TorMeta.Ex10Server: dynamic
            #$TorMeta.Ex10ServerXO: lynms650.global.ad.toro.com
            # force unresolved to dyn
            if((Get-Variable  -name "$($TenOrg)Meta").value.Ex10ServerXO){
                write-host -foregroundcolor darkgray "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Adding EMS (connecting to $($TorMeta.Ex10ServerXO))..." ;
            } ;

            $pltNSess = @{
                ConnectionURI = "http://$((Get-Variable  -name "$($TenOrg)Meta").value.Ex10ServerXO)/powershell";
                ConfigurationName = 'Microsoft.Exchange' ;
                name = 'Exchange2010' ;
            } ;
            if ((Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant) {
              # use variant IIS Webpool
              $pltNSess.ConnectionURI = $pltNSess.ConnectionURI.replace("/powershell", "/$((Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant)") ;
            }
            $pltNSess.Add("Credential", $Credential); # just use the passed $Credential vari
            $cMsg = "Connecting to OP Ex20XX ($($credDom))";
            Write-Host $cMsg ;
            write-verbose "`n$((get-date).ToString('HH:mm:ss')):New-PSSession w`n$(($pltNSess|out-string).trim())" ;

            $error.clear() ;
            TRY { $global:E10Sess = New-PSSession @pltNSess -ea STOP
            } CATCH {
                $ErrTrapd = $_ ;
                write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ((Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant) {
                  # switch to stock pool and retry
                  $pltNSess.ConnectionURI = $pltNSess.ConnectionURI.replace("/$((Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant)", "/powershell") ;
                  write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TARGETING VARIANT POOL`nRETRY W STOCK POOL: New-PSSession w`n$(($pltNSess|out-string).trim())" ;
                  $global:E10Sess = New-PSSession @pltNSess -ea STOP  ;
                } else {
                    STOP ;
                } ;
            } ; # try-E

            if(!$global:E10Sess){
                write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO RETURN PSSESSION!`nAUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                Break ;
            } ;

            $pltIMod=@{Global = $true ;PassThru = $true;DisableNameChecking = $true ; verbose=$true ;} ;
            $pltISess = [ordered]@{
                Session             = $global:E10Sess ;
                DisableNameChecking = $true  ;
                AllowClobber        = $true ;
                ErrorAction         = 'Stop' ;
                Verbose             = $false ;
            } ;
            if ($CommandPrefix) {
                write-host -foregroundcolor white "$((get-date).ToString("HH:mm:ss")):Note: Prefixing this Mod's Cmdlets as [verb]-$($CommandPrefix)[noun]" ;
                $pltIMod.add('Prefix',$CommandPrefix) ;
                $pltISess.add('Prefix',$CommandPrefix) ;
            } ;
            write-verbose "`n$((get-date).ToString('HH:mm:ss')):Import-PSSession w`n$(($pltISess|out-string).trim())`nImport-Module w`n$(($pltIMod|out-string).trim())" ;

            # Verbose:Continue is VERY noisey for module loads. Bracketed suppress:
            if($VerbosePreference = "Continue"){
                $VerbosePrefPrior = $VerbosePreference ;
                $VerbosePreference = "SilentlyContinue" ;
                $verbose = ($VerbosePreference -eq "Continue") ;
            } ;
            Try {
                $Global:E10Mod = Import-Module (Import-PSSession @pltISess) @pltIMod  ;
                #$Global:EOLModule = Import-Module (Import-PSSession @pltISess) -Global -Prefix $CommandPrefix -PassThru -DisableNameChecking   ;
            } catch {
                Write-Warning -Message "Tried but failed to import the EXO PS module.`n`nError message:" ;
                throw $_ ;
            } ;
            # reenable VerbosePreference:Continue, if set, during mod loads
            if($VerbosePrefPrior -eq "Continue"){
                $VerbosePreference = $VerbosePrefPrior ;
                $verbose = ($VerbosePreference -eq "Continue") ;
            } ;
            Add-PSTitleBar $sTitleBarTag -verbose:$($VerbosePreference -eq "Continue")  ;

        } ; #  # if-E $bExistingREms
    } ;  # PROC-E
    END {
        if($bExistingREms -eq $false){
            if( Get-PSSession | where-object {$_.ConfigurationName -eq "Microsoft.Exchange" -AND $_.Name -match $rgxRemsPSSName -AND $_.State -eq "Opened" -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') }  ){
                $bExistingREms= $true ;
            } else {
                write-error "(Credential mismatch:disconnecting from existing Ex201X:$($eEXO.Identity) tenant)" ;
                Disconnect-Ex2010 ;
                $bExistingREms = $false ;
            } ;
        } ;
    } ; # END-E
}

#*------^ Connect-Ex2010XO.ps1 ^------
