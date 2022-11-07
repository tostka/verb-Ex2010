#*------v Reconnect-Ex2010XO.ps1 v------
Function Reconnect-Ex2010XO {
   <#
    .SYNOPSIS
    Reconnect-Ex2010XO - Reconnect Remote Exch2010 Mgmt Shell connection Cross-Org (XO)
    .NOTES
    Updated By: : Todd Kadrie
    Website:	http://toddomation.com
    Twitter:	@tostka http://twitter.com/tostka
    Based on original function Author: ExactMike Perficient, Global Knowl... (Partner)
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    REVISIONS   :
    # 3:18 PM 5/18/2021 somehow lost $credOpTORSID, so flipped lost default $credOPTor -> $credTORSID
    * 1:57 PM 3/31/2021 wrapped long lines for vis
    * 8:30 AM 10/22/2020 ren'd $TentantTag -> $TenOrg, swapped looping meta resolve with 1-liner approach ; added AcceptedDom caching to the middle status test (suppress one more get-exoaccepteddomain call if possible), replaced all $Meta.value with the $TenOrg version
    * 1:19 PM 10/15/2020 converted connect-exo to Ex2010, adding onprem validation
    .DESCRIPTION
    Reconnect-Ex2010XO - Reconnect Remote Exch2010 Mgmt Shell connection Cross-Org (XO)
    .PARAMETER  Credential
    Credential to use for this connection [-credential 'account@domain.com']
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Reconnect-Ex2010XO;
    Reconnect EXO connection
    .EXAMPLE
    $batchsize = 10 ;
    $RecordCount=$mr.count #this is the array of whatever you are processing ;
    $b=0 #this is the initialization of a variable used in the do until loop below ;
    $mrs = @() ;
    do {
        Write-Progress -Activity "Getting move request statistics for all $wave move requests." -Status "Processing Records $b through $($b+$batchsize) of $RecordCount." -PercentComplete ($b/$RecordCount*100) ;
        $mrs += $mr | Select-Object -skip $b -first $batchsize | foreach-object {Reconnect-Ex2010XO; $_ | Get-OLMoveRequestStatistics} ;
        $b=$b+$batchsize ;
        } ;
    until ($b -gt $RecordCount) ;

    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    #>
    [CmdletBinding()]
    [Alias('rx10xo')]
    <#
    Param(
      [Parameter(HelpMessage="Use Proxy-Aware SessionOption settings [-ProxyEnabled]")]
      [boolean]$ProxyEnabled = $False,
      [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")]
      [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
      [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
      [switch] $showDebug
    ) ;
    #>
     Param(
        [Parameter(HelpMessage = 'Credential object')][System.Management.Automation.PSCredential]$Credential = $credTORSID,
        [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
        [switch] $showDebug
    )  ;
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

        $sTitleBarTag = "EMS" ;
        $CommandPrefix = $null ;

        $TenOrg=get-TenantTag -Credential $Credential ;
        if($TenOrg -ne 'TOR'){
            # explicitly leave this tenant (default) untagged
            $sTitleBarTag += $TenOrg ;
        } ;
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
        $verbose = ($VerbosePreference -eq "Continue") ;
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

        write-verbose "(Removing $($Rems2Broken.count) Broken sessions)" ;
        if ($Rems2Broken.count -gt 0){ for ($index = 0 ;$index -lt $Rems2Broken.count ;$index++){Remove-PSSession -session $Rems2Broken[$index]}  };
        write-verbose "(Removing $($Rems2Closed.count) Closed sessions)" ;
        if ($Rems2Closed.count -gt 0){for ($index = 0 ;$index -lt $Rems2Closed.count ; $index++){Remove-PSSession -session $Rems2Closed[$index] } } ;
        write-verbose "(Removing $($Rems2WrongOrg.count) sessions connected to the WRONG ORG)" ;
        if ($Rems2WrongOrg.count -gt 0){for ($index = 0 ;$index -lt $Rems2WrongOrg.count ; $index++){Remove-PSSession -session $Rems2WrongOrg[$index] } } ;
        # preclear until proven *up*
        $bExistingREms = $false ;

        if( Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') } ){

            $bExistingREms= $true ;
            write-verbose "(Authenticated to Ex20XX:$($Credential.username.split('\')[0].tostring()))" ;

        } else {
            write-verbose "(NOT Authenticated to Credentialed Ex20XX Org:$($Credential.username.split('\')[0].tostring()))" ;
            $tryNo=0 ; $1F=$false ;
            Do {
                if($1F){Sleep -s 5} ;
                $tryNo++ ;
                write-host "." -NoNewLine; if($tryNo -gt 1){Start-Sleep -m (1000 * 5)} ;
                write-verbose "$((get-date).ToString('HH:mm:ss')):Reconnecting:No existing PSSESSION matching`n (ConfigurationName -eq 'Microsoft.Exchange') -AND (Name -match $($rgxRemsPSSName)) -AND ($_.ComputerName -match $((Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName))`nwith valid Open/Availability:$((Get-PSSession | where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ($_.Name -match $rgxRemsPSSName)} |ft -a Id,Name,ComputerName,ComputerType,State,ConfigurationName,Availability|out-string).trim())" ;
                Disconnect-Ex2010 ; Disconnect-PssBroken ;Start-Sleep -Seconds 3;

                $bExistingREms = $false ;

                Connect-Ex2010xo -credential:$($Credential) ;

                $1F=$true ;
                if($tryNo -gt $DoRetries ){throw "RETRIED EX20XX CONNECT $($tryNo) TIMES, ABORTING!" } ;
            } Until ( Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') } ) ;

        } ;

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

#*------^ Reconnect-Ex2010XO.ps1 ^------
