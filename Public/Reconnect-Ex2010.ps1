#*------v Reconnect-Ex2010.ps1 v------
Function Reconnect-Ex2010 {
  <#
    .SYNOPSIS
    Reconnect-Ex2010 - Reconnect Remote ExchOnPrem Mgmt Shell connection (validated functional Exch2010 - Exch2016)
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    AddedCredit : Inspired by concept code by ExactMike Perficient, Global Knowl... (Partner)
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Version     : 1.1.0
    CreatedDate : 2020-02-24
    FileName    : Reonnect-Ex2010()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell
    REVISIONS   :
    * 10:52 AM 4/2/2021 updated cbh
    * 1:56 PM 3/31/2021 rewrote to dyn detect pss, rather than reading out of date vari
    * 10:14 AM 3/23/2021 fix default $Cred spec, pointed at an OP cred
    * 8:29 AM 11/17/2020 added missing $Credential param 
    * 9:33 AM 5/28/2020 actually added the alias:rx10 
    * 12:20 PM 5/27/2020 updated cbh, moved alias: rx10 win func
    * 6:59 PM 1/15/2020 cleanup
    * 8:09 AM 11/1/2017 updated example to pretest for reqMods
    * 1:26 PM 12/9/2016 split no-session and reopen code, to suppress notfound errors, add pshelpported to local EMSRemote
    * 2/10/14 posted version
    .DESCRIPTION
    Reconnect-Ex2010 - Reconnect Remote Exch2010 Mgmt Shell connection
    .PARAMETER  Credential
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    $reqMods="connect-Ex2010;Disconnect-Ex2010;".split(";") ;
    $reqMods | % {if( !(test-path function:$_ ) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing $($_) function. EXITING." } } ;
    Reconnect-Ex2010 ;
    .LINK
    #>
    [CmdletBinding()]
    [Alias('rx10','rxOP','reconnect-ExOP')]
    Param(
        [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")][System.Management.Automation.PSCredential]
        $Credential = $global:credOpTORSID 
    )
    # checking stat on canned copy of hist sess, says nothing about current, possibly timed out, check them manually
    $rgxRemsPSSName = "^(Session\d|Exchange\d{4})$" ;
    # back the TenOrg out of the Credential
    $TenOrg = get-TenantTag -Credential $Credential ;

    $Rems2Good = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') } ;
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
    if(!$E10Sess){
        if (!$Credential) {
            Connect-Ex2010
        } else {
            Connect-Ex2010 -Credential:$($Credential) ;
        } ;
    } elseif($tSess = get-pssession -id $e10sess.id -ea 0 |?{$_.computername -eq $e10sess.computername -ANd $_.name -eq $e10sess.name}){
        # matches historical session
        if( $tSess | where-object { ($_.State -eq "Opened") -AND ($_.Availability -eq 'Available') } ){
            $bExistingREms= $true ;
        } else {
            $bExistingREms= $false ;
        } ;
    } else { 
        # doesn't match histo
        $bExistingREms= $false ;
    } ; 
    $propsPss =  'Id','Name','ComputerName','ComputerType','State','ConfigurationName','Availability' ; 
    if($bExistingREms){
        $smsg = "existing connection Open/Available:`n$(($tSess| ft -auto $propsPss |out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    } else { 
        $smsg = "(resetting any existing EX10 connection and re-establishing)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        Disconnect-Ex2010 ; Start-Sleep -S 3;
        if (!$Credential) {
            Connect-Ex2010 ;
        } else {
            Connect-Ex2010 -Credential:$($Credential) ;
        } ;
    } ; 
}

#*------^ Reconnect-Ex2010.ps1 ^------