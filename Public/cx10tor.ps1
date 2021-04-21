#*------v cx10tor.ps1 v------
function cx10tor {
    <#
    .SYNOPSIS
    cx10tor - Connect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Connect-EX2010 - Connect-EX2010 to specified on-prem Exchange
    .NOTES
    REVISIONS   :
    # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
    .EXAMPLE
    cx10tor
    #>
    [CmdletBinding()] 
    [Alias('cxOPtor')]
    Param([Parameter(HelpMessage = 'Credential object')][System.Management.Automation.PSCredential]$Credential = $credTorSID)
    $Verbose = ($VerbosePreference -eq 'Continue') ;
    if(!$Credential){
        $pltGHOpCred=@{TenOrg="TOR" ;userrole=@('SID','ESVC','LSVC') ;verbose=$($verbose)} ;
        if($Credential=(get-HybridOPCredentials @pltGHOpCred).cred){
            #Connect-EX2010 -cred $credTorSID #-Verbose:($VerbosePreference -eq 'Continue') ; 
        } else {
            $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole $($UserRole -join '|') value!"
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Break ;
        } ;
    } ; 
    Connect-EX2010 -cred $Credential #-Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ cx10tor.ps1 ^------