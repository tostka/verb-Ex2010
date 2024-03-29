﻿#*------v cx10cmw.ps1 v------
function cx10cmw {
    <#
    .SYNOPSIS
    cx10cmw - Connect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Connect-EX2010 - Connect-EX2010 to specified on-prem Exchange
    .EXAMPLE
    cx10cmw
    #>
    [CmdletBinding()] 
    [Alias('cxOPcmw')]
    Param()
    $Verbose = ($VerbosePreference -eq 'Continue') ;
    $pltGHOpCred=@{TenOrg="CMW" ;userrole=@('ESVC','LSVC','SID') ;verbose=$($verbose)} ;
    if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
        Connect-EX2010 -cred $OPCred -Verbose:($VerbosePreference -eq 'Continue') ; 
    } else {
        $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole $($UserRole -join '|') value!"
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
        exit ;
    } ;
}

#*------^ cx10cmw.ps1 ^------
