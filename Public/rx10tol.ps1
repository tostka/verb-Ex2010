#*------v rx10tol.ps1 v------
function rx10tol {
    <#
    .SYNOPSIS
    rx10tol - Reonnect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Reconnect-EX2010 - Reonnect to specified on-prem Exchange
    .EXAMPLE
    rx10tol
    #>
    [CmdletBinding()] 
    [Alias('rxOPtol')]
    Param()
    $Verbose = ($VerbosePreference -eq 'Continue') ;
    $pltGHOpCred=@{TenOrg="TOL" ;userrole=@('ESVC','LSVC','SID') ;verbose=$($verbose)} ;
    if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
        ReConnect-EX2010 -cred $OPCred -Verbose:($VerbosePreference -eq 'Continue') ; 
    } else {
        $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole $($UserRole -join '|') value!"
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
        exit ;
    } ;
}

#*------^ rx10tol.ps1 ^------