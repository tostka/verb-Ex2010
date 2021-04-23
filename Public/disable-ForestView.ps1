#*------v disable-ForestView.ps1 v------
Function disable-ForestView {
<#
.SYNOPSIS
disable-ForestView.ps1 - Disable Exchange onprem AD ViewEntireForest setting (permits org-wide object access, wo use of proper explicit -domaincontroller sub.domain.com)
.NOTES
Version     : 1.0.2
Author      : Todd Kadrie
Website     : http://www.toddomation.com
Twitter     : @tostka / http://twitter.com/tostka
CreatedDate : 2020-10-26
FileName    :
License     : MIT License
Copyright   : (c) 2020 Todd Kadrie
Github      : https://github.com/tostka/verb-XXX
Tags        : Powershell
REVISIONS
* 10:56 AM 4/2/2021 cleaned up; added recstat & wlt
* 11:44 AM 3/5/2021 variant of toggle-fv
.DESCRIPTION
disable-ForestView.ps1 - Disable Exchange onprem AD ViewEntireForest setting (permits org-wide object access, wo use of proper explicit -domaincontroller sub.domain.com)
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output
.EXAMPLE
disable-ForestView
.LINK
https://github.com/tostka/verb-ex2010
.LINK
#>
[CmdletBinding()]
PARAM() ;
    # toggle forest view
    if (get-command -name set-AdServerSettings){
        if ((get-AdServerSettings).ViewEntireForest ) {
              write-verbose "(set-AdServerSettings -ViewEntireForest `$False)"
              set-AdServerSettings -ViewEntireForest $False
        } ;
    } else {
        #-=-record a STATUSERROR=-=-=-=-=-=-=
        $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
        if(Get-Variable passstatus -scope Script){$script:PassStatus += $statusdelta } ;
        if(Get-Variable -Name PassStatus_$($tenorg) -scope Script){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
        #-=-=-=-=-=-=-=-=
        $smsg = "MISSING:set-AdServerSettings`nOPEN an Exchange OnPrem connection FIRST!"
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        BREAK ;
    } ;
}

#*------^ disable-ForestView.ps1 ^------