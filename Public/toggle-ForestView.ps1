#*------v toggle-ForestView.ps1 v------
Function toggle-ForestView {
<#
.SYNOPSIS
toggle-ForestView.ps1 - Toggle Exchange onprem AD ViewEntireForest setting (permits org-wide object access, wo use of proper explicit -domaincontroller sub.domain.com)
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
* 10:53 AM 4/2/2021 typo fix
* 10:07 AM 10/26/2020 added CBH
.DESCRIPTION
toggle-ForestView.ps1 - Toggle Exchange onprem AD ViewEntireForest setting (permits org-wide object access, wo use of proper explicit -domaincontroller sub.domain.com)
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output
.EXAMPLE
toggle-ForestView
.LINK
https://github.com/tostka/verb-ex2010
.LINK
#>
[CmdletBinding()]
PARAM() ;
    # toggle forest view
    if (get-command -name set-AdServerSettings){ 
        if (!(get-AdServerSettings).ViewEntireForest ) {
              write-warning "Enabling WholeForest"
              write-host "`a"
              if (get-command -name set-AdServerSettings -ea 0) { set-AdServerSettings -ViewEntireForest $TRUE } ;
        } else {
          write-warning "Disabling WholeForest"
          write-host "`a"
          set-AdServerSettings -ViewEntireForest $FALSE ;
        } ;
    } else {
        THROW "MISSING:set-AdServerSettings`nOPEN an Exchange OnPrem connection FIRST!"
    } ; 
}
#*------^ toggle-ForestView.ps1 ^------