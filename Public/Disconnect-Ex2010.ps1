#*------v Disconnect-Ex2010.ps1 v------
Function Disconnect-Ex2010 {
  <#
    .SYNOPSIS
    Disconnect-Ex2010 - Clear Remote Exch2010 Mgmt Shell connection
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
    Tags        : Powershell,Exchange,ExchangeOnline
    REVISIONS   :
    # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
    * 1:14 PM 3/1/2021 added color reset
    * 4:13 PM 10/22/2020 added pretest of $Global:*'s before running at remove-module (suppresses errors)
    * 12:23 PM 5/27/2020 updated cbh, moved aliases:Disconnect-EMSR','dx10' win func
    * 10:51 AM 2/24/2020 updated attrib
    * 6:59 PM 1/15/2020 cleanup
    * 8:01 AM 11/1/2017 added Remove-PSTitlebar 'EMS', and Disconnect-PssBroken to the bottom - to halt growth of unrepaired broken connections. Updated example to pretest for reqMods
    * 12:54 PM 12/9/2016 cleaned up, add pshelp, implented and debugged as part of verb-Ex2010 set
    * 2:37 PM 12/6/2016 ported to local EMSRemote
    * 2/10/14 posted version
    .DESCRIPTION
    Disconnect-Ex2010 - Clear Remote Exch2010 Mgmt Shell connection
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    $reqMods="Remove-PSTitlebar".split(";") ;
    $reqMods | % {if( !(test-path function:$_ ) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing $($_) function. EXITING." } } ;
    Disconnect-Ex2010 ;
    .LINK
    #>
    [CmdletBinding()]
    [Alias('Disconnect-EMSR','dx10')]
    Param()
    if($Global:E10Mod){$Global:E10Mod | Remove-Module -Force -verbose:$($false) } ;
    if($Global:E10Sess){$Global:E10Sess | Remove-PSSession -verbose:$($false)} ;
    # 7:56 AM 11/1/2017 remove titlebar tag
    Remove-PSTitlebar 'EMS' ;
    # kill any other sessions using distinctive name; add verbose, to ensure they're echo'd that they were missed
    Get-PSSession | Where-Object { $_.name -eq 'Exchange2010' } | Remove-PSSession -verbose:$($false);
    # kill any broken PSS, self regen's even for L13 leave the original borked and create a new 'Session for implicit remoting module at C:\Users\', toast them, they don't reopen. Same for Ex2010 REMS, identical new PSS, indistinguishable from the L13 regen, except the random tmp_xxxx.psm1 module name. Toast them, it's just a growing stack of broken's
    Disconnect-PssBroken ;
    #[console]::ResetColor()  # reset console colorscheme
}

#*------^ Disconnect-Ex2010.ps1 ^------