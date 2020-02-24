#*------v Function Reconnect-Ex2010 v------
Function Reconnect-Ex2010 {
  <#
    .SYNOPSIS
    Reconnect-Ex2010 - Reconnect Remote Exch2010 Mgmt Shell connection
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter:	http://twitter.com/tostka
    Author: ExactMike Perficient, Global Knowl... (Partner)
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    REVISIONS   :
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
  if (!$E10Sess) {
    if (!$Credential) {
      Connect-Ex2010
    }
    else {
      Connect-Ex2010 -Credential:$($Credential) ;
    } ;
  }
  elseif ($E10Sess.state -ne 'Opened' -OR $E10Sess.Availability -ne 'Available' ) {
    Disconnect-Ex2010 ; Start-Sleep -S 3;
    if (!$Credential) {
      Connect-Ex2010
    }
    else {
      Connect-Ex2010 -Credential:$($Credential) ;
    } ;
  } ;
}#*------^ END Function Reconnect-Ex2010 ^------ ;
if (!(get-alias rx10 -ea 0)) { set-alias -name rx10 -value Reconnect-Ex2010 } ;