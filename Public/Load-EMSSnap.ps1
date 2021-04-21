#*------v Load-EMSSnap.ps1 v------
function Load-EMSSnap {
  <#
    .SYNOPSIS
    Checks local machine for registred Exchange2010 EMS, and loads the component
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter:	http://twitter.com/tostka

    REVISIONS   :
    # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods/add-Pssnapins
    * 6:59 PM 1/15/2020 cleanup
    vers: 9:39 AM 8/12/2015: retool into generic switched version to support both modules & snappins with same basic code ; building a stock EMS version (vs the fancier load-EMSSnapLatest)
    vers: 10:43 AM 1/14/2015 fixed return & syntax expl to true/false
    vers: 10:20 AM 12/10/2014 moved commentblock into function
    vers: 11:40 AM 11/25/2014 adapted to Lync
    ers: 2:05 PM 7/19/2013 typo fix in 2013 code
    vers: 1:46 PM 7/19/2013
    .INPUTS
    None.
    .OUTPUTS
    Outputs $true if successful. $false if failed.
    .EXAMPLE
    $EMSLoaded = Load-EMSSnap ; Write-Debug "`$EMSLoaded: $EMSLoaded" ;
    Stock free-standing Exchange Mgmt Shell load
    .EXAMPLE
    $EMSLoaded = Load-EMSSnap ; Write-Debug "`$EMSLoaded: $EMSLoaded" ; get-exchangeserver | out-null ;
    Example utilizing a workaround for bug in EMS, where loading ADMS causes Powershell/ISE to crash if ADMS is loaded after EMS, before EMS has executed any commands
    .EXAMPLE
    TRY {
        if(($host.version.major -lt 3) -AND (get-service MSExchangeADTopology -ea SilentlyContinue)){
                write-verbose -verbose:$bshowVerbose  "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Using Local Server EMS10 Snapin" ;
                $sName="Microsoft.Exchange.Management.PowerShell.E2010"; if (!(Get-PSSnapin | where {$_.Name -eq $sName})) {Add-PSSnapin $sName -ea Stop} ;
        } else {
             write-verbose -verbose:$bshowVerbose  "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Initiating REMS connection" ;
            $reqMods="connect-Ex2010;Disconnect-Ex2010;".split(";") ;
            $reqMods | % {if( !(test-path function:$_ ) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing $($_) function. EXITING." } } ;
            Reconnect-Ex2010 ;
        } ;
    } CATCH {
        Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
        Exit #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ;
    } ;
    Example demo'ing check for local psv2 & ADtopo svc to defer
    #>

  # check registred v loaded ;
  # style of plugin we want to test/load
  $PlugStyle = "Snapin"; # for Exch EMS
  #"Module" ; # for Lync/ADMS
  $PlugName = "Microsoft.Exchange.Management.PowerShell.E2010" ;

  switch ($PlugStyle) {
    "Module" {
      # module-style (for LMS or ADMS
      $PlugsReg = Get-Module -ListAvailable;
      $PlugsLoad = Get-Module;
    }
    "Snapin" {
      $PlugsReg = Get-PSSnapin -Registered ;
      $PlugsLoad = Get-PSSnapin ;
    }
  } # switch-E

  TRY {
    if ($PlugsReg | where { $_.Name -eq $PlugName }) {
      if (!($PlugsLoad | where { $_.Name -eq $PlugName })) {
        #
        switch ($PlugStyle) {
          "Module" {
            Import-Module $PlugName -ErrorAction Stop -verbose:$($false); write-output $TRUE ;
          }
          "Snapin" {
            Add-PSSnapin $PlugName -ErrorAction Stop -verbose:$($false); write-output $TRUE
          }
        } # switch-E
      }
      else {
        # already loaded
        write-output $TRUE;
      } # if-E
    }
    else {
      Write-Error { "$(Get-TimeStamp):($env:computername) does not have $PlugName installed!"; };
      #return $FALSE ;
      write-output $FALSE ;
    } # if-E ;
  } CATCH {
        Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
        Exit #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ;
  } ;

}

#*------^ Load-EMSSnap.ps1 ^------