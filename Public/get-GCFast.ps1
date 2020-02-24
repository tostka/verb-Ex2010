#*----------------v Function get-GCFast v----------------
function get-GCFast {

  <#
    .SYNOPSIS
    get-GCFast - function to locate a random sub-100ms response gc in specified domain & optional AD site
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter:	http://twitter.com/tostka
    Additional Credits: Originated in Ben Lye's GetLocalDC()
    Website:	http://www.onesimplescript.com/2012/03/using-powershell-to-find-local-domain.html
    REVISIONS   :
    # 2:19 PM 4/29/2019 add [lab dom] to the domain param validateset & site lookup code, also copied into tsksid-incl-ServerCore.ps1
    # 2:39 PM 8/9/2017 ADDED some code to support labdom.com, also added test that $LocalDcs actually returned anything!
    # 10:59 AM 3/31/2016 fix site param valad: shouln't be sitecodes, should be Site names; updated Site param def, to validate, cleanup, cleaned up old remmed code, rearranged comments a bit
    # 1:12 PM 2/11/2016 fixed new bug in get-GCFast, wasn't detecting blank $site, for PSv2-compat, pre-ensure that ADMS is loaded
    12:32 PM 1/8/2015 - tweaked version of Ben lye's script, replaced broken .NET site query with get-addomaincontroller ADMT module command
    .PARAMETER  Domain
    Which AD Domain [Domain fqdn]
    .PARAMETER  Site
    DCs from which Site name (defaults to AD lookup against local computer's Site)
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns one DC object, .Name is name pointer
    .EXAMPLE
    C:\> get-gcfast -domain dom.for.domain.com -site Site
    Lookup a Global domain gc, with Site specified (whether in Site or not, will return remote site dc's)
    .EXAMPLE
    C:\> get-gcfast -domain dom.for.domain.com
    Lookup a Global domain gc, default to Site lookup from local server's perspective
  #>

  [CmdletBinding()]
  param(
    [Parameter(HelpMessage = 'Target AD Domain')]
    [string]$Domain
    , [Parameter(Position = 1, Mandatory = $False, HelpMessage = "Optional: DCs from what Site name? (default=Discover)")]
    [string]$Site
  ) ;
  $SpeedThreshold = 100 ;
  $ErrorActionPreference = 'SilentlyContinue' ; # Set so we don't see errors for the connectivity test
  $env:ADPS_LoadDefaultDrive = 0 ; $sName = "ActiveDirectory"; if ( !(Get-Module | Where-Object { $_.Name -eq $sName }) ) {
    if ($bDebug) { Write-Debug "Adding ActiveDirectory Module (`$script:ADPSS)" };
    $script:AdPSS = Import-Module $sName -PassThru -ea Stop ;
  } ;
  if (!$Domain) {
    $Domain = (get-addomain).DNSRoot ; # use local domain
    write-host -foregroundcolor yellow   "Defaulting domain: $Domain";
  }
  # Get all the local domain controllers
  if ((!$Site)) {
    # if no site, look the computer's Site Up in AD
    $Site = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name ;
    write-host -foregroundcolor yellow   "Using local machine Site: $Site";
  } ;

  # gc filter
  #$LocalDCs = Get-ADDomainController -filter { (isglobalcatalog -eq $true) -and (Site -eq $Site) } ;
  #$LocalDCs = Get-ADDomainController -filter { (isglobalcatalog -eq $true) -and (Site -eq $Site) } ;
  $LocalDCs = Get-ADDomainController -filter { (isglobalcatalog -eq $true) -and (Site -eq $Site) -and (Domain -eq $Domain) } ;
  # any dc filter
  #$LocalDCs = Get-ADDomainController -filter {(Site -eq $Site)} ;

  $PotentialDCs = @() ;
  # Check connectivity to each DC against $SpeedThreshold
  if ($LocalDCs) {
    foreach ($LocalDC in $LocalDCs) {
      $TCPClient = New-Object System.Net.Sockets.TCPClient ;
      $Connect = $TCPClient.BeginConnect($LocalDC.Name, 389, $null, $null) ;
      $Wait = $Connect.AsyncWaitHandle.WaitOne($SpeedThreshold, $False) ;
      if ($TCPClient.Connected) {
        $PotentialDCs += $LocalDC.Name ;
        $Null = $TCPClient.Close() ;
      } # if-E
    } ;
    write-host -foregroundcolor yellow  "`$PotentialDCs: $PotentialDCs";
    $DC = $PotentialDCs | Get-Random ;
    write-output $DC  ;
  }
  else {
    write-host -foregroundcolor yellow  "NO DCS RETURNED BY GET-GCFAST()!";
    write-output $false ;
  } ;
} #*----------------^ END Function get-GCFast ^----------------