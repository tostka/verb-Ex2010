#*-----v Function load-EMSLatest v-----
function load-EMSLatest {
  #  #Checks local machine for registred E20[13|10|07] EMS, and then loads the newest one found
  #Returns the string 2013|2010|2007 for reuse for version-specific code

  <#
  .SYNOPSIS
  load-EMSLatest - Checks local machine for registred E20[13|10|07] EMS, and then loads the newest one found.
  Attempts remote Ex2010 connection if no local EMS installed
  Returns the string 2013|2010|2007 for reuse for version-specific code
    .NOTES
  Author: Todd Kadrie
  Website:	http://tinstoys.blogspot.com
  Twitter:	http://twitter.com/tostka
  REVISIONS   :
  * 6:59 PM 1/15/2020 cleanup
  9:39 AM 2/4/2015 updated to remote to a local hub, updated latest TOR
    .INPUTS
  None. Does not accepted piped input.
    .OUTPUTS
  Returns version number connected to: [2013|2010|2007]
    .EXAMPLE
  .\load-EMSLatest
    .LINK
  #>

  # check registred & loaded ;
  $SnapsReg = Get-PSSnapin -Registered ;
  $SnapsLoad = Get-PSSnapin ;
  $Snapin13 = "Microsoft.Exchange.Management.PowerShell.E2013";
  $Snapin10 = "Microsoft.Exchange.Management.PowerShell.E2010";
  $Snapin7 = "Microsoft.Exchange.Management.PowerShell.Admin";
  # check/load E2013, E2010, or E2007, stop at newest (servers wouldn't be running multi-versions)
  if (($SnapsReg | where { $_.Name -eq $Snapin13 })) {
    if (!($SnapsLoad | where { $_.Name -eq $Snapin13 })) {
      Add-PSSnapin $Snapin13 -ErrorAction SilentlyContinue ; return "2013" ;
    }
    else {
      return "2013" ;
    } # if-E
  }
  elseif (($SnapsReg | where { $_.Name -eq $Snapin10 })) {
    if (!($SnapsLoad | where { $_.Name -eq $Snapin10 })) {
      Add-PSSnapin $Snapin10 -ErrorAction SilentlyContinue ; return "2010" ;
    }
    else {
      return "2010" ;
    } # if-E
  }
  elseif (($SnapsReg | where { $_.Name -eq $Snapin7 })) {
    if (!($SnapsLoad | where { $_.Name -eq $Snapin7 })) {
      Add-PSSnapin $Snapin7 -ErrorAction SilentlyContinue ; return "2007" ;
    }
    else {
      return "2007" ;
    } # if-E
  }
  else {
    Write-Verbose "Unable to locate Exchange tools on localhost, attempting to remote to Exchange 2010 server...";
    #Try implicit remoting-only works for Exchange 2010
    Try {
      # connect to a local hub (leverages ADSI function)
      $Ex2010Server = (Get-ExchangeServerInSite | ? { $_.Roles -match "^(36|38)$" })[0].fqdn
      $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Ex2010Server/PowerShell/ -ErrorAction Stop ;
      Import-PSSession $ExchangeSession -ErrorAction Stop;
    }
    Catch {
      Write-Host -ForegroundColor Red "Unable to import Exchange tools from $Exchange2010Server, is it running Exchange 2010?" ;
      Write-Host -ForegroundColor Magenta "Error:  $($Error[0])" ;
      Exit;
    } # try-E
  }# if-E
} #*-----^END Function load-EMSLatest ^-----