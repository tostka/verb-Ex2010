#*------v Function get-DCLocal v------
Function get-DCLocal {
    <#
    .SYNOPSIS
    get-DCLocal - Function to locate a random DC in the local AD site (sub-250ms response)
    .NOTES
    Author: Todd Kadrie
    Website:	http://tinstoys.blogspot.com
    Twitter:	http://twitter.com/tostka

    Additional Credits: Originated in Ben Lye's GetLocalDC()
    Website:	http://www.onesimplescript.com/2012/03/using-powershell-to-find-local-domain.html
    REVISIONS   :
    12:32 PM 1/8/2015 - tweaked version of Ben lye's script, replaced broken .NET site query with get-addomaincontroller ADMT module command
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns one DC object, .Name is name pointer
    .EXAMPLE
    C:\> get-dclocal
    #>

    #  alt command: Return one unverified connectivityDC in SITE site:
    # Get-ADDomainController -discover -site "SITE"

    # Set $ErrorActionPreference to continue so we don't see errors for the connectivity test
    $ErrorActionPreference = 'SilentlyContinue'
    # Get all the local domain controllers
    # .Net call below fails in LYN, because LYN'S SITE LISTS NO SERVERS ATTRIBUTE!
    #$LocalDCs = ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()).Servers
    # use get-addomaincontroller to do it
    $Site = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name ;
    # gc filter
    #$LocalDCs = Get-ADDomainController -filter {(isglobalcatalog -eq $true) -AND (Site -eq $Site)} ;
    # any dc filter
    $LocalDCs = Get-ADDomainController -filter { (Site -eq $Site) } ;
    # Create an array for the potential DCs we could use
    $PotentialDCs = @()
    # Check connectivity to each DC
    ForEach ($LocalDC in $LocalDCs) {
        #write-verbose -verbose $localdc
        # Create a new TcpClient object
        $TCPClient = New-Object System.Net.Sockets.TCPClient
        # Try connecting to port 389 on the DC
        $Connect = $TCPClient.BeginConnect($LocalDC.Name, 389, $null, $null)
        # Wait 250ms for the connection
        $Wait = $Connect.AsyncWaitHandle.WaitOne(250, $False)
        # If the connection was succesful add this DC to the array and close the connection
        If ($TCPClient.Connected) {
            # Add the FQDN of the DC to the array
            $PotentialDCs += $LocalDC.Name
            # Close the TcpClient connection
            $Null = $TCPClient.Close()
        } # if-E
    } # loop-E
    # Pick a random DC from the list of potentials
    $DC = $PotentialDCs | Get-Random
    #write-verbose -verbose $DC
    # Return the DC
    Return $DC
} #*------^ END Function get-DCLocal ^------