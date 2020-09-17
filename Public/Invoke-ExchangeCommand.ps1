
#*------v Function Invoke-ExchangeCommand v------
function Invoke-ExchangeCommand{
    <#
    .SYNOPSIS
    Invoke-ExchangeCommand.ps1 - PowerShell function allows you to run PowerShell commands and script blocks on Exchange servers in different forests without a forest trust.
    .NOTES
    Version     : 0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2020-09-15
    FileName    : IInvoke-ExchangeCommand.ps1
    License     : MIT License
    Copyright   : (c) 2015 Mark Gossa
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell,Exchange,ExchangeForestMigration,CrossForest,ExchangeRemotePowerShell,ExchangePowerShell
    AddedCredit : Mark Gossa
    AddedWebsite: https://gallery.technet.microsoft.com/Exchange-Cross-Forest-e25d48eb
    AddedTwitter:
    REVISIONS
    * 4:28 PM 9/15/2020 cleanedup, added CBH, added to verb-Ex2010
    * 10/26/2015 posted vers
    .DESCRIPTION
    Invoke-ExchangeCommand.ps1 - PowerShell function allows you to run PowerShell commands and script blocks on Exchange servers in different forests without a forest trust.
    This PowerShell function allows you to run PowerShell commands and script blocks on Exchange servers in different forests without a forest trust.
    To run Invoke-ExchangeCommand, you must connect to the Exchange server using a hostname and not an IP address. Invoke-ExchangeCommand works best on Server 2012 R2/Windows 8.1 and later but also works on Server 2008 R2/Windows 7. Tested on Exchange 2010 and later. More information on cross-forest Exchange PowerShell can be found here: http://markgossa.blogspot.com/2015/10/exchange-2010-2013-cross-forest-remote-powershell.html
    Usage:
    1. Enable connections to all PowerShell hosts:
    winrm s winrm/config/client '@{TrustedHosts="*"}'
    # or better: selective hosts:
    Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value 'client.domain.com' -Concatenate –Force ; Get-Item -Path WSMan:\localhost\Client\TrustedHosts | fl Name, Value ;
    cd WSMan:\localhost\Client ;
    dir | format-table –auto ; # review existing settings:
    # AllowEncrypted is defined on the client end, via the WSMAN: drive
    set-item .\allowunencrypted $true ;
    # You probably will need to set the AllowUnencrypted config setting in the Service as well, which has to be changed in the remote server using the following:
    set-item -force WSMan:\localhost\Service\AllowUnencrypted $true ;
    # And don't forget to also enable Digest Authorization:
    set-item -force WSMan:\localhost\Service\Auth\Digest $true ;
    # (to allow the system to digest the new settings)
    TSK: I don't even see the path existing on the lab Ex651
    WSMan:\localhost\Service\Auth\Digest
    Need to set to permit Basic Auth too?
    cd .\Auth ;
    Set-Item Basic $True ;
    Check if the user you're connecting with has proper authorizations on the remote machine (triggers GUI after the confirm prompt; use –force to suppress).
    Set-PSSessionConfiguration -ShowSecurityDescriptorUI -Name Microsoft.PowerShell ;
    .PARAMETER  ExchangeServer
    Target Exchange Server[-ExchangeServer server.domain.com]
    .PARAMETER  Scriptblock
    Scriptblock/Command to be executed on target server[-ScriptBlock {Get-Mailbox | ft}]
    .PARAMETER  $Credential
    Credential object to be used for connection[-Credential cred]
    .PARAMETER Whatif
    Parameter to run a Test no-change pass [-Whatif switch]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns objects returned to pipeline
    .EXAMPLE
    .\Invoke-ExchangeCommand.ps1
    .EXAMPLE
    .\Invoke-ExchangeCommand.ps1
    .LINK
    https://github.com/tostka/verb-Ex2010
    .LINK
    https://gallery.technet.microsoft.com/Exchange-Cross-Forest-e25d48eb
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, HelpMessage = "Target Exchange Server[-ExchangeServer server.domain.com]")]
        [string] $ExchangeServer,
        [Parameter(Mandatory = $true, HelpMessage = "Scriptblock/Command to be executed[-ScriptBlock {Get-Mailbox | ft}]")]
        [string] $ScriptBlock,
        [Parameter(Mandatory = $true, HelpMessage = "Credentials [-Credential credobj]")]
        [System.Management.Automation.PSCredential] $Credential
    ) ;
    BEGIN {
        #${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        # Get parameters this function was invoked with
        #$PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        # silently stop any running transcripts
        $stopResults = try { Stop-transcript -ErrorAction stop } catch {} ;
        $WarningPreference = "SilentlyContinue" ;
    } ; # BEGIN-E
    PROCESS {
        $Error.Clear() ;
        #Connect to DC and pass through credential variable
        $pltICPS = @{
            ComputerName   = $ExchangeServer ;
            ArgumentList   = $Credential, $ExchangeServer, $ScriptBlock, $WarningPreference ;
            Credential     = $Credential ;
            Authentication = 'Negotiate'
        } ;
        write-verbose "Invoke-Command  w`n$(($pltICPS|out-string).trim())`n`$ScriptBlock:`n$(($ScriptBlock|out-string).trim())" ;
        #Invoke-Command -ComputerName $ExchangeServer -ArgumentList $Credential,$ExchangeServer,$ScriptBlock,$WarningPreference -Credential $Credential -Authentication Negotiate
        Invoke-Command @pltICPS -ScriptBlock {

                #Specify parameters
                param($Credential,$ExchangeServer,$ScriptBlock,$WarningPreference)

                #Create new PS Session
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ `
                -Authentication Kerberos -Credential $Credential

                #Import PS Session
                Import-PSSession $Session | Out-Null

                #Run commands
                foreach($Script in $ScriptBlock){
                    Invoke-Expression $Script
                }

                #Close all open sessions
                Get-PSSession | Remove-PSSession -Confirm:$false
            }
    } ; # PROC-E
    END {    } ; # END-E
} #*------^ END Function Invoke-ExchangeCommand ^------


