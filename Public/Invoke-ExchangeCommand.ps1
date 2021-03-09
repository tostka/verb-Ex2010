#*------v Invoke-ExchangeCommand.ps1 v------
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
    * 2:40 PM 9/17/2020 cleanup <?> encode damage (emdash's for dashes)
    * 4:28 PM 9/15/2020 cleanedup, added CBH, added to verb-Ex2010
    * 10/26/2015 posted vers
    .DESCRIPTION
    Invoke-ExchangeCommand.ps1 - PowerShell function allows you to run PowerShell commands and script blocks on Exchange servers in different forests without a forest trust.
    This PowerShell function allows you to run PowerShell commands and script blocks on Exchange servers in different forests without a forest trust.
    To run Invoke-ExchangeCommand, you must connect to the Exchange server using a hostname and not an IP address. Invoke-ExchangeCommand works best on Server 2012 R2/Windows 8.1 and later but also works on Server 2008 R2/Windows 7. Tested on Exchange 2010 and later. More information on cross-forest Exchange PowerShell can be found here: http://markgossa.blogspot.com/2015/10/exchange-2010-2013-cross-forest-remote-powershell.html
    Usage:
    1. Enable connections to all PowerShell hosts:
    winrm s winrm/config/client '@{TrustedHosts="*"}'
    # TSK: OR BETTER: _SELECTIVE_ HOSTS:
    Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value 'LYNMS7330.global.ad.toro.com' -Concatenate -Force ; Get-Item -Path WSMan:\localhost\Client\TrustedHosts | fl Name, Value ;
    cd WSMan:\localhost\Client ;
    dir | format-table -auto ; # review existing settings:
    # AllowEncrypted is defined on the client end, via the WSMAN: drive
    set-item .\allowunencrypted $true ;
    # You probably will need to set the AllowUnencrypted config setting in the Service as well, which has to be changed in the remote server using the following:
    set-item -force WSMan:\localhost\Service\AllowUnencrypted $true ;
    # tsk: reverted it back out:
    #-=-=-=-=-=-=-=-=
    [PS] WSMan:\localhost\Service\Auth> set-item -force WSMan:\localhost\Service\AllowUnencrypted $false ;
    cd ..
    [PS] WSMan:\localhost\Service>dir | format-table -auto ;
       WSManConfig: Microsoft.WSMan.Management\WSMan::localhost\Service
    Type          Name                             SourceOfValue Value
    ----          ----                             ------------- -----
    System.String RootSDDL                                       O:NSG:BAD:P(A;;GA;;;BA)(A;;GR;;;IU)S:P(AU;FA;GA;;;WD)(A...
    System.String MaxConcurrentOperations                        4294967295
    System.String MaxConcurrentOperationsPerUser                 1500
    System.String EnumerationTimeoutms                           240000
    System.String MaxConnections                                 300
    System.String MaxPacketRetrievalTimeSeconds                  120
    System.String AllowUnencrypted                               false
    Container     Auth
    Container     DefaultPorts
    System.String IPv4Filter                                     *
    System.String IPv6Filter                                     *
    System.String EnableCompatibilityHttpListener                false
    System.String EnableCompatibilityHttpsListener               false
    System.String CertificateThumbprint
    System.String AllowRemoteAccess                              true
    #-=-=-=-=-=-=-=-=
    TSK: try it *without* AllowUnencrypted before opening it up
    # And don't forget to also enable Digest Authorization:
    set-item -force WSMan:\localhost\Service\Auth\Digest $true ;
    # (to allow the system to digest the new settings)
    TSK: I don't even see the path existing on the lab Ex651
    WSMan:\localhost\Service\Auth\Digest
    TSK: but winrm shows the config enabled with Digest:
    winrm get winrm/config/client
    #-=-=-=-=-=-=-=-=
    Client
      NetworkDelayms = 5000
      URLPrefix = wsman
      AllowUnencrypted = true
      Auth
          Basic = true
          Digest = true
          Kerberos = true
          Negotiate = true
          Certificate = true
          CredSSP = false
      DefaultPorts
          HTTP = 5985
          HTTPS = 5986
      TrustedHosts = LYNMS7330
    #-=-=-=-=-=-=-=-=
    #-=-=L650'S settings-=-=-=-=-=-=
    # SERVICE AUTH
    [PS] C:\scripts>winrm get winrm/config/service/auth
    Auth
        Basic = false
        Kerberos = true
        Negotiate = true
        Certificate = false
        CredSSP = false
        CbtHardeningLevel = Relaxed
    # SERVICE OVERALL
    [PS] C:\scripts>winrm get winrm/config/service
    Service
    RootSDDL = O:NSG:BAD:P(A;;GA;;;BA)S:P(AU;FA;GA;;;WD)(AU;SA;GWGX;;;WD)
    MaxConcurrentOperations = 4294967295
    MaxConcurrentOperationsPerUser = 15
    EnumerationTimeoutms = 60000
    MaxConnections = 25
    MaxPacketRetrievalTimeSeconds = 120
    AllowUnencrypted = false
    Auth
        Basic = false
        Kerberos = true
        Negotiate = true
        Certificate = false
        CredSSP = false
        CbtHardeningLevel = Relaxed
    DefaultPorts
        HTTP = 5985
        HTTPS = 5986
    IPv4Filter = *
    IPv6Filter = *
    EnableCompatibilityHttpListener = false
    EnableCompatibilityHttpsListener = false
    CertificateThumbprint
    #-=-=-=-=-=-=-=-=
    ==3:22 PM 9/17/2020:POST settings on CurlyHoward:
    #-=-=-=-=-=-=-=-=
    [PS] WSMan:\localhost\Client>winrm get winrm/config/client
    Client
        NetworkDelayms = 5000
        URLPrefix = wsman
        AllowUnencrypted = true
        Auth
            Basic = true
            Digest = true
            Kerberos = true
            Negotiate = true
            Certificate = true
            CredSSP = false
        DefaultPorts
            HTTP = 5985
            HTTPS = 5986
        TrustedHosts = LYNMS7330
    [PS] WSMan:\localhost\Client>winrm get winrm/config/client
    Client
        NetworkDelayms = 5000
        URLPrefix = wsman
        AllowUnencrypted = true
        Auth
            Basic = true
            Digest = true
            Kerberos = true
            Negotiate = true
            Certificate = true
            CredSSP = false
        DefaultPorts
            HTTP = 5985
            HTTPS = 5986
        TrustedHosts = LYNMS7330
    [PS] WSMan:\localhost\Client>winrm get winrm/config/service/auth
    Auth
        Basic = false
        Kerberos = true
        Negotiate = true
        Certificate = false
        CredSSP = false
        CbtHardeningLevel = Relaxed
    [PS] WSMan:\localhost\Client>winrm get winrm/config/service
    Service
        RootSDDL = O:NSG:BAD:P(A;;GA;;;BA)(A;;GR;;;IU)S:P(AU;FA;GA;;;WD)(AU;SA;GXGW;;;WD)
        MaxConcurrentOperations = 4294967295
        MaxConcurrentOperationsPerUser = 1500
        EnumerationTimeoutms = 240000
        MaxConnections = 300
        MaxPacketRetrievalTimeSeconds = 120
        AllowUnencrypted = true
        Auth
            Basic = false
            Kerberos = true
            Negotiate = true
            Certificate = false
            CredSSP = false
            CbtHardeningLevel = Relaxed
        DefaultPorts
            HTTP = 5985
            HTTPS = 5986
        IPv4Filter = *
        IPv6Filter = *
        EnableCompatibilityHttpListener = false
        EnableCompatibilityHttpsListener = false
        CertificateThumbprint
        AllowRemoteAccess = true
    #-=-=-=-=-=-=-=-=
    #-=-ABOVE SETTINGS VIA WSMAN: PSDRIVE=-=-=-=-=-=-=
    [PS] WSMan:\localhost\Client>cd WSMan:\localhost\Client ;
    [PS] WSMan:\localhost\Client>dir | format-table -auto ;
       WSManConfig: Microsoft.WSMan.Management\WSMan::localhost\Client
    Type          Name             SourceOfValue Value
    ----          ----             ------------- -----
    System.String NetworkDelayms                 5000
    System.String URLPrefix                      wsman
    System.String AllowUnencrypted               true
    Container     Auth
    Container     DefaultPorts
    System.String TrustedHosts                   LYNMS7330

    [PS] WSMan:\localhost\Client>cd WSMan:\localhost\Service
    [PS] WSMan:\localhost\Service>dir | format-table -auto ;
       WSManConfig: Microsoft.WSMan.Management\WSMan::localhost\Service
    Type          Name                             SourceOfValue Value
    ----          ----                             ------------- -----
    System.String RootSDDL                                       O:NSG:BAD:P(A;;GA;;;BA)(A;;GR;;;IU)S:P(AU;FA;GA;;;WD)(A...
    System.String MaxConcurrentOperations                        4294967295
    System.String MaxConcurrentOperationsPerUser                 1500
    System.String EnumerationTimeoutms                           240000
    System.String MaxConnections                                 300
    System.String MaxPacketRetrievalTimeSeconds                  120
    System.String AllowUnencrypted                               true
    Container     Auth
    Container     DefaultPorts
    System.String IPv4Filter                                     *
    System.String IPv6Filter                                     *
    System.String EnableCompatibilityHttpListener                false
    System.String EnableCompatibilityHttpsListener               false
    System.String CertificateThumbprint
    System.String AllowRemoteAccess                              true

    [PS] WSMan:\localhost\Service>cd WSMan:\localhost\Service\Auth\
    [PS] WSMan:\localhost\Service\Auth>dir | format-table -auto ;
       WSManConfig: Microsoft.WSMan.Management\WSMan::localhost\Service\Auth
    Type          Name              SourceOfValue Value
    ----          ----              ------------- -----
    System.String Basic                           false
    System.String Kerberos                        true
    System.String Negotiate                       true
    System.String Certificate                     false
    System.String CredSSP                         false
    System.String CbtHardeningLevel               Relaxed
    #-=-=-=-=-=-=-=-=
    # ^ clearly digest doesn't even exist in the list on the service\auth
    
    Need to set to permit Basic Auth too?
    cd .\Auth ;
    Set-Item Basic $True ;
    Check if the user you're connecting with has proper authorizations on the remote machine (triggers GUI after the confirm prompt; use -force to suppress).
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
}

#*------^ Invoke-ExchangeCommand.ps1 ^------