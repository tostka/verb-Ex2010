# start-Ex16ServicesInDependencyOrder.ps1


#region START_EX16SERVICESINDEPENDENCYORDER ; #*------v start-Ex16ServicesInDependencyOrder v------
function start-Ex16ServicesInDependencyOrder{
        <#
        .SYNOPSIS
        start-Ex16ServicesInDependencyOrder.ps1 - Cold start Exchange 2016 server, without any running Exchange servers of the revision online. Runs from native powershell, without any Remote EMS dependancy, also ensures all services - Automatic or Manual - startup settings are back to stock settings (restoring function after a service disable maintenance down).
        .NOTES
        Version     : 0.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2025-03-19
        FileName    : start-Ex16ServicesInDependencyOrder.ps1
        License     : (none asserted)
        Copyright   : (none asserted)
        Github      : https://github.com/tostka/powershell/
        Tags        : Powershell,Exchange,ExchangeServer,Service,Maintenance
        AddedCredit : Ali Tajran
        AddedWebsite: https://www.alitajran.com/restart-exchange-services-powershell-script/
        AddedTwitter: URL
        REVISIONS
        * 1:21 PM 8/12/2025 fix: start-service has no -ComputerName $server param; fixed typo in valid on the -targetrole; 
        * 10:45 AM 8/6/2025 added write-myOutput|Warning|Verbose support (for xopBuildLibrary/install-Exchange15.ps1 compat) 
        * 1:33 PM 7/15/2025 ported to func, added to xopBuildLibrary.ps1
        * 3:15 PM 3/27/2025 added -nostart to restore startup settings to default, but not start the services (followup with a reboot). added -identity param, and pipeline looping, along with -ComputerName spec on start-service & set-service & get-service commands, to permit remote batch runs against array of servers
            added CBH; added svc installed pretest from MS Ex16 setup servicecontrol.ps1; updated to run remotely
        * April 22, 2021 Ali Tajran's posted version
        .DESCRIPTION
        start-Ex16ServicesInDependencyOrder.ps1 - Cold start Exchange 2016 server, without any running Exchange servers of the revision online. Runs from native powershell, without any Remote EMS dependancy, also ensures all services - Automatic or Manual - startup settings are back to stock settings (restoring function after a service disable maintenance down)

        Runs a list of automatic services, and a list of manual services: 
            sets the autos to startuptype:automatic, 
            sets the manuals properly to startuptype:manual; 
            then runs both and starts the automatic startup services (the manuals should follow along automatically)

        Has advantage over the old...

        ```powershell
        Test-ServiceHealth |  Select-Object ServicesNotRunning –Expand ServicesNotrunning | Start-Service ;
        ```

        ... in that you can't use tsh, if there's no running ex to open an EMS session into. 

        This script runs in stock powershell, no dependancy on EMS or it's cmdlets.

        This also functions to reverse disable-xopExServer() (and is a configured alias on this function)

        Tweaked variant of: 
        [Restart Exchange services with PowerShell script - ALI TAJRAN](https://www.alitajran.com/restart-exchange-services-powershell-script/)
        Updated on April 22, 2021


        .PARAMETER Identity ; 
        Specify the identity of the Exchange Server. This can be piped from Get-ExchangeServer or specified explicitly using a string (defaults to local computer)
        .PARAMETER nostart
        Optional switch to suppress start of servicese (e.g. solely restore normal Startup setting, and *manually* perform a follow-up reboot, to let the entire system come up refreshed) [-noStart]
        .PARAMETER TargetRole
        Role specification (Mailbox|EdgeTransport) that overrides automatic role detection (which is normally predicated on presense of MSExchangeADTopology|MSExchangeEdgeCredential services) to suppress autodetect of target role [-TargetRole EdgeTransport]
        .INPUTS
        None. Does not accepted piped input.(.NET types, can add description)
        .OUTPUTS
        None. Returns no objects or output (.NET types)
        System.Boolean
        .EXAMPLE
        PS> .\start-Ex16ServicesInDependencyOrder.ps1 -whatif -verbose
        .EXAMPLE
        PS> .\start-Ex16ServicesInDependencyOrder.ps1 -identity SERVER1 -nostart; 
        Run against a specific machine, specifying -nostart to restore default service startup settings, wo performing a full service startup (manual reboot would be the normal follow-on). 
        .EXAMPLE
        PS> 'server1','server2' | get-exchangeserver | select -expand fqdn | .\start-Ex16ServicesInDependencyOrder.ps1
        Run against a series; pipeline usage, routing through get-exchangeserver to ensure are actual mail servers, and expanding into specific full fqdns
        .LINK
        https://github.com/tostka/verb-Ex2010
        .LINK
        https://github.com/tostka/powershellBB/
        #>
        [CmdletBinding()]
        [Alias('enable-xopExServer')]
        PARAM (
            [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$true,
                HelpMessage="Specify the identity of the Exchange Server. This can be piped from Get-ExchangeServer or specified explicitly using a string")]
                [ValidateNotNullOrEmpty()]
                [PSCustomObject]$Identity=$env:computername,
            [Parameter(HelpMessage="Switch to suppress start of re-enabled servicese (e.g. manually perform a follow-up reboot, to let the entire system come up refreshed) [-noStart]")]
                [switch]$noStart,
            [Parameter(HelpMessage="Role specification (Mailbox|EdgeTransport) that overrides automatic role detection (which is normally predicated on presense of MSExchangeADTopology|MSExchangeEdgeCredential services) to suppress autodetect of target role [-TargetRole EdgeTransport]")]
                [ValidateSet('Mailbox','EdgeTransport')]
                [string]$TargetRole
        ) ; 
        BEGIN{
            #region LOCAL_CONSTANTS ; #*------v LOCAL_CONSTANTS v------

            ## test we successfully connected to v15/Ex16:
            #$ExVersNo = '15.1' ; 

            if(-not $TargetRole){
                # determine target role, based on services: get-service -ComputerName lynms6400 -Name MSExchangeADTopology # non-Edge
                # MSExchangeEdgeCredential # edge
                if(get-service -ComputerName $Identity -name MSExchangeEdgeCredential -ea 0){
                    $TargetRole = 'Mailbox'
                } elseif(get-service -ComputerName $Identity -name MSExchangeADTopology -ea 0){ 
                    $TargetRole = 'EdgeTransport'
                } else {
                    $smsg = "UNABLE TO RESOLVE A SPECIFIC -TARGETROLE (UNABLE TO REMOTE LOCATE MSExchangeADTopology|MSExchangeEdgeCredential SERVICE TO STEER LOGIC!" ; 
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 
                    THROW $smsg ; 
                    BREAK ; 
                }; 
                $smsg = "Using resolved -TargetRole: $($TargetRole)" ; 
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            } else{
                $smsg = "Using specified -TargetRole: $($TargetRole)" ; 
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            }; 

            switch ($TargetRole){
                'Mailbox' {
                    # Automatic services
                    $auto = "MSExchangeADTopology",
                        "MSExchangeAntispamUpdate",
                        "MSComplianceAudit",  # ADD: 1:54 PM 7/15/2025
                        "MSExchangeCompliance", # ADD: 1:54 PM 7/15/2025
                        "MSExchangeDagMgmt",
                        "MSExchangeDiagnostics",
                        "MSExchangeEdgeSync",
                        "MSExchangeFrontEndTransport",
                        "MSExchangeHM",
                        "MSExchangeHMRecovery", # ADD: 1:56 PM 7/15/2025
                        "MSExchangeImap4",
                        "MSExchangeIMAP4BE",
                        "MSExchangeIS",
                        "MSExchangeMailboxAssistants",
                        "MSExchangeMailboxReplication",
                        "MSExchangeDelivery",
                        "MSExchangeSubmission",
                        "MSExchangeNotificationsBroker", # ADD: 1:57 PM 7/15/2025
                        "MSExchangeRepl",
                        "MSExchangeRPC",
                        "MSExchangeFastSearch",
                        "HostControllerService",
                        "MSExchangeServiceHost",
                        "MSExchangeThrottling",
                        "MSExchangeTransport",
                        "MSExchangeTransportLogSearch",
                        "MSExchangeUM",
                        "MSExchangeUMCR",
                        "MSExchange Mitigation", # 1:55 PM 7/15/2025 ADD
                        "FMS",
                        "IISADMIN",
                        "RemoteRegistry",
                        "SearchExchangeTracing",
                        "Winmgmt",
                        "W3SVC"

                    # Manual services
                    $man = "MSExchangePop3",
                      "MSExchangePOP3BE",
                      "WSBExchange",
                      "AppIDSvc",
                      "pla"
                }
                'EdgeTransport'{
                    # Auto Exchange services on Edge Transport servers ADD: all 2:01 PM 7/15/2025
                    $auto = 'ADAM_MSExchange',
                        'MSExchangeAntispamUpdate',
                        'MSExchangeEdgeCredential',
                        'MSExchangeDiagnostics',
                        'MSExchangeHM',
                        'MSExchangeHMRecovery',
                        'MSExchangeServiceHost',
                        'MSExchangeTransport',
                        'MSExchangeTransportLogSearch'
                } ; 
            };
            #endregion LOCAL_CONSTANTS ; #*------^ END LOCAL_CONSTANTS ^------ 
        } # BEG-E
        PROCESS{
    
            foreach($server in $Identity){
                $sBnrS="`n`n#*------v PROCESSING : $($server.toupper()) v------" ; 

                # Enable Services
                foreach ($service in $auto) {
                    if(Get-Service -Name "$service*" -ComputerName $server -ea 0  | ?{$_.Name -eq $service}){
                        Set-Service -Name $service -StartupType Automatic -ComputerName $server -ea 0 -verbose:$true ; 
                        Write-Host "Enabling "$service
                    }else{
                        $smsg = "$service is not installed on $($server)" ; 
                        if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                            if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        } ;
                    } ; 
                }
                foreach ($service in $man) {
                    if(Get-Service "$service*" -ComputerName $server -ea 0  | ?{$_.Name -eq $service}){
                        Set-Service -Name $service2 -StartupType Manual -ComputerName $server -ea 0  -verbose:$true ; 
                        Write-Host "Enabling "$service2
                    }else{
                        $smsg = "$service is not installed on $($server)" ; 
                        if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                            if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        } ;
                    } ; 
                }

                if(-not $noStart){
                    # Start Services
                    foreach ($service in $auto) {
                        if(Get-Service "$service*"  -ComputerName $server -ea 0 | ?{$_.Name -eq $service}){
                            Write-Host "Starting "$service 
                            #Start-Service -Name $service  -ComputerName $server -ea 0
                            # has no -comp param
                            Start-Service -Name $service -ea 0                 
                        }else{
                            $smsg = "$service is not installed on $($server)" ; 
                            if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                            if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        } ;
                        } ; 
                    } ; 
                }else {
                    $smsg = "Skipping startup of services: Perform a reboot of the sytem to perform a full refresh startup`n(or rerun this script without -NoStart parameter" ; 
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } ; 

                $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))`n`n" ;
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            } ; 
        } ;  # PROC-E
        END{

        } ;
    }
#endregion START_EX16SERVICESINDEPENDENCYORDER ; #*------^ END start-Ex16ServicesInDependencyOrder ^------

