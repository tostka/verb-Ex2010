# stop-Ex16ServicesInDependencyOrder.ps1


#region STOP_EX16SERVICESINDEPENDENCYORDER ; #*------v stop-Ex16ServicesInDependencyOrder v------
function stop-Ex16ServicesInDependencyOrder {
        <#
        .SYNOPSIS
        stop-Ex16ServicesInDependencyOrder.ps1 - Stop Exchange 2016 server, without any running Exchange servers of the revision online. Runs from native powershell, without any Remote EMS dependancy, also ensures all services - Automatic or Manual - startup settings are back to stock settings (restoring function after a service disable maintenance down).
        .NOTES
        Version     : 0.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2025-03-19
        FileName    : stop-Ex16ServicesInDependencyOrder.ps1
        License     : (none asserted)
        Copyright   : (none asserted)
        Github      : https://github.com/tostka/powershell/
        Tags        : Powershell,Exchange,ExchangeServer,Service,Maintenance
        AddedCredit : Ali Tajran
        AddedWebsite: https://www.alitajran.com/restart-exchange-services-powershell-script/
        AddedTwitter: URL
        REVISIONS
        * 1:21 PM 8/12/2025 fix: start-service has no -ComputerName $server param; fix validate typo in -targetrole
        * 3:44 PM 7/15/2025 had never actually finished conv fr start-Ex16ServicesInDep, now done;  ported to func, added to xopBuildLibrary.ps1
        * 3:15 PM 3/27/2025 added -nostart to restore startup settings to default, but not start the services (followup with a reboot). added -identity param, and pipeline looping, along with -ComputerName spec on start-service & set-service & get-service commands, to permit remote batch runs against array of servers
            added CBH; added svc installed pretest from MS Ex16 setup servicecontrol.ps1; updated to run remotely
        * April 22, 2021 Ali Tajran's posted version
        .DESCRIPTION
        stop-Ex16ServicesInDependencyOrder.ps1 - Stop Exchange 2016 server, without any running Exchange servers of the revision online. Runs from native powershell, without any Remote EMS dependancy, also ensures all services - Automatic or Manual - startup settings are back to stock settings (restoring function after a service disable maintenance down).

        Runs a list of automatic services, and a list of manual services: 
            sets the autos to startuptype:Disabled, 
            sets the manuals to startuptype:Disabled; 
            then runs both and stops the automatic startup services (the manuals should follow along automatically)

        Has no dependancy on running ex to open an EMS session into for native EMS cmdlets

        This script runs in stock powershell, no dependancy on EMS or it's cmdlets.
    

        Tweaked variant of: 
        [Restart Exchange services with PowerShell script - ALI TAJRAN](https://www.alitajran.com/restart-exchange-services-powershell-script/)
        Updated on April 22, 2021


        .PARAMETER Identity ; 
        Specify the identity of the Exchange Server. This can be piped from Get-ExchangeServer or specified explicitly using a string (defaults to local computer)
        .PARAMETER noStop
        Switch to suppress Stop of disabled servicese (e.g. manually perform a follow-up reboot, to let the entire system come up refreshed) [-noStop]
        .PARAMETER noDisable
        Switch to suppress Startup:Disabled of services (e.g. simply stop the server; a follow-up reboot will bring the system back up refreshed) [-noDisable]          
        .PARAMETER TargetRole
        Role specification (Mailbox|EdgeTransport) that overrides automatic role detection (which is normally predicated on presense of MSExchangeADTopology|MSExchangeEdgeCredential services) to suppress autodetect of target role [-TargetRole EdgeTransport]
        .INPUTS
        None. Does not accepted piped input.(.NET types, can add description)
        .OUTPUTS
        None. Returns no objects or output (.NET types)
        System.Boolean
        .EXAMPLE
        PS> .\stop-Ex16ServicesInDependencyOrder.ps1 -whatif -verbose
        .EXAMPLE
        PS> .\stop-Ex16ServicesInDependencyOrder.ps1 -identity SERVER1 -nostop; 
        Run against a specific machine, specifying -nostop to down default service startup settings, wo performing a full service disable (manual reboot would be the normal follow-on). 
        .EXAMPLE
        PS> 'server1','server2' | get-exchangeserver | select -expand fqdn | .\stop-Ex16ServicesInDependencyOrder.ps1
        Run against a series; pipeline usage, routing through get-exchangeserver to ensure are actual mail servers, and expanding into specific full fqdns
        .LINK
        https://github.com/tostka/verb-Ex2010
        .LINK
        https://github.com/tostka/powershellBB/
        #>
        <# Stops the minimum key services, to cleanly bring down an Exchange server: MSExchangeTransport,MSExchangeIS,MSExchangeMailboxAssistants,MSExchangeMailboxReplication
        Below runs in stock powershell, no dependancy
        [Restart Exchange services with PowerShell script - ALI TAJRAN](https://www.alitajran.com/restart-exchange-services-powershell-script/)
        Updated on April 22, 2021

        Citations of processes to do the above: 


        [Exchange Server 2016 CU Setup Cannot Stop Service due to Access Denied – Granikos GmbH & Co. KG](https://granikos.eu/exchange-server-2016-cu-setup-cannot-stop-service-due-to-access-denied/)

        > The PowerShell code executed as part of the CU Setup sets the startup type of
        > Exchange and some Windows services to **Disabled**. This ensures that in the
        > case of a server reboot, an automatic service start will not interfere with a
        > partially executed setup. After setting the startup type to Disabled, the
        > services are stopped.  
        > 
        > The services are controlled by the ServiceControl.ps1 script is located on the Exchange Server installation media in \Setup\ServerRoles\Common.
        > 
        Locating the above: (corrrected un-delimited path in comment above)
        gci -path 'D:\cab\ExchangeServer2016-x64-CU23-ISO\unpacked\ServiceControl.ps1'  -recur
            Directory: D:\cab\ExchangeServer2016-x64-CU23-ISO\unpacked\Setup\ServerRoles\Common
        Mode                LastWriteTime         Length Name
        ----                -------------         ------ ----
        -ar---        3/26/2022   1:27 PM          53801 ServiceControl.ps1


        > The function StopServices stops services using the Stop-Service cmdlet. Due to timing issues, some services are _stopped_ by killing the running processes using Stop-Process -Force.
        > 
        > The services stopped by stopping the running process are:
        > 
        > -   FMS
        > -   MSExchangeServiceHost
        > -   MSExchangeTransport
        > -   MSExchangeInferenceService
        > -   MSExchangeDagMgmt

        Lifting the mailbox role spec from the .ps1:
        $script:servicesToControl['Mailbox']            = @( 'MSExchangeMonitoring', 'IISAdmin', 'MSExchangeIS', 'MSExchangeMailboxAssistants', 'MSFTESQL-Exchange', 'MSExchangeThrottling', 'MSExchangeADTopology' ,'MSExchangeTopologyService', 'MSExchangeRepl', 'MSExchangeDagMgmt', 'MSExchangeWatchDog', 'MSExchangeTransportLogSearch', 'MSExchangeRPC', 'MSExchangeServiceHost', 'W3Svc', 'HTTPFilter', 'wsbexchange', 'MSExchangeTransportSyncManagerSvc', 'MSExchangeFastSearch', 'hostcontrollerservice', 'SearchExchangeTracing', 'MSExchangeSubmission', 'MSExchangeDelivery', 'MSExchangeMigrationWorkflow', 'MSExchangeDiagnostics', 'MSExchangeProcessUtilizationManager', 'MSExchangeHM', 'MSExchangeHMRecovery', 'MSExchangeInferenceService')

        ["exchange 2016 server" stop services powershell reboot - Google Search](https://www.google.com/search?q=%22exchange+2016+server%22+stop+services+powershell+reboot&num=10&client=firefox-b-1-d&sca_esv=5e95eb279cb35b6f&ei=NjzjZ4uoOs610PEPn5_zmAo&ved=0ahUKEwiLm7assaaMAxXOGjQIHZ_PHKMQ4dUDCBA&uact=5&oq=%22exchange+2016+server%22+stop+services+powershell+reboot&gs_lp=Egxnd3Mtd2l6LXNlcnAiNiJleGNoYW5nZSAyMDE2IHNlcnZlciIgc3RvcCBzZXJ2aWNlcyBwb3dlcnNoZWxsIHJlYm9vdDIFEAAY7wUyBRAAGO8FMgUQABjvBTIIEAAYogQYiQUyBRAAGO8FSKshUJUMWPkUcAF4AJABAJgBiwGgAcACqgEDMS4yuAEDyAEA-AEBmAIDoAL0AcICChAAGLADGNYEGEfCAggQIRigARjDBJgDAIgGAZAGCJIHAzIuMaAHkwyyBwMxLjG4B-YB&sclient=gws-wiz-serp)

        The below takese transport & CAS out, but doesn't actually down dbs etc...

        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        > AI Overview
        > Learn more
        > To stop Exchange 2016 services using PowerShell and then reboot the server, first open the Exchange Management Shell as an administrator, then use the Stop-Service cmdlet to stop the services, and finally reboot the server using Shutdown /r /t 0. 
        > Here's a more detailed breakdown:
        > 1. Open Exchange Management Shell:
        > 
        >     Click the Start button, type "Exchange Management Shell", and right-click the result to select "Run as administrator".
        > 
        > 2. Stop Exchange Services:
        > 
        >     Use the Stop-Service cmdlet to stop the Exchange services you need. For example:
        > 
        > Code
        > 
        >        Stop-Service -Name MSExchangeMailboxReplication
        >        Stop-Service -Name MSExchangeTransport
        >        Stop-Service -Name MSExchangeRpcClientService
        > 
        >     You can use Get-Service to view the services and their names to ensure you are stopping the correct services.
        > 
        > 3. Reboot the Server:
        > 
        >     Use the following command to initiate a reboot:
        > 
        > Code
        > 
        >        Shutdown /r /t 0
        > 
        >     This command shuts down the server and restarts it immediately.
        > 
        > Important Considerations:
        > 
        >     Backup:
        >     Before making any significant changes, ensure you have a recent backup of your Exchange Server.
        > 
        > Order of Operations:
        > Stop services in a logical order, ensuring that dependencies are handled correctly. For example, stop the transport service before the mailbox replication service. 
        > Troubleshooting:
        > If you encounter issues, check the Exchange logs for errors and consult Microsoft documentation for troubleshooting steps. 
        > Restart IIS:
        > After stopping and starting Exchange services, you may need to restart IIS to ensure that the changes are reflected. You can do this by using the command iisreset in the command prompt. 
        > 
        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        TSK: The below MSExchangeFBA doesn't even exist as an Exchange 2016 service!

        [Exchange command or script to disable all Exchange services - Collaboration - Spiceworks Community](https://community.spiceworks.com/t/exchange-command-or-script-to-disable-all-exchange-services/790293/2)

        > obb4
        > Jalapeno
        > Feb 2021
        > A script containing these commands would stop all the services.
        > ```cmd
        > net stop MSExchangeADTopology /y
        > net stop MSExchangeFBA
        > net stop wsbexchange
        > net stop MSExchangeMonitoring
        > net stop MSExchangeIS
        > net stop MSExchangeSA
        > ```
        > note: the above is C:\Windows\system32\net.exe, works in ps, no conflicts
        > 
        > To disable them:
        > ```cmd
        > sc config "Name of Service" start= disabled
        > ```
        > note: the above is C:\Windows\system32\sc.exe, not the powershell alias sc => set-content
        > likely safer, in powershell to run as: sc.exe config "Name of Service" start= disabled
        > 
        > I don't think any IT service firm would prank a customer.
        > Have they installed any kind of RMM system? Perhaps they inadvertently copied a config to your DAG.
        > 
        #>
        [CmdletBinding()]
        [Alias('enable-xopExServer')]
        PARAM (
            [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$true,
                HelpMessage="Specify the identity of the Exchange Server. This can be piped from Get-ExchangeServer or specified explicitly using a string")]
                [ValidateNotNullOrEmpty()]
                [PSCustomObject]$Identity=$env:computername,
            [Parameter(HelpMessage="Switch to suppress Stop of disabled servicese (e.g. manually perform a follow-up reboot, to let the entire system come up refreshed) [-noStop]")]
                [switch]$noStop,
            [Parameter(HelpMessage="Switch to suppress Startup:Disabled of services (e.g. simply stop the server; a follow-up reboot will bring the system back up refreshed) [-noDisable]")]
                [switch]$noDisable,
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

                # Disable Services
                if($noDisable){
                    $SetStartType = 'Automatic' ; 
                }else{
                    $SetStartType = 'Disabled' ; 
                } ; 
                foreach ($service in $auto) {
                    if(Get-Service -Name "$service*" -ComputerName $server -ea 0  | ?{$_.Name -eq $service}){
                        Set-Service -Name $service -StartupType $SetStartType -ComputerName $server -ea 0 -verbose:$true ; 
                        Write-Host "Enabling "$service
                    }else{
                        $smsg = "$service is not installed on $($server)" ; 
                        if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                            if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        } ;
                    } ; 
                } # loop-E auto
                # Disable Services
                if($noDisable){
                    $SetStartType = 'Manual' ; 
                }else{
                    $SetStartType = 'Disabled' ; 
                } ; 
                foreach ($service in $man) {
                    if(Get-Service "$service*" -ComputerName $server -ea 0  | ?{$_.Name -eq $service}){
                        Set-Service -Name $service2 -StartupType $SetStartType -ComputerName $server -ea 0  -verbose:$true ; 
                        Write-Host "Enabling "$service2
                    }else{
                        $smsg = "$service is not installed on $($server)" ; 
                        if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                            if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        } ;
                    } ; 
                } # loop-E man

                if(-not $noStop){
                    # Stop Services
                    foreach ($service in $auto) {
                        if(Get-Service "$service*"  -ComputerName $server -ea 0 | ?{$_.Name -eq $service}){
                            Write-Host "Starting "$service 
                            #Start-Service -Name $service  -ComputerName $server -ea 0
                            Start-Service -Name $service -ea 0 # has no -computer param!
                        }else{
                            $smsg = "$service is not installed on $($server)" ; 
                            if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                                if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            } ;
                        } ; 
                    } ; 
                }else {
                    $smsg = "Skipping stop of services: Perform a reboot of the sytem to perform a full refresh`n(or rerun this script without -NoStop parameter" ; 
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } ;

                $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))`n`n" ;
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            } ;  # loop-E Identity
        } ;  # PROC-E
        END{

        } ; 
    }
#endregion STOP_EX16SERVICESINDEPENDENCYORDER ; #*------^ END stop-Ex16ServicesInDependencyOrder ^------

