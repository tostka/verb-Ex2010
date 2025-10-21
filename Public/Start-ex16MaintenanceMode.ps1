# Start-ex16MaintenanceMode.ps1


#region START_EX16MAINTENANCEMODE ; #*------v Start-ex16MaintenanceMode v------
function Start-ex16MaintenanceMode{
        <#
        .SYNOPSIS
        Start-ex16MaintenanceMode.ps1 - Puts a Microsoft Exchange Server 2016 computer into maintenance mode.
        .NOTES
        Version     : 0.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2025-03-19
        FileName    : Start-ex16MaintenanceMode.ps1
        License     : MIT License
        Copyright   : (c) 2025 Todd Kadrie
        Github      : https://github.com/tostka/verb-XXX
        Tags        : Powershell
        AddedCredit : PietroCiaccio
        AddedWebsite: https://github.com/PietroCiaccio/
        AddedTwitter: URL
        REVISIONS
        * 1:58 PM 8/17/2025 updated svcrestart for Edge non-remote icm support
        * 5:47 PM 8/16/2025 Start-ex16MaintenanceMode: add Edge detection support
        * 9:06 AM 7/24/2025 add return of Get-ex16MaintenanceModeTDO results
        * 9:09 AM 3/26/2025 added ipsn post connect-ExchangeServer() call (wasn't including import/non-functional [headscratch]); beyond that, used wo issues on 1st of the Ex16 builds
        * 8/12/2020 Pietro Ciaccio's PSG-posted ExchangePowerShell module, v0.11.0
        .DESCRIPTION
        Puts a Microsoft Exchange Server 2016 computer into maintenance mode. CmdLet will - ;
                    - drain queues ;
                    - restart transport services ;
                    - redirect messages to a redirection server ;
                    - move off active database copies to an available DAG member ;
                    - suspend the cluster node ;
                    - prevent database activation on the server ;
                    - suspend passive copies ;
                    - set all server component states to inactive ;

        Based on PietroCiaccio's github-posted function from his ExchangePowerShell module
        https://github.com/PietroCiaccio/ExchangePowerShell
        ... and as posted to PSG:
        https://www.powershellgallery.com/packages/ExchangePowerShell/0.2.2/Content/ExchangePowerShell.psm1
        Appears the PSG is 0.11.0 (current version) 	575,293 	8/12/2020
        As is github: (from .psm1 header):

        ```text
        # EP (ExchangePowerShell) Powershell Module 0.11.0 - Oct 14, 2021,
        # Author: Pietro Ciaccio | LinkedIn: https://www.linkedin.com/in/pietrociaccio | Twitter: @PietroCiac
        # PSG copy: https://www.powershellgallery.com/packages/ExchangePowerShell/0.11.0 | 0.11.0 (current version) 	575,293 	8/12/2020
        # => the avail gh vers is as current as they come (below)
        ```

        issue: Ex16 Connect-ExchangeServer -auto -ClientApplication:ManagementShell"
        If it can't find the local - services down - it WILL DIVERT INTO older EXCH versions!
        WHICH DOESN'T SEE EX16 SERVERS IN GET-EXCHANGESERVER!
        So we need to version check the session we pull (and use the -serverFQDN param of connect-Exchangeserver()), to ensure we got the server we want.

            get-exchangeserver [server] returns Admindisplayversion like :
            Version nn.n (Build nnn.n)
            The Buildnumber Shortstring can be constructed by combining the 'Version nn.n' digits with the '(Build nnn.n)' digits: nn.n.nnn.n

                [regex]::Matches($AdminDisplayVersion,"(\d*\.\d*)").value -join '.'

            Do matching on the nn.n for each Major rev, if you want to be sure of your supported commandset targets
                - 2019, 'Version 15.2'
                - 2016, 'Version 15.1'
                - 2013, 'Version 15.0'
                - 2010sp3, 'Version 14.3'
                - 2010sp2, 'Version 14.2'
                - 2010sp1, 'Version 14.1'
                - 2010, 'Version 14.0'

        .PARAMETER Identity ;
        Specify the identity of the Exchange Server. This can be piped from Get-ExchangeServer or specified explicitly using a string.
        .PARAMETER RedirectionTarget ;
        Specify the identity of the Exchange Server you wish to redirect pending messages to.
        .PARAMETER MoveActiveDatabaseCopies ;
        Specify whether to move active database copies to other DAG members, if possible. The default is false.
        .INPUTS
        None. Does not accepted piped input.(.NET types, can add description)
        .OUTPUTS
        None. Returns no objects or output (.NET types)
        System.Boolean
        [| get-member the output to see what .NET obj TypeName is returned, to use here]
        .EXAMPLE
        PS> .\Start-ex16MaintenanceMode.ps1 -whatif -verbose
        EXSAMPLEOUTPUT
        Run with whatif & verbose
        .EXAMPLE
        PS> .\Start-ex16MaintenanceMode.ps1
        EXSAMPLEOUTPUT
        EXDESCRIPTION
        .LINK
        https://github.com/tostka/verb-XXX
        .LINK
        https://github.com/tostka/powershell/
        #>
        # #Requires -PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
        [CmdletBinding()]
        Param (
            [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$true,
                HelpMessage="Specify the identity of the Exchange Server. This can be piped from Get-ExchangeServer or specified explicitly using a string")]
                [ValidateNotNullOrEmpty()]
                [PSCustomObject]$Identity=$env:computername,
            [Parameter(valuefrompipelinebypropertyname=$false,HelpMessage="Specify the identity of the Exchange Server you wish to redirect pending messages to.")]
                [PSCustomObject]$RedirectionTarget,
            [Parameter(HelpMessage="Specify whether to move active database copies to other DAG members, if possible. The default is false.")]
                [boolean]$MoveActiveDatabaseCopies = $false
        ) ;
        BEGIN{
            if(-not (gcm Set-ServerComponentState -ea 0)){
                if(connect-XopLocalManagementShell){ 
                    write-host -foregroundcolor green "Connected" 
                } else { 
                    $smsg = "NOT CONNECTED!"
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } ;
            } ; 
        } # BEG-E
        PROCESS {
            
            # Validate identity
            if ($input) {
                $ExchangeServer = $null; $ExchangeServer = $input | Get-ex16ExchangeServerTDO
            } else {
                $ExchangeServer = $null; $ExchangeServer = Get-ex16ExchangeServerTDO -Identity $Identity
            }

            # Validate Redirection Server
            $RedirectionServer = $null;
            if ($RedirectionTarget) {
                if($ExchangeServer.ServerRole -eq 'Edge'){
                    $smsg = "Edge Role detected, with -RedirectionServer!"
                    $smsg += "`nRedirect-Message *can* redirect between *UNSUBSCRIBED* edge servers, with suitable cross perms," ;
                    $smsg += "`nBut in general, it's not a function used with normal SUBSCRIBED EDGE ROLE" ;
                    $smsg += "`nBLANKING SPECIFIED -REDIRECTIONSERVER SPEC:$($RedirectionServer)" ;
                    $smsg += "`n(HW LoadBalancing state handling should detect down status on role and steer traffic redundantly)" ;
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $RedirectionServer = $NULL ; 

                } else { 
                    TRY {
                        $RedirectionServer = Get-ex16ExchangeServerTDO -Identity $RedirectionTarget
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        throw $smsg ;
                    }
                } ; 
            }
            if($ExchangeServer.ServerRole -eq 'Edge'){
                $smsg = "(ServerRole:Edge, skipping DAG checks)" ;
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
            } else { 
                # Determine DAG membership
                $isDAGMember = $false
                TRY {
                    $RecipientServer = $null; $RecipientServer = Get-MailboxServer -identity $Exchangeserver.fqdn -erroraction stop
                    if ($($RecipientServer | measure-object).count -ne 1) {
                        $smsg = "$($($RecipientServer | measure-object).count) servers returned from query. Unable to continue."
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        throw $smsg ;
                    }
                    if ($RecipientServer.DatabaseAvailabilityGroup -ne $null) {
                        $isDAGMember = $true
                    }
                } CATCH {
                    throw $_.exception.message
                }
            } ;  # if-E isDAG
            # Draining queues
            $smsg = "Putting '$($ExchangeServer.fqdn.toupper())' into maintenance mode."
            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } ;
            if ($RedirectionServer){
                $smsg = "Using '$($RedirectionServer.fqdn.toupper())' for message redirection."
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
            }
            $smsg = "Draining mail queues."
            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } ;
            TRY {
                Set-ServerComponentState -Identity $($ExchangeServer.fqdn) -Component HubTransport -State Draining -Requester Maintenance -erroraction stop
            } CATCH {
                throw $_.exception.message
            }

            # Restarting transport services
            if($ExchangeServer.ServerRole -eq 'Edge'){
                $smsg = "Edge:Restarting MSExchangeTransport services."                
            }else{
                $smsg = "Restarting MSExchangeTransport and MSExchangeFrontEndTransport services."
            } ; 
            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } ;
            $n = 0
            Do {
                TRY {
                    if($ExchangeServer.ServerRole -eq 'Edge'){
                        # remote icm doesn't work for Edge
                        "MSExchangeTransport"| restart-service -WarningAction SilentlyContinue
                    }else{
                        invoke-command -ComputerName $($ExchangeServer.fqdn) -scriptblock {"MSExchangeTransport","MSExchangeFrontEndTransport" | restart-service -WarningAction SilentlyContinue} -ErrorAction stop -WarningAction SilentlyContinue
                    } ; 
                    break

                } CATCH {
                    $n++
                    if($ExchangeServer.ServerRole -eq 'Edge'){
                        $smsg = "WARNING: Issue restarting MSExchangeTransportservice. Waiting 60 seconds then retrying." #-nonewline -ForegroundColor Yellow
                    }else{
                        $smsg = "WARNING: Issue restarting MSExchangeTransport and MSExchangeFrontEndTransport services. Waiting 60 seconds then retrying." #-nonewline -ForegroundColor Yellow
                    } ; 
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;
                    start-sleep -Seconds 60
                    $smsg = " Retry attempt $n of 3." #-ForegroundColor Yellow
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;
                }
                if ($n -eq 3) {
                    $smsg = "Issue restarting MSExchangeTransport and MSExchangeFrontEndTransport services. Continuing."
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    break
                }
            } while ($true)

            # Redirect messages
            <# 5:46 PM 8/16/2025: applic with Edge: *can* but only unsubscribed, disable it with edge for now.
            #-=-=-=-=-=-=-=-=
            [Redirect-Message (ExchangePowerShell) | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/redirect-message?view=exchange-ps)
            > Use the Redirect-Message cmdlet to drain the active messages from all the delivery queues on a Mailbox server, and transfer those messages to another Mailbox server.
            #-=-=-=-=-=-=-=-=
            [Redirect-Message edge server - Google Search](https://www.google.com/search?client=firefox-b-1-d&q=Redirect-Message+edge+server)
            #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
            AI Overview
            The Redirect-Message cmdlet in Exchange Server is used to redirect active messages from the delivery queues of one Mailbox server to another. This is typically done during maintenance or troubleshooting to move email traffic off a server that is being taken offline. While Redirect-Message is primarily used with Mailbox servers, it can also interact with Edge Transport servers in certain scenarios, particularly when managing mail flow between an Exchange organization and an unsubscribed Edge Transport server. 
            Here's a more detailed explanation:
            1. Redirecting Messages with Redirect-Message:

                The Redirect-Message cmdlet in the Exchange Management Shell is used to move messages from one Mailbox server to another.
                This cmdlet drains all active messages from the delivery queues on the source server and routes them to the target server.
                The source server stops accepting new messages while the redirection process is ongoing.
                Only active messages are redirected; shadow queues and poison messages are not affected.

            2. Interaction with Edge Transport Servers:

                Edge Transport Servers and Mail Flow:
                .

            Edge Transport servers act as a perimeter network for Exchange, handling inbound and outbound mail flow. 
            Edge Subscriptions:
            .
            Typically, Edge Transport servers are subscribed to an Active Directory site within the Exchange organization, which automates mail flow and configuration. 
            Unsubscribed Edge Servers:
            .
            In situations where an Edge Transport server is not subscribed, you may need to manually configure Send and Receive connectors to manage mail flow. 
            Redirecting Messages to/from Edge Servers:
            .
            While Redirect-Message primarily moves messages between Mailbox servers, it can be involved in scenarios where you need to redirect mail flow to or from an unsubscribed Edge Transport server. For example, you might redirect messages to an Edge server for processing (like address rewriting) or redirect messages from an Edge server to a Mailbox server for delivery.

            3. Example Usage (with Edge Transport):

                If you need to take an Edge Transport server offline for maintenance, you might redirect messages to a different Edge Transport server or back to Mailbox servers within the organization.

            You could use Redirect-Message in conjunction with other cmdlets like Set-ServerComponentState (to put the Hub Transport service into "Draining" mode) to manage the process.

            4. Key Considerations:

                Permissions: You need the appropriate permissions to run Redirect-Message.

            Target Server: The target server for redirection must be online and accessible. 
            Redirection Scope: Only active messages are redirected; shadow queues and poison messages are not affected.

            In summary, Redirect-Message is a powerful cmdlet for managing message flow between Exchange servers. While primarily used for Mailbox servers, it can be part of the process for managing mail flow involving Edge Transport servers, especially when dealing with unsubscribed Edge servers or maintenance scenarios. 

            #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
            #>
            if ($RedirectionServer) {
                $smsg = "Redirecting messages."
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
                TRY {
                    Redirect-Message -Server $($ExchangeServer.fqdn) -Target $($RedirectionServer.fqdn) -confirm:$false -erroraction stop -WarningAction SilentlyContinue
                } CATCH {
                    throw $_.exception.message
                }
            }

            # DAG members only
            if ($isDAGMember) {

                # Move active database copies off
                TRY {
                    $smsg = "Setting DatabaseCopyActivationDisabledAndMoveNow to 'True'."
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;
                    Set-MailboxServer -Identity $($ExchangeServer.fqdn) -DatabaseCopyActivationDisabledAndMoveNow $True -erroraction Stop -confirm:$false
                } CATCH {
                    throw $_.exception.message
                }

                # Move active copies immediately
                TRY {
                    $actives = $null; $actives = Get-MailboxDatabaseCopyStatus *\$($ExchangeServer.name) | ? {$_.activecopy -eq $true}
                    $smsg = "$($($actives | measure-object).count) active database copies found."
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;

                    if ($($($actives | measure-object).count) -eq 0 -and $MoveActiveDatabaseCopies) {
                        $smsg = "No active database copies to move."
                        if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ;
                    }

                    if ($actives -and $MoveActiveDatabaseCopies) {
                        $smsg = "Moving active databases to other DAG members."
                        if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ;
                        $actives | . {
                            process {
                                if ($($($($_ | . { process {(get-Mailboxdatabase $_.databasename).servers}}) | measure-object).count) -lt 2) {
                                    $smsg = "No other database copies exist. Unable to move active database copy."
                                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    } ;
                                } else {
                                    $move = $null;
                                    TRY {
                                        $move = Get-Mailboxdatabase $($_.databasename) |  Move-ActiveMailboxDatabase -MountDialOverride lossless -SkipClientExperienceChecks -SkipMaximumActiveDatabasesChecks -confirm:$false -erroraction stop
                                        if ($move.status -ne "Succeeded") {
                                            $smsg = "$($move.identity) Issue moving active database copy."
                                            if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                            } ;
                                            throw $smsg ;
                                        }
                                    } CATCH {
                                        write-warning $_.exception.message
                                    }
                                }
                            }
                        }
                    }

                } CATCH {
                    throw $_.exception.message
                }

                # Suspend cluster node
                $smsg = "Suspending cluster node."
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
                TRY {
                    invoke-command -ComputerName $($ExchangeServer.fqdn) -ScriptBlock {
                        if ((Get-ClusterNode $($using:ExchangeServer.fqdn)).state -ne "Paused") {
                            Suspend-ClusterNode $($using:ExchangeServer.fqdn)
                        }
                    } -ErrorAction Stop | out-null
                } CATCH {
                    throw $_.exception.message
                }

                # Set activation policy to blocked
                TRY {
                    $smsg = "Setting DatabaseCopyAutoActivationPolicy to 'Blocked'."
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;
                    Set-MailboxServer -Identity $($ExchangeServer.fqdn) -DatabaseCopyAutoActivationPolicy Blocked -erroraction Stop -confirm:$false
                } CATCH {
                    throw $_.exception.message
                }

                # Suspend passive copies
                TRY {
                    $Copies = $null; $Copies = Get-MailboxDatabaseCopyStatus *\$($ExchangeServer.name) | ? {$_.activecopy -eq $false}
                    if ($Copies) {
                        $smsg = "Suspending passive copies."
                        if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ;
                        $Copies | . { process {
                                $_  | Suspend-MailboxDatabaseCopy -confirm:$false -erroraction stop
                            }
                        }
                    }
                } CATCH {
                    throw $_.exception.message
                }

            }

            # Complete maintenance mode
            TRY {
                $smsg = "Completing maintenance mode."
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
                Set-ServerComponentState -Identity $($ExchangeServer.fqdn) -Component ServerWideOffline -State Inactive -Requester Maintenance -erroraction stop
            } CATCH {
                throw $_.exception.message
            }

            $smsg = "Done."
            $smsg += "`nConfirming Status: Get-ex16MaintenanceModeTDO -Identity $($ExchangeServer.fqdn)" ;
            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } ;

            # 12:30 PM 3/27/2025 add trailing stat confirmation:
            $testResults = Get-ex16MaintenanceModeTDO -Identity $ExchangeServer.fqdn ;
            $smsg = "Results`n$(($testResults|out-string).trim())" ; 
            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } ;
            $smsg = "Returning status to pipeline" ;
            if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            } ;
            $testResults | write-output  ; 

        }  # PROC-E
        END {
        } ;  # END-E
    }
#endregion START_EX16MAINTENANCEMODE ; #*------^ END Start-ex16MaintenanceMode ^------

