# Stop-ex16MaintenanceMode.ps1


#region STOP_EX16MAINTENANCEMODE ; #*------v Stop-ex16MaintenanceMode v------
function Stop-ex16MaintenanceMode{
        <#
        .SYNOPSIS
        Stop-ex16MaintenanceMode.ps1 - Takes a Microsoft Exchange Server 2016 computer out of maintenance mode.
        .NOTES
        Version     : 0.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2025-03-19
        FileName    : Stop-ex16MaintenanceMode.ps1
        License     : MIT License
        Copyright   : (c) 2025 Todd Kadrie
        Github      : https://github.com/tostka/verb-XXX
        Tags        : Powershell
        AddedCredit : PietroCiaccio
        AddedWebsite: https://github.com/PietroCiaccio/
        AddedTwitter: URL
        REVISIONS
        * 5:53 PM 8/16/2025 Stop-ex16MaintenanceMode: add Edge detection support; code to do the $isDAG test within the function (and not inherited)
        * 2:30 PM 3/26/2025 merge PietroCiaccio's Disable-EPMaintenanceMode with my updates under Start-ex16MaintenanceMode.ps1
        * 9:09 AM 3/26/2025 added ipsn post connect-ExchangeServer() call (wasn't including import/non-functional [headscratch]); beyond that, used wo issues on 1st of the Ex16 builds
        * 8/12/2020 Pietro Ciaccio's PSG-posted ExchangePowerShell module, v0.11.0
        .DESCRIPTION
        Stop-ex16MaintenanceMode.ps1 - Takes a Microsoft Exchange Server 2016 computer out of maintenance mode. CmdLet will: 

                - set all server component states to active
                - resume the cluster node
                - enable database activation on the server
                - resume passive database copies
                - resume transport
                - restart transport services 

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
        PS> .\Stop-ex16MaintenanceMode.ps1 -whatif -verbose
        EXSAMPLEOUTPUT
        Run with whatif & verbose
        .EXAMPLE
        PS> .\Stop-ex16MaintenanceMode.ps1
        EXSAMPLEOUTPUT
        EXDESCRIPTION
        .LINK
        https://github.com/tostka/powershellBB/
        #>
        # #Requires -PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
        [CmdletBinding()]
        Param (
            [Parameter(mandatory=$false,valuefrompipelinebypropertyname=$true,
                HelpMessage="Specify the identity of the Exchange Server. This can be piped from Get-ExchangeServer or specified explicitly using a string")]
                [ValidateNotNullOrEmpty()]
                [PSCustomObject]$Identity=$env:computername
        ) ; 
        BEGIN{
            if(-not (gcm Set-ServerComponentState -ea 0)){
                if(connect-XopLocalManagementShell){
                     write-host -foregroundcolor green "Connected" } 
                else { 
                    $smsg = "NOT CONNECTED!" ; 
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    BREAK} ;
            } ; 
        } # BEG-E
        PROCESS{
            # Validate identity
            # In the process block, the $input variable contains the current object in the pipeline.

            if ($input) {            
                $ExchangeServer = $null; $ExchangeServer = $input | Get-ex16ExchangeServerTDO
            } else {
                $ExchangeServer = $null; $ExchangeServer = Get-ex16ExchangeServerTDO -Identity $Identity
            }

            # Remove from maintenance mode
            TRY {
                $smsg = "Removing '$($ExchangeServer.fqdn.toupper())' from maintenance mode."
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
                Set-ServerComponentState -Identity $($ExchangeServer.fqdn) -Component ServerWideOffline -State Active -Requester Maintenance -erroraction stop
            } CATCH {
                throw $_.exception.message
            } 
            
            if($ExchangeServer.ServerRole -eq 'Edge'){
                $smsg = "(ServerRole:Edge, skipping DAG checks)" ;
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
            } else { 
                # 5:52 PM 8/16/2025 doesn't have $isDAGmember populating code in this func - shouldn't be inheriting from external tests
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

                # DAG members only
                if ($isDAGMember) {
            
                    # Resume cluster node
                    $smsg = "Resuming cluster node."  
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;     
                    TRY {
                        invoke-command -ComputerName $($ExchangeServer.fqdn) -ScriptBlock {
                            if ((Get-ClusterNode $($using:ExchangeServer.fqdn)).state -ne "Up") {
                                Resume-ClusterNode $($using:ExchangeServer.fqdn)
                            }     
                        } -ErrorAction Stop | out-null
                    } CATCH {
                        throw $_.exception.message
                    }

                    # Move active database copies on
                    TRY {
                        $smsg = "Setting DatabaseCopyActivationDisabledAndMoveNow to 'False'."
                        if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ;
                        Set-MailboxServer -Identity $($ExchangeServer.fqdn) -DatabaseCopyActivationDisabledAndMoveNow $false -erroraction Stop -confirm:$false
                    } CATCH {
                        throw $_.exception.message
                    }

                    # Set activation policy to unrestricted
                    TRY {
                        $smsg = "Setting DatabaseCopyAutoActivationPolicy to 'Unrestricted'."
                        if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ;
                        Set-MailboxServer -Identity $($ExchangeServer.fqdn) -DatabaseCopyAutoActivationPolicy Unrestricted -erroraction Stop -confirm:$false
                    } CATCH {
                        throw $_.exception.message
                    }   
            
                    # Resume passive copies
                    TRY {                
                        $Copies = $null; $Copies = Get-MailboxDatabaseCopyStatus *\$($ExchangeServer.name) | ? {$_.activecopy -eq $false}
                        if ($Copies) {
                            $smsg = "Resuming passive copies."
                            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            } ;
                            $Copies | . { process {
                                    $_  | Resume-MailboxDatabaseCopy -confirm:$false -erroraction stop
                                }                
                            }
                        }
                    } CATCH {
                        throw $_.exception.message
                    }

                }     
            } ;  # if-E Edge/Mailbox detect
            # Resume transport
            $smsg = "Resuming transport."
            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } ;
            TRY {
                Set-ServerComponentState -Identity $($ExchangeServer.fqdn) -Component HubTransport -State Active -Requester Maintenance -erroraction stop
            } CATCH {
                throw $_.exception.message
            }

            # Restarting transport services
            $smsg = "Restarting MSExchangeTransport and MSExchangeFrontEndTransport services."
            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } ;
            $n = 0
            Do {
                TRY {
                    invoke-command -ComputerName $($ExchangeServer.fqdn) -scriptblock {"MSExchangeTransport","MSExchangeFrontEndTransport" | restart-service -WarningAction SilentlyContinue} -ErrorAction stop -WarningAction SilentlyContinue
                    break
                } CATCH {
                    $n++
                    $smsg = "WARNING: Issue restarting MSExchangeTransport and MSExchangeFrontEndTransport services. Waiting 60 seconds then retrying." #-nonewline -ForegroundColor Yellow
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;
                    Start-Sleep -Seconds 60
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

            $smsg = "Done." 
            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } ;
  
        
        }  # PROC-E
        END {
        } ; 
    }
#endregion STOP_EX16MAINTENANCEMODE ; #*------^ END Stop-ex16MaintenanceMode ^------

