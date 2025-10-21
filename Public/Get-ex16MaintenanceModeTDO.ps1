# Get-ex16MaintenanceModeTDO.ps1


#region GET_EX16MAINTENANCEMODETDO ; #*------v Get-ex16MaintenanceModeTDO v------
function Get-ex16MaintenanceModeTDO {
        <#
        .SYNOPSIS
        Checks if a Microsoft Exchange Server 2016 computer is in maintenance mode. (only checks DAG members, skips any other role or non-DAG mailbox role!)
        .NOTES
        Version     : 0.0.1
        Author      : PietroCiaccio
        Website     : https://github.com/PietroCiaccio/
        Twitter     : 
        CreatedDate : 2025-03-19
        FileName    : Get-ex16MaintenanceModeTDO.ps1
        License     : MIT License
        Copyright   : (c) 2025 Todd Kadrie
        Github      : https://github.com/tostka/verb-XXX
        Tags        : Powershell
        AddedCredit : Todd Kadrie
        AddedWebsite: http://www.toddomation.com
        AddedTwitter: @tostka / http://twitter.com/tostka
        REVISIONS
        * 5:32 PM 8/16/2025 Get-ex16MaintenanceModeTDO(): add Edge detection support
        * 10:45 AM 8/6/2025 added write-myOutput|Warning|Verbose support (for xopBuildLibrary/install-Exchange15.ps1 compat) 
        * 1:40 PM 8/5/2025 added connect-xop...()
        * 9:04 AM 7/24/2025 ren: Get-EPMaintenanceMode -> Get-ex16MaintenanceModeTDO (alias orig) ; 
        updated return obj, now includes raw component status and a xxxFmt formated output (visible in dumps wo manual expansion); added ServerWideOffline that can be used as central check (assuming it can't be set if other components are active)
        * 4:30 PM 7/23/2025 rejiggered to output a customobject with more useful sub-properites, and easier to review info.
        * 12:52 PM 3/27/2025 TK: added aggregated Tests summary, returned to pipeline (to evaluate status, for follow-on processing).
        * 8/12/2020 Pietro Ciaccio's PSG-posted ExchangePowerShell module, v0.11.0
        .DESCRIPTION
            Checks if a Microsoft Exchange Server 2016 computer is in maintenance mode.
        .PARAMETER Identity
            Specify the identity of the computer. This can be piped from Get-ExchangeServer or specified explicitly using a string.
        .PARAMETER KeyComponents
        Key Components that are critical for 'Down' status of a server - to prevent CAS access or mail-handling (defaults to 'ServerWideOffline|HubTransport|FrontendTransport|AutoDiscoverProxy|ActiveSyncProxy|EcpProxy|EwsProxy|ImapProxy|OabProxy|OwaProxy|PopProxy|RpsProxy|RpcProxy|MapiProxy|EdgeTransport|MailboxDeliveryProxy')
        .OUTPUTS
        Returns a PSCustomObject to pipeline, summarizing status of tested components.
        .EXAMPLE
        $testResults = Get-ex16MaintenanceModeTDO -identity SERVER1 ; 
        #>
        [cmdletbinding()]
        [Alias('Get-EPMaintenanceMode')]
        PARAM (
            [Parameter(mandatory=$true,valuefrompipelinebypropertyname=$true)]
                [PSCustomObject]$Identity,
            [Parameter(mandatory=$false,HelpMessage="Key Components that are critical for 'Down' status of a server - to prevent CAS access or mail-handling (defaults to 'ServerWideOffline|HubTransport|FrontendTransport|AutoDiscoverProxy|ActiveSyncProxy|EcpProxy|EwsProxy|ImapProxy|OabProxy|OwaProxy|PopProxy|RpsProxy|RpcProxy|MapiProxy|EdgeTransport|MailboxDeliveryProxy')")]
                [ValidateNotNullOrEmpty()]
                [string[]]$KeyComponents = @('ServerWideOffline','HubTransport','FrontendTransport','AutoDiscoverProxy','ActiveSyncProxy','EcpProxy','EwsProxy','ImapProxy','OabProxy','OwaProxy','PopProxy','RpsProxy','RpcProxy','MapiProxy','EdgeTransport','MailboxDeliveryProxy')
        )
        BEGIN{
            [regex]$rgxKeyComponents = [regex]$rgx = ('(' + (($KeyComponents |%{[regex]::escape($_)}) -join '|') + ')') ;
            if(-not (gcm Set-ServerComponentState -ea 0)){
                if(connect-XopLocalManagementShell){ 
                    $smsg = "Connected" 
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;

                } else { 
                    $smsg = "NOT CONNECTED!"
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } ;
            } ; 
        } ; 
        PROCESS {            
            $Summary = [ordered]@{
                ExchangeServer = $null ;
                isMailboxServer = $false ;
                isDAGMember = $false ;
                DatabaseCopyActivationDisabledAndMoveNow = $null ;
                ClusternodeUp = $null ;
                DatabaseCopyAutoActivationPolicy = $null ;
                MailboxDatabaseCopyStatus = $null ;
                isEdgeServer = $false ; # 5:23 PM 8/16/2025 add Edge reporting, it's in the $ExchangeServer.ServerRole: Edge
                ServerComponentStateActive = $null ;
                ServerComponentStateActiveFmt = $null ;
                ServerComponentStateINActive = $null ;
                ServerComponentStateINActiveFmt = $null ;
                KeyComponentsState = $null ;
                KeyComponentsStateFmt = $null ;
                ServerWideOffline = $null
            }

            # Validate identity
            if ($input) {
                $ExchangeServer = $null; $ExchangeServer = $input | Get-ex16ExchangeServerTDO
            } else {
                $ExchangeServer = $null; $ExchangeServer = Get-ex16ExchangeServerTDO -Identity $Identity
            }
            $smsg = "$($ExchangeServer.fqdn.toupper())."
            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;

            $Summary.ExchangeServer = $ExchangeServer.fqdn.toupper() ; 
            if($ExchangeServer.ServerRole -eq 'Edge'){
                $Summary.isEdgeServer = [boolean]($ExchangeServer.ServerRole -eq 'Edge') ; 
            } else {
                # Determine DAG membership
                $isDAGMember = $false
                TRY {
                    $RecipientServer = $null; 
                    $RecipientServer = Get-MailboxServer -identity $Exchangeserver.fqdn -erroraction stop
                    if ($($RecipientServer | measure).count -ne 1) {
                        $smsg = "$($($RecipientServer | measure).count) servers returned from query. Unable to continue."
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        throw $smsg ;
                    } else {
                        $Summary.isMailboxServer = $true ; 
                    } 
                    if ($RecipientServer.DatabaseAvailabilityGroup -ne $null) {
                        $Summary.isDAGMember = $isDAGMember = $true                    
                    }
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;                
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } ; 
            } ;  # if-E Edge/Mailbox role 
            if($RecipientServer){
                TRY {
                    # DAG members only
                    if ($isDAGMember) {
                        $MBServer = $null; $MBServer = $ExchangeServer | Get-MailboxServer -erroraction stop
                        if ($MBServer.DatabaseCopyActivationDisabledAndMoveNow -eq $false) {
                            $smsg = "DatabaseCopyActivationDisabledAndMoveNow is False."     
                            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;                 
                        } else {
                            $smsg = "DatabaseCopyActivationDisabledAndMoveNow is True."     
                            if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;                   
                        }
                        $Summary.DatabaseCopyActivationDisabledAndMoveNow = $($MBServer.DatabaseCopyActivationDisabledAndMoveNow)
                        $cn = $null; $cn = invoke-command -ComputerName $($ExchangeServer.fqdn) -ScriptBlock {Get-ClusterNode $($using:ExchangeServer.fqdn)} -ErrorAction Stop
                        if ($cn.state -eq "up") {
                            $smsg = "Cluster node is Up." 
                            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;                     
                            $Summary.ClusternodeUp = $true ; 
                        } else {
                            $smsg = "Cluster node is not Up."   
                            if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;                     
                            $Summary.ClusternodeUp = $false ; 
                        }
                        if ($MBServer.DatabaseCopyAutoActivationPolicy -eq "unrestricted") {
                            $smsg = "DatabaseCopyAutoActivationPolicy is Unrestricted."
                            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            } ;
                        } else {
                            $smsg = "DatabaseCopyAutoActivationPolicy is $($MBServer.DatabaseCopyAutoActivationPolicy)."   
                            if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                        }
                        $Summary.DatabaseCopyAutoActivationPolicy = $MBServer.DatabaseCopyAutoActivationPolicy ; 
                        $Copies = $null; 
                        $Copies = Get-MailboxDatabaseCopyStatus *\$($ExchangeServer.name)
                        if ($Copies) {
                            $Copies | . { process {
                                if ($_.status -notmatch "^healthy$|^mounted$") {
                                    $smsg = "$($_.name) database copy status is $($_.status)."                                
                                    $Summary.MailboxDatabaseCopyStatus += "$($_.name) database copy status:$($_.status.toUpper())" ; 
                                } else {
                                    $smsg = "$($_.name) database copy status is $($_.status)."   
                                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                    } ;                             
                                    $Summary.MailboxDatabaseCopyStatus += "$($_.name) database copy status:$($_.status.toUpper())."
                                }
                            }}
                        }
                    } # dag only
                 
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } ; 
            };  # if-E $RecipientServer
            # move the SCS outside the DAG tests
            TRY{
                $CS = $null; 
                #$CS = Get-ServerComponentState -Identity $($ExchangeServer.fqdn) -erroraction stop | ? {$_.state -ne "active"}
                $CS = Get-ServerComponentState -Identity $($ExchangeServer.fqdn) -erroraction stop ; 
                #if (-not($CS | ? {$_.state -ne "active"})) {
                if ($CS | ? {$_.state -eq "active"}) {
                    $smsg = "Server component states active: $(($CS  | ? {$_.state -EQ "active"}).component -join ';')" 
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;                   
                    $Summary.ServerComponentStateActive += $CS  | ? {$_.state -eq "active"}; 
                    $Summary.ServerComponentStateActiveFmt += @("Server component states ACTIVE:`n$(($CS  | ? {$_.state -eq "active"}|ft -a | out-string).trim())") ; 
                }
                #} else {
                if ($CS | ? {$_.state -ne "active"}) {                    
                    $Summary.ServerComponentStateINActive += $CS  | ? {$_.state -ne "active"} ; 
                    $Summary.ServerComponentStateINActiveFmt += @("Server component states INACTIVE:`n$(($CS  | ? {$_.state -ne "active"}|ft -a | out-string).trim())") ; 
                }
                $Summary.KeyComponentsState = $cs | ?{$_.component -match $rgxKeyComponents} ; 
                $Summary.KeyComponentsStateFmt = @("Key component states:`n$(($cs | ?{$_.component -match $rgxKeyComponents}|ft -a | out-string).trim())") ; 
                $Summary.ServerWideOffline = $cs | ?{$_.component -eq 'ServerWideOffline'} ; 
                $smsg = "Done."
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
                write-verbose "Returning test results to pipeline" ; 
                [pscustomobject]$Summary | write-output ; 
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            } ; 
        } ;  # PROC-E
    }
#endregion GET_EX16MAINTENANCEMODETDO ; #*------^ END Get-ex16MaintenanceModeTDO ^------

