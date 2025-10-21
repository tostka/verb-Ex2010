# disable-xopExServer.ps1


#region DISABLE_XOPEXSERVER ; #*------v disable-xopExServer v------
function disable-xopExServer{
        <#
        .SYNOPSIS
        disable-xopExServer - Down and disable services on Exchange 2016 server, to keep offline during reboots, while performing non-Exchange maintenance. Supports feeding an array of computernames to do a group of servers.
        .NOTES
        REVISION
        * 12:48 PM 8/6/2025 add $rgxExSvcNamesFull = '^(MSEx|W3SVC|ClusSvc)', to capture broader svc state for returns ; capture svc changes via -passthru, echo to console ; added write-myOutput|Warning|Verbose support (for xopBuildLibrary/install-Exchange15.ps1 compat)
        * 9:33 AM 4/11/2025 add alias: cfmdt (reflects standard verbalias)
        .DESCRIPTION

        Down and disable services on Exchange 2016 server, to keep offline during reboots, while performing non-Exchange maintenance. Supports feeding an array of computernames to do a group of servers.

        Specifically disables Automatic starupt services: MSExchangeADTopology|MSExchangeEdgeCredential & w3svc (IIS), which will keep the rest of the Exchange services from restarting after reboots, 
    
        Afterward, restore function, renable and start services using: 

        PS> start-Ex16ServicesInDependencyOrder

        [Overview of Exchange services on Exchange servers | Microsoft Learn](https://learn.microsoft.com/en-us/exchange/plan-and-deploy/deployment-ref/services-overview)

        .PARAMETER Identity
        Specify the identity of the computer.
        .PARAMETER whatIf
        Whatif switch (defaults TRUE, -whatif:`$false to override) [-whatIf]
        .PARAMETER force
        Added force paramter to skip YYY prompts
        .INPUTS
        None. Does not accepted piped input.(.NET types, can add description)
        .OUTPUTS
        None. Returns no objects or output (.NET types)
        .EXAMPLE
        PS> disable-xopExServer -Identity $env:COMPUTERNAME -whatif:$false -verbose ;
        Run with verbose, override default -whatif
        .EXAMPLE
        PS> Get-ExchangeServer Server1  |  select -expand fqdn | disable-xopExServer
        .LINK
        https://github.com/tostka/powershellBB/    
        #>
        [cmdletbinding()]
        PARAM (
            [Parameter(mandatory=$true,ValueFromPipeline=$true,Helpmessage="Specify the identity of the computer.")]
                [string[]]$Identity,
            [Parameter(HelpMessage="Whatif switch (defaults TRUE, -whatif:`$false to override) [-whatIf]")]
                [switch]$whatif,
            [Parameter(HelpMessage="Force switch (overrides prompts) [-force]")]
                [switch]$force
        )
        BEGIN{
            $prpSvc = 'Status','Name','DisplayName','StartType' ;
            $rgxExSvcNamesFull = '^(MSEx|W3SVC|ClusSvc)' ; 
            $smsg = "THIS SCRIPT WILL *DOWN* AND SERVICE-DISABLE THE SPECIFIED LIST OF SERVERS!: $($Identity -join ',')" ; 
            if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;
            if(-not $force){
                $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ; 
                if ($bRet.ToUpper() -eq "YYY") {
                    $smsg = "(Moving on)" ; 
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } else {
                     $smsg = "Invalid response. Exiting" ; 
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    #exit 1
                    break ; 
                }  ; 
            } else { 
                $smsg = "-force specified: Skipping prompts, immediate down of target server" ; 
                if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
            } ; 
        } ; 
        PROCESS {
            foreach($server in $Identity){
                $smsg = $sBnrS="`n`n#*------v PROCESSING : $($server.toupper()) v------" ; 
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
                TRY {

                    $transcript = ".\logs" ; if(-not (test-path -path $transcript)){ mkdir $transcript -verbose:$true ; } ;
                    # - FOR FUNCTIONS, build from ${CmdletName}
                    #$transcript +=  "\$($CmdletName)" ;
                    # - OR FOR SCRIPTS, use $ScriptBaseName/$ScriptNameNoExt (which reflect name of hosting .psm1/.ps1 a function was loaded from)
                    $transcript +=  "\disable-xopExServer-$($server)" ;
                    # -- common v
                    if(get-variable whatif -ea 0){
                        $transcript += "-WHATIF-$(get-date -format 'yyyyMMdd-HHmmtt')-trans.txt" ; 
                        if(-not $whatif){$transcript = $transcript.replace('-WHATIF-','-EXECUTE-')} 
                    } ; 

                    if($transcript){
                        $stopResults = TRY {Stop-transcript -ErrorAction stop} CATCH {} ;
                        if($stopResults){
                            $smsg = "Stop-transcript:$($stopResults)" ; 
                            if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                                if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            } ;
                        } ; 
                        $startResults = start-Transcript -path $transcript ;
                        if($startResults){
                            $smsg = "start-transcript:$($startResults)" ; 
                            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                        } ; 
                    } else {
                        $smsg = "UNPOPULATED `$transcript! - ABORTING!" ; 
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ; 
                        throw $smsg ; 
                        break ; 
                    } ;  

                    $smsg = "Disable & Stop-services -force: MSExchangeADTopology" ; 
                    if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                        if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    } ;
                    $svcChanges = @() ; 
                    $svcChanges += @(get-service -Name MSExchangeADTopology -ComputerName $server -ea 0 | set-service -StartupType Disabled -WhatIf:$($whatif) -verbose -PassThru)  ;
                    $svcChanges += @(get-service -name MSExchangeADTopology -ComputerName $server -ea 0 | stop-service -force -WhatIf:$($whatif) -verbose -PassThru)  ;
                    $smsg = "Disable & Stop-services -force: w3svc (IIS)" ; 
                    if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                            if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        } ;
                    $svcChanges += @(get-service -Name w3svc -ComputerName $server -ea 0 | set-service -StartupType Disabled  -WhatIf:$($whatif) -verbose -PassThru) ;                
                    $svcChanges += @(get-service -name w3svc -ComputerName $server -ea 0|  stop-service -WhatIf:$($whatif) -verbose -PassThru)  ;
                    $smsg = "Stop-services all remaining running msex* services..." ; 
                    if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                        if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    } ;

                    $svcChanges += @(get-service -name msex* -ComputerName $server -ea 0| ? status -ne 'Stopped' | stop-service -WhatIf:$($whatif) -verbose -PassThru)  ;
                    $smsg = "Service Changes:`n$(($svcChanges | ft -a |out-string).trim())" ; 
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;

                    $smsg = "`n`n==>POST: get-service -name msex*`n$((get-service | ?{$_.Name -match '^(w3svc|msex)'}| select $prpSvc | ft -a |out-string).trim())`n`n" ; 
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;

                    # add services info as return
                    get-service | ?{$_.Name -match $rgxExSvcNamesFull}| select $prpSvc | write-output ; 
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } ; 
                #region STOP_TRANS ; #*------v STOP_TRANS v------
                $stopResults = TRY {Stop-transcript -ErrorAction stop} CATCH {} ;
                if($stopResults){
                    $smsg = "Stop-transcript:$($stopResults)" ; 
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } ; 
                #endregion STOP_TRANS ; #*------^ END STOP_TRANS ^------
                $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))`n`n" ;
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            } ; # loop-E
        } # PROC-E
    }
#endregion DISABLE_XOPEXSERVER ; #*------^ END disable-xopExServer ^------

