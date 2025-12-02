# connect-XopLocalManagementShell.ps1


#region CONNECT_XOPLOCALMANAGEMENTSHELL ; #*------v connect-xopLocalManagementShell v------
    #if (-not(gi function:connect-xopLocalManagementShell -ea 0)) {
        function connect-XopLocalManagementShell {
            <#
            .SYNOPSIS
            connect-XopLocalManagementShell - Detects *local* machine has Exchange Server installed, and role (Mailbox v Edge); resolves & configs dependant Env Varis; configures and loads local Exchange Management Shell (EMS) connection into the server.
            .NOTES
            Version     : 0.0.1
            Author      : Todd Kadrie
            Website     : http://www.toddomation.com
            Twitter     : @tostka / http://twitter.com/tostka
            CreatedDate : 2025-03-19
            FileName    : connect-XopLocalManagementShell
            License     : MIT License
            Copyright   : (c) 2025 Todd Kadrie
            Github      : https://github.com/tostka/verb-XXX
            Tags        : Powershell
            AddedCredit : PietroCiaccio
            AddedWebsite: https://github.com/PietroCiaccio/
            AddedTwitter: URL
            REVISIONS
            * 2:47 PM 12/2/2025 💡 updated the CBH demo to test for missing cmdlet, before doing reimport (conditional on actual fail; 
                imports are needed when this is called out of the .psm1 by another freestanding .ps1; 
                tends to work fine wo remedial ip from funcs inside the psm1). Tested in latest set-exLicense.
            * 4:13 PM 10/7/2025 updated CBH non-func $idfqdn for $env:computername in pss test ; 
                for demo of new remedial call import code; code to detect preexisting pssessions, and skip rexec redund, also to remove those as they accumulate; 
            * 5:12 PM 8/16/2025 connect-XopLocalManagementShell():Edge lacks RemoteExchange.ps1, added regkey edge test to exempt test for the file (was causing premature throw).
            * 10:45 AM 8/6/2025 added write-myOutput|Warning|Verbose support (for xopBuildLibrary/install-Exchange15.ps1 compat)
            * 9:40 AM 8/5/2025 add pop missing $env:ExchangeInstallPath; confirm dep RemoteExchange.ps1 is accessible
            * 8:59 AM 7/30/2025 add CBH xmpl, test cmdlet trigger
            * 3:36 PM 7/24/2025 defaulted ExVers Ex2016
            * 1:13 PM 7/14/2025 ADDED -identity; updated code from test-xopLocal
            * 1:08 PM 3/27/2025 added min items to make it a functional Adv Function (w verbose support etc)
            * 9:09 AM 3/26/2025 added ipsn post connect-ExchangeServer() call (wasn't including import/non-functional [headscratch]); beyond that, used wo issues on 1st of the Ex16 builds
            * 8/12/2020 Pietro Ciaccio's PSG-posted ExchangePowerShell module, v0.11.0
            .DESCRIPTION
            connect-XopLocalManagementShell - Detects *local* machine has Exchange Server installed, the role (Mailbox v Edge); resolves & configs dependant Env Varis; configures and loads local Exchange Management Shell (EMS) connection into the server. Then returns the PSSession object to the pipeline.
            Functions as a ground-up autodiscovery & configure EMS load, without pre-existing Exchange Server EMS shortcut. 
            .OUTPUT
            System.Management.Automation.Runspaces.PSSession (for successful non-Edge Exchange Server role connection)
            System.Management.Automation.PSModuleInfo (for  successful Edge Exchange Server role connection)
            .EXAMPLE
            PS> if(connect-XopLocalManagementShell){ write-host -foregroundcolor green "Connected" } else { write-warning "NOT CONNECTED!"} ;
            Simple test - returns session on connection.
            .EXAMPLE
            PS> $tcmdlet = 'Set-ClientAccessServer' ; 
            PS> $cmd = $null; $cmd = get-command $tcmdlet -erroraction 0 ;
            PS> if(-not $cmd){
            PS>     if($xopconn = connect-XopLocalManagementShell){
            PS>         if($ExPSS = get-pssession | ? { $_.ComputerName -match "^$($env:computername)" -AND $_.ConfigurationName -eq 'Microsoft.Exchange' } | sort id -Descending | select -first 1 ){
            PS>             TRY{
            PS>                 $cmd = $null; $cmd = get-command $tcmdlet -erroraction 0 ;
            PS>                 if(-not $cmd){
            PS>                     $smsg = "Missing $($tcmdlet): re-importing PSSession..." ;
            PS>                     if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
            PS>                         if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            PS>                         #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            PS>                     } ;
            PS>                     $ExIPSS = Import-PSSession $ExPSS -allowclobber -ea STOP ;
            PS>                 } ; 
            PS>                 $cmd = $null; $cmd = get-command 'Get-OrganizationConfig' -erroraction stop ;
            PS>                 $smsg = "Connected to: $($expss.computername)" ;
            PS>                 if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
            PS>                     if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            PS>                     #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            PS>                 } ;
            PS>             } CATCH {
            PS>                 $ErrTrapd=$Error[0] ;
            PS>                 $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
            PS>                 if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
            PS>                     if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            PS>                 } ;
            PS>                 BREAK ;
            PS>             } ;
            PS>         } ;
            PS>     } else {
            PS>         $smsg = "NOT CONNECTED!"
            PS>         if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
            PS>             if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
            PS>             else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            PS>         } ;
            PS>         BREAK ; 
            PS>     } ;
            PS> } ;            
            Well rounded demo that rediscovers out of scope/unimported open/avail pssessions and does a remedial import.
            .LINK
            https://github.com/tostka/powershellBB/    
            #>           
            [CmdletBinding()]
            PARAM(
                [Parameter(Mandatory = $false, HelpMessage = "Target Exchange OnPrem Server (defaults to `$env:computername)[-ExVers Ex2016]")]
                    [ValidateNotNullOrEmpty()]
                    [string]$identity = $env:COMPUTERNAME,
                [Parameter(Mandatory = $false, HelpMessage = "Target Exchange version string (ExSE|Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000)[-ExVers Ex2016]")]
                    [ValidateSet('ExSE', 'Ex2019', 'Ex2016', 'Ex2013', 'Ex2010', 'Ex2007', 'Ex2003', 'Ex2000')]
                    [string]$ExVers= 'Ex2016'
            )
            BEGIN{
                # make the vers a vari, for future ex19 recycling
                #$ExVers = "V15" ;
                # ExVers = [string] 'ExSE|Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000'
                switch -regex ($ExVers) {
                    'ExSE' { $ExInstallPathVers = 'V15' }
                    'Ex2019' { $ExInstallPathVers = 'V15' }
                    'Ex2016' { $ExInstallPathVers = 'V15' }
                    'Ex2013' { $ExInstallPathVers = 'V15' }
                    'Ex2010' { $ExInstallPathVers = 'V14' }
                    'Ex2007' { $ExInstallPathVers = 'V8' }
                    'Ex2003' { $ExInstallPathVers = 'V6' }
                    'Ex2000' { $ExInstallPathVers = 'V6' }
                    default {
                        $smsg = "Unrecognized Exchange Version: '$($ExVers)'!" ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        if(gcm Write-MyError -ea 0){Write-MyError $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error} else{ write-ERROR "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        throw $smsg ; 
                    }
                } ;
                $smsg = "Loading Exchange PowerShell Module..." ; 
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
                # populate the key stock xop EnvVaris            
                if (-not($env:ExchangeInstallPath)) {
                    #throw "Exchange Server system variable ExchangeInstallPath missing." ;
                    # stock Ex enviro variables if missing
                    if ($null -eq $ExInstall -OR $null -eq $ExBin) {
                        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\Setup') {
                            $Global:ExInstall = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\Setup).MsiInstallPath ;
                            $Global:ExBin = $Global:ExInstall + "\Bin"
                            $smsg = ("Set ExInstall: {0}" -f $Global:ExInstall)
                            # set Process 'temp' variable ; It's required for discovery below!
                            $env:ExchangeInstallPath = $Global:ExInstall ; 
                            $smsg += ("`nSet `$env:ExchangeInstallPath: {0}" -f $Global:ExInstall) ;
                            if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                                if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            } ;
                            $smsg = ("Set ExBin: {0}" -f $Global:ExBin)
                            if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                                if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            } ;
                        } else {
                            $smsg = "Exchange Server Install Path not found in registry! (HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\Setup)" ;
                            if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                            throw $smsg ;
                        } ;
                    }elseif($ExInstall){
                        # ExchangeInstallPath still blanked, dep for below
                        # set Process 'temp' variable ; It's required for discovery below!
                        $env:ExchangeInstallPath = $Global:ExInstall ; 
                        $smsg += ("`nSet `$env:ExchangeInstallPath: {0}" -f $Global:ExInstall) ;
                        if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                            if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        } ;
                        if($null -eq $ExBin) {
                            $Global:ExBin = $Global:ExInstall + "\Bin"                        
                        }
                    } ;  ;
                } ;
                # 5:11 PM 8/16/2025 Edge doesn't have RemoteExchange.ps1 in bin!, test the role key too.
                if(-not (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\EdgeTransportRole')){
                    if(test-path "$($env:ExchangeInstallPath)\bin\RemoteExchange.ps1" -PathType Leaf){
                        $smsg += "validated `$env:ExchangeInstallPath\bin\RemoteExchange.ps1" ; 
                        if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                            if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        } ;
                    }else{
                        $smsg = "INVALID $env:ExchangeInstallPath)\bin\RemoteExchange.ps1!" ; 
                        $smsg += "`n(may be Edge role, naturally lacks: further testing...)" ;
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        #THROW $smsg ; 
                    } ; 
                } else { 
                    $smsg += "Edge Role regkey: Skipping test for: `$env:ExchangeInstallPath\bin\RemoteExchange.ps1" ; 
                    if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                        if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    } ;
                } ; 
                # 1:36 PM 8/17/2025 fix: vari expan doesn't work in single quotes => dbl quot
                if ($env:ExchangeInstallPath -notmatch "\\$($ExVers)\\") {
                    $smsg = "The Microsoft Exchange Management Shell will be loaded from '$($env:ExchangeInstallPath)'" ;
                    $smsg += "`nExchange Server $($ExVers) powershell module not detected. There may be issues." ;
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } ;

                if (-not (Get-Service -ComputerName $env:computername -Name msex*)) {
                    $smsg = "No Exchange Servicese (msex*) installed on local machine! (get-service msex*)" ;
                    $smsg += "`n(any recently completed Ex Setup pass, did not successfully install Ex services)" ;
                    $smsg += "`n(this function cannot bring up *remote* EMS connections)" ;
                    $smsg = "Exchange Server 2016 powershell module not detected. There may be issues." ;
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } ;
            }  # BEG-E
            PROCESS{
                TRY {
                    # preremove broken:
                    get-pssession | ?{$_.State -eq 'Broken' -AND $_.Availability -eq 'None'} | foreach-object {
                      write-host "(RmvPSS:State:Broken & Availability:None)" ; $_ | Remove-PSSession -verbose ; 
                    } ; 
                    $cmd = $null; $cmd = get-command 'Get-OrganizationConfig' -erroraction stop ;
                } CATCH {
                    if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\EdgeTransportRole')){
                        $smsg = "Connecting to local EdgeTransportRole server"; 
                        # implement local snapins access on edge role: Only way to get access to EMS commands.
                        [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exshell.psc1" -ErrorAction Stop
                        ForEach($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn){
                            write-verbose ("Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name)
                            Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop
                        } ; 
                        $oIpMo = Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop -PassThru ; 
                        # Import-Module -PassThru: Returns an object representing the imported module. By default, this cmdlet doesn't generate any output.
                        $passed = $true #We are just going to assume this passed.
                    }else{
                        $smsg = "Connecting to local Exchange server (via resolved fqdn)"; 
                        # orig code: open auto-discovery, connects into legacy Ex by pref, needs to be steered.
                        #$invexpr = ". '$env:ExchangeInstallPath\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell" ;
                        # supports -ServerFqdn param, why let it auto into obso Ex versions?, resolve the $identity to a name, and then to the full fqdn (A record)
                        $idName = switch -regex ($identity.gettype().fullname) {
                            'System.Management.Automation.PSObject' {
                                if ($identity.name) { 
                                    $identity.name | write-output 
                                }else { 
                                    $smsg = "PSObject with no name property (get-exchangeserver piped output is supported)!" 
                                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    } ;
                                }
                            }
                            'System.String' { $identity | write-output }
                            default { 
                                $smsg = "Unrecognized -identity object type!" 
                                if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                } ;
                                throw $smsg; 
                            }
                        } ;
                        # Resolve FQDN from IP PTR: exclude autoassign DHCP subnets, pick first return if multiple DNS A type records on nbhostname, use default local DNS servers in stack.
                        if ($idFQDN = Resolve-DnsName -type A $idName | ? { $_.ipaddress -notmatch '^169\.254\.' } | select -first 1 | select -expand name) {
                            $smsg = "Resolve-DnsName: `$idFQDN:$($idFQDN)" ; 
                            if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                                if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            } ;
                        } else {
                            $smsg = "UNABLE TO Resolve-DnsName -type A $($idName)" ;
                            $smsg += "`nFalling back to `$env:COMPUTERNAME (nbname):$($env:COMPUTERNAME)" ;
                            if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                            $idFQDN = $env:COMPUTERNAME ; 
                        } ;
                        if ($idFQDN){
                            if(-not (get-pssession | ? { $_.ComputerName -match "^$($idFQDN)" -AND $_.ConfigurationName -eq 'Microsoft.Exchange' -AND $_.State -eq 'Opened' -AND $_.Availability -EQ 'Available'})){
                                $invexpr = ". '$env:ExchangeInstallPath\bin\RemoteExchange.ps1'; Connect-ExchangeServer -serverFQDN $idFQDN -ClientApplication:ManagementShell" ;
                                Invoke-Expression $invexpr ;
                            } ; ; 
                            # TK: above doesn't leave the session imported; requires manual import; can also have an array, sort on ExpiresOn (no, null), sort on id which is an incrmenting integer each new add, take most recent;  can't import an array of sessions                            
                            # add/enforce state & Avail, sort single latest
                            # if there are multiples, remove dupes, to stop accumulation
                            $ExPSS = get-pssession | ? { $_.ComputerName -match "^$($idFQDN)" -AND $_.ConfigurationName -eq 'Microsoft.Exchange' -AND $_.State -eq 'Opened' -AND $_.Availability -EQ 'Available'} ; 
                            if($ExPSS -is [array]){                                                                
                                # no xop ExpiresOn is empty, sort on Id which increments with each added session
                                $ExPSS | sort id -Descending | select-object -Skip 1 | remove-pssession -verbose  ; 
                                $ExPSS =  $ExPSS | sort id -Descending | select -first 1  ;                                 
                            } ; 
                            # iex will passthru content, IF the exec'd cmd does & uses -passthru; Undetermined if the MS Connect-ExchangeServer() supports passthru
                            <# [PowerShell Gallery | connect-ExchangeServer.ps1 1.2.6](https://www.powershellgallery.com/packages/ExchangeAntiSpamReport/1.2.6/Content/connect-ExchangeServer.ps1)
                                includes a spec for it:
                                [OutputType([Management.Automation.Runspaces.PSSession])]
                                => it *should*, but lets stick with discovery for belt & suspenders
                            #>
                            $ExIPSS = Import-PSSession $ExPSS -allowclobber ;
                            #$ExPSS | write-output ; # return the session to the pipeline - no, we're returning true/false further down
                            $smsg = "Connected to:`n$(($ExIPSS | ft -a |out-string).trim())" ;
                            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;

                        } else {
                            $smsg = "UNABLE TO Resolve-DnsName -type A $($idName), (or use functional `$ENV:computername)" ;
                            if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                        } ;
                    } ;  # if-E Edge/Normal server connection
                } ;

                <# issue: Ex16 Connect-ExchangeServer -auto -ClientApplication:ManagementShell"
            if it can't find the local - services down - it WILL DIVERT INTO older EXCH versions!
            WHICH DOESN'T SEE EX16 SERVERS IN GET-EXCHANGESERVER!

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
            #>
                # test we successfully connected to v15/Ex16+:
                #$ExVersNo = '15.1' ;
                switch -regex ($ExVers){
                    'ExSE'{$ExVersNo = '15.2'}
                    'Ex2019'{$ExVersNo = '15.2'}
                    'Ex2016'{$ExVersNo = '15.1'}
                    'Ex2013'{$ExVersNo = '15.0'}
                    'Ex2010'{$ExVersNo = '14.'}
                    'Ex2007'{$ExVersNo = '8.'}
                    'Ex2003'{$ExVersNo = '6.5'}
                    'Ex2000'{$ExVersNo = '6'}
                    default{
                        $smsg = "-ExVers $($ExVers) is *not* a recognized/configured version string in this script!" ; 
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ; 
                        throw $smsg ; 
                        break ; 
                    }
                } ; 
                if (Get-ExchangeServer -ea STOP | ? { $_.AdminDisplayVersion -Match ( [regex]::Escape("Version $($ExVersNo)")) }) {
                    $smsg = "(Confirmed v15 ExOP servers returned by Get-ExchangeServer)" ;
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } else {
                    $smsg = "Get-ExchangeServer DID NOT RETURN ANY V15+ SERVERS!" ;
                    $smsg += "`n(Get-ExchangeServer | ?{`$_.AdminDisplayVersion -Match '^Version 15'})" ;
                    $smsg += "`nABORTING: Make sure target server $($identity) msex* services are started & running!" ;
                    $smsg += "`n(clearing any mis-version PSSessions and the local function:get-mailbox command...)" ;
                    $smsg += "`nRemediate the target system, and re-run" ; ;
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    Get-PSSession | ? { $_.ConfigurationName -eq 'Microsoft.Exchange' } | Remove-PSSession -verbose ;
                    get-item function:get-mailbox -ea 0 | remove-item -force -verbose ;
                    break ;
                } ;
            } ;  # PROC-E
            END{
                TRY {
                    $cmd = $null; $cmd = get-command 'Get-OrganizationConfig' -erroraction stop ; # 'get-mailbox'
                    #$true | write-output ;
                    # TK, lets return the PSS
                    # edge won't have $ExIPSS populated, & non-Edge won't have $oIpMo ; return the -passthru captured $oIpMo for Edge confirm
                    if($cmd){
                        #if($ExIPSS){
                        if($ExIPSS){
                            $smsg = "(Returning Import-PSSession result to pipeline)" ;
                            if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                                if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            } ;
                            $ExIPSS | write-output ;
                        }elseif($oIpMo ){
                            $smsg = "(Returning Import-Module result to pipeline)" ;
                            if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                                if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            } ;
                            $oIpMo | write-output ; 
                        }elseif($cmd ){
                            $smsg = "(Returning gcm Get-OrganizationConfig result to pipeline)" ;
                            if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                                if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            } ;
                            $cmd | write-output ; 
                        }else{
                            $smsg = "(connect-xopLocalManagementShell:undefined misconnection issue!)" ;
                            if(gcm Write-MyError -ea 0){Write-MyError $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error} else{ write-ERROR "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                        }
                    }else {
                        $smsg = "NOT CONNECTED! (unable to gcm Get-OrganizationConfig & locate populated `$ExIPSS (Non-Edge role) or `$oIpMo (Edge role) to return)"
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        $false | write-output ;
                    }
                } CATCH {
                    $false | write-output
                    $smsg = "Unable to load the Microsoft Exchange Management Shell.(via Connect-ExchangeServer -serverFQDN $($idFQDN))" ;
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } ;
            } ;  # END-E
        }
    #}
    #endregion CONNECT_XOPLOCALMANAGEMENTSHELL ; #*------^ END connect-xopLocalManagementShell ^------
