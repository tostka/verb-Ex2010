# Connect-ExchangeServerTDO.ps1

#region CONNECT_EXCHANGESERVERTDO ; #*------v Connect-ExchangeServerTDO v------
#if(-not(gci function:Connect-ExchangeServerTDO -ea 0)){
    Function Connect-ExchangeServerTDO {
        <#
        .SYNOPSIS
        Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
        will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellREmote (REMS) connect to each server, 
        stopping at the first successful connection.
        .NOTES
        REVISIONS
        * 3:58 PM 5/14/2025 restored prior dropped earlier rev history (routinely trim for psparamt inclu)
        * 10;07 am 4/30/2025 fixed borked edge conn, typo, and rev logic for Ex & role detection in raw PS - lacks evaris for exchange (EMS/REMS only), so leverage reg & stock install loc hunting to discover setup.exe for vers & role confirm).
        * 2:46 PM 4/22/2025 add: -Version (default to Ex2010), and postfiltered returned ExchangeServers on version. If no -Version, sort on newest Version, then name, -descending.
        * 4:25 PM 1/15/2025 seems to work at this point, move to rebuild
        * 4:49 PM 1/9/2025 reworked connect-exchangeserverTdo() to actually use the credentials passed in, and 
        added the missing import-module $PSS, to _connect-ExOP, to make the session actually functional 
        for running cmds, wo popping cred prompts. 
        * 12:24 PM 12/4/2024 removed bracket bnr echos around _connect-ExOP
        * 3:54 PM 11/26/2024 integrated back TLS fixes, and ExVersNum flip from June; syncd dbg & vx10 copies.
        * 12:57 PM 6/11/2024 Validated, Ex2010 & Ex2019, hub, mail & edge roles: tested ☑️ on CMW mail role (Curly); and Jumpbox; 
        copied in CBH from repo copy, which has been updated/debugged compat on CMW Edge 
        includes local snapin detect & load for edge role (simplest EMS load option for Edge role, from David Paulson's original code; no longer published with Ex2010 compat)
        * 1:30 PM 9/5/2024 added  update-SecurityProtocolTDO() SB to begin
        * 12:49 PM 6/21/2024 flipped PSS Name to Exchange$($ExchVers[dd])
        * 11:28 AM 5/30/2024 fixed failure to recognize existing functional PSSession; Made substantial update in logic, validate works fine with other orgs, and in our local orgs.
        * 4:02 PM 8/28/2023 debuged, updated CBH, renamed connect-ExchangeSErver -> Connect-ExchangeServerTDO (avoid name clashes, pretty common verb-noun combo).
        * 12:36 PM 8/24/2023 init
         .PARAMETER name
        FQDN of a specific Exchange server[-Name EXSERVER.DOMAIN.COM]
        .PARAMETER discover
        Boolean paraameter that drives auto-discovery of target Exchange servers for connection (defaults `$true)[-discover:`$false]
        .PARAMETER credential
        Use specific Credentials[-Credentials [credential object]
            .PARAMETER Site
        Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-Site 'SITENAME']
        .PARAMETER RoleNames
        Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']
        .PARAMETER Version
        Specific Exchange Server Version to connect to('Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000')[-Version 'Ex2016']
        .PARAMETER TenOrg
        Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
        .INPUTS
        None. Does not accepted piped input.(.NET types, can add description)
        .OUTPUTS
        [system.object] Returns a system object containing a successful PSSession
        System.Boolean
        .EXAMPLE
        PS> $PSSession = Connect-ExchangeServerTDO -siteName SITENAME -RoleNames @('HUB','CAS') -verbose 
        Demo's connecting to a functional Hub or CAS server in the SITENAME site with verbose outputs, the `PSSession variable will contain information about the successful connection. Makes automatic Exchangeserver discovery calls into AD (using ADSI) leveraging the separate get-ADExchangeServerTDO()
        .EXAMPLE
        PS> TRY{$Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name}CATCH{$Site=$env:COMPUTERNAME} ;
        PS> $PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
        Demo including support for EdgeRole, which is detected on it's lack of AD Site specification (which gets fed through to call, by setting the Site to the machine itself).
        .EXAMPLE
        PS> $PSSession = Connect-ExchangeServerTDO -siteName SITENAME -RoleNames @('HUB','CAS') -Version Ex2016 -verbose 
        Demo's connecting to a functional Hub or CAS server Version Ex2016 in the SITENAME site with verbose outputs, the `PSSession variable will contain information about the successful connection. Makes automatic Exchangeserver discovery calls into AD (using ADSI) leveraging the separate get-ADExchangeServerTDO()
        #>        
        [CmdletBinding(DefaultParameterSetName='discover')]
        PARAM(
            [Parameter(Position=0,Mandatory=$true,ParameterSetName='name',HelpMessage="FQDN of a specific Exchange server[-Name EXSERVER.DOMAIN.COM]")]
                [String]$name,
            [Parameter(Position=0,ParameterSetName='discover',HelpMessage="Boolean paraameter that drives auto-discovery of target Exchange servers for connection (defaults `$true)[-discover:`$false]")]
                [bool]$discover=$true,
            [Parameter(Position=1,HelpMessage = "Use specific Credentials[-Credentials [credential object]")]
                [Management.Automation.PSCredential]$credential,
            [Parameter(Position=1,HelpMessage="Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-Site 'SITENAME']")]
                [Alias('Site')]
                [string]$SiteName = (gv -name "$($TenOrg)_ADSiteDefaultName" -ea 0).Value,
            [Parameter(Position=2,HelpMessage="Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']")]
                [ValidateSet('MBX','CAS','HUB','UM','MBX','EDGE')]
                [string[]]$RoleNames = @('HUB','CAS'),
            [Parameter(Position=2,HelpMessage="Specific Exchange Server Version to connect to('Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000')[-Version 'Ex2016']")]
                [ValidateSet('Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000')]
                [string[]]$Version = 'Ex2010',
            [Parameter(Mandatory=$FALSE,HelpMessage="Tenant Tag (3-letter abbrebiation - defaults to global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']")]
                [ValidateNotNullOrEmpty()]
                [string]$TenOrg = $global:o365_TenOrgDefault
        ) ;
        BEGIN{
            $Verbose = ($VerbosePreference -eq 'Continue') ;
            $CurrentVersionTlsLabel = [Net.ServicePointManager]::SecurityProtocol ; # Tls, Tls11, Tls12 ('Tls' == TLS1.0)  ;
      write-verbose "PRE: `$CurrentVersionTlsLabel : $($CurrentVersionTlsLabel )" ;
      # psv6+ already covers, test via the SslProtocol parameter presense
      if ('SslProtocol' -notin (Get-Command Invoke-RestMethod).Parameters.Keys) {
        $currentMaxTlsValue = [Math]::Max([Net.ServicePointManager]::SecurityProtocol.value__,[Net.SecurityProtocolType]::Tls.value__) ;
        write-verbose "`$currentMaxTlsValue : $($currentMaxTlsValue )" ;
        $newerTlsTypeEnums = [enum]::GetValues('Net.SecurityProtocolType') | Where-Object { $_ -gt $currentMaxTlsValue }
        if($newerTlsTypeEnums){
          write-verbose "Appending upgraded/missing TLS `$enums:`n$(($newerTlsTypeEnums -join ','|out-string).trim())" ;
        } else {
          write-verbose "Current TLS `$enums are up to date with max rev available on this machine" ;
        };
        $newerTlsTypeEnums | ForEach-Object {
          [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $_
        } ;
      } ;
    
            # 5:15 PM 4/22/2025 on CMW, have to patch version to Ex2016

            #*------v Function _connect-ExOP v------
            function _connect-ExOP{
                    [CmdletBinding()]
                    PARAM(
                        [Parameter(Position=0,Mandatory=$true,HelpMessage="Exchange server AD Summary system object[-Server EXSERVER.DOMAIN.COM]")]
                            [system.object]$Server,
                        [Parameter(Position=1,HelpMessage = "Use specific Credentials[-Credentials [credential object]")]
                            [Management.Automation.PSCredential]$credential
                    );
                    $verbose = $($VerbosePreference -eq "Continue") ;
                    if([double]$ExVersNum = [regex]::match($Server.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
                        switch -regex ([string]$ExVersNum) {
                            '15.2' { $isEx2019 = $true ; $ExVers = 'Ex2019' }
                            '15.1' { $isEx2016 = $true ; $ExVers = 'Ex2016'}
                            '15.0' { $isEx2013 = $true ; $ExVers = 'Ex2013'}
                            '14.*' { $isEx2010 = $true ; $ExVers = 'Ex2010'}
                            '8.*' { $isEx2007 = $true ; $ExVers = 'Ex2007'}
                            '6.5' { $isEx2003 = $true ; $ExVers = 'Ex2003'}
                            '6' {$isEx2000 = $true ; $ExVers = 'Ex2000'} ;
                            default {
                                $smsg = "UNRECOGNIZED ExVersNum.Major.Minor string:$($ExVersNum)! ABORTING!" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                THROW $SMSG ;
                                BREAK ;
                            }
                        } ;
                    }else {
                        $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$Server.version:$($Server.version)!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        throw $smsg ;
                        break ;
                    } ;
                    if($Server.RoleNames -eq 'EDGE'){
                        if(($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup')) -or
                            ($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup')) -or
                            $ByPassLocalExchangeServerTest)
                        {
                            if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole') -or
                                    (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'))
                            {
                                $smsg = "We are on Exchange Edge Transport Server"
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                $IsEdgeTransport = $true
                            }
                            TRY {
                                Get-ExchangeServer -ErrorAction Stop | Out-Null
                                $smsg = "Exchange PowerShell Module already loaded."
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                $passed = $true 
                            }CATCH {
                                $smsg = "Failed to run Get-ExchangeServer"
                                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                if($isLocalExchangeServer){
                                    write-host  "Loading Exchange PowerShell Module..."
                                    TRY{
                                        if($IsEdgeTransport){
                                            # implement local snapins access on edge role: Only way to get access to EMS commands.
                                            [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exshell.psc1" -ErrorAction Stop
                                            ForEach($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn){
                                                write-verbose ("Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name)
                                                Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop
                                            } ; 
                                            Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop ; 
                                            $passed = $true #We are just going to assume this passed.
                                        }else{
                                            Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                                            Connect-ExchangeServer -Auto -ClientApplication:ManagementShell
                                            $passed = $true #We are just going to assume this passed.
                                        } 
                                    }CATCH {
                                        $smsg = "Failed to Load Exchange PowerShell Module..." ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                    }                               
                                } ;
                            } FINALLY {
                                if($LoadExchangeVariables -and $passed -and $isLocalExchangeServer){
                                    if($ExInstall -eq $null -or $ExBin -eq $null){
                                        if(Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup'){
                                            $Global:ExInstall = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath
                                        }else{
                                            $Global:ExInstall = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath
                                        }

                                        $Global:ExBin = $Global:ExInstall + "\Bin"

                                        $smsg = ("Set ExInstall: {0}" -f $Global:ExInstall)
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                        $smsg = ("Set ExBin: {0}" -f $Global:ExBin)
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                    } ; 
                                } ; 
                            } ; 
                        } else  {
                            $smsg = "Does not appear to be an Exchange 2010 or newer server." ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                        }
                        if(get-command -Name Get-OrganizationConfig -ea 0){
                            $smsg = "Running in connected/Native EMS" ; 
                            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            Return $true ; 
                        } else { 
                            TRY{
                                $smsg = "Initiating Edge EMS local session (exshell.psc1 & exchange.ps1)" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                # 5;36 PM 5/30/2024 didn't work, went off to nowhere for a long time, and exited the script
                                #& (gcm powershell.exe).path -PSConsoleFile "$($env:ExchangeInstallPath)bin\exshell.psc1" -noexit -command ". '$($env:ExchangeInstallPath)bin\Exchange.ps1'"
                                <# [Adding the Transport Server to Exchange - Mark Lewis Blog](https://marklewis.blog/2020/11/19/adding-the-transport-server-to-exchange/)
                                To access the management console on the transport server, I opened PowerShell then ran
                                exshell.psc1
                                Followed by
                                exchange.ps1
                                At this point, I was able to create a new subscription using he following PowerShel
                                #>
                                invoke-command exshell.psc1 ; 
                                invoke-command exchange.ps1
                                if(get-command -Name Get-OrganizationConfig -ea 0){
                                    $smsg = "Running in connected/Native EMS" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                    Return $true ;
                                } else { return $false };  
                            }CATCH{
                                Write-Error $_ ;
                            } ;
                        } ; 
                    } else {
                        $pltNPSS=@{ConnectionURI="http://$($Server.FQDN)/powershell"; ConfigurationName='Microsoft.Exchange' ; name="Exchange$($ExVersNum.tostring())"} ;
                        $pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ;} ;
                        # use ExVersUnm dd instead of hardcoded (Exchange2010)
                        if($ExVersNum -ge 15){
                            $smsg = "EXOP.15+:Adding -Authentication Kerberos" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            $pltNPSS.add('Authentication',"Kerberos") ;
                            $pltNPSS.name = $ExVers ;
                        } ;
                        $smsg = "Adding EMS (connecting to $($Server.FQDN))..." ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $smsg = "New-PSSession w`n$(($pltNPSS|out-string).trim())" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $ExPSS = New-PSSession @pltNPSS  ;
                        $ExIPSS = Import-PSSession $ExPSS -allowclobber ;
                        # 3:59 PM 1/9/2025 appears credprompting is due to it's missing the import-module $ExIPSS ! 
                        $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $Global:E10Mod = Import-Module $ExIPSS @pltIMod ;
                        $ExPSS | write-output ;
                        $ExPSS= $ExIPSS = $null ;
                    } ; 
                } ;
            #*------^ END Function _connect-ExOP ^------
            $pltGADX=@{
                ErrorAction='Stop';
            } ;
        } ;
        PROCESS{
            if($PSBoundParameters.ContainsKey('credential')){
                $pltGADX.Add('credential',$credential) ;
            }
            if($SiteName){
                $pltGADX.Add('siteName',$siteName) ;
            } ;
            if($RoleNames){
                $pltGADX.Add('RoleNames',$RoleNames) ;
            } ;
            TRY{
                if($discover){
                    $smsg = "Getting list of Exchange Servers" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    $exchServers=get-ADExchangeServerTDO @pltGADX | sort responsetime ;
                }else{
                    $exchServers=get-ADExchangeServerTDO @pltGADX | sort responsetime ;
                } ;
                $pltTW=@{
                    'ErrorAction'='Stop';
                } ;
                $pltCXOP = @{
                    verbose = $($VerbosePreference -eq "Continue") ;
                } ;
                if($pltGADX.credential){
                    $pltCXOP.Add('Credential',$pltGADX.credential) ;
                } ;
                if($Version){
                    switch ($Version){
                      'Ex2000'{$rgxExVersNum = '6' } 
                      'Ex2003'{$rgxExVersNum = '6.5' } 
                      'Ex2007'{$rgxExVersNum = '8.*' } 
                      'Ex2010'{$rgxExVersNum = '14.*'} 
                      'Ex2013'{$rgxExVersNum = '15.0' } 
                      'Ex2016'{$rgxExVersNum = '15.1'} 
                      'Ex2019'{$rgxExVersNum = '15.2' } 
                    } ; 
                    $exchServers  = $exchServers | ?{ [double]([regex]::match( $_.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value) -match $rgxExVersNum } ; 

                } else {
                    write-verbose "no -Version: Sorting Newest first, then names, descending" ; 
                    $exchServers  = $exchServers | sort version,name -desc
                } ; 
                $prpPSS = 'Id','Name','ComputerName','ComputerType','State','ConfigurationName','Availability' ; 
                foreach($exServer in $exchServers){
                    $smsg = "testing conn to:$($exServer.name.tostring())..." ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig' -ea SilentlyContinue){
                        if($pssEXOP = Get-PSSession |  where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ( $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND ( $_.runspace.ConnectionInfo.Port -eq '80') }){
                            if($pssEXOP.State -ne "Opened" -OR $pssEXOP.Availability -ne "Available"){
                                $pssEXOP | remove-pssession ; $pssEXOP = $null ;
                            } ;
                        } ; 
                    } else {
                        $smsg = "(mangled ExOP conn: disconnect/reconnect...)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        if($pssEXOP = Get-PSSession |  where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ( $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND ( $_.runspace.ConnectionInfo.Port -eq '80') }){
                            if($pssEXOP.State -ne "Opened" -OR $pssEXOP.Availability -ne "Available"){
                                $pssEXOP | remove-pssession ; $pssEXOP = $null ;
                            } ;
                        } ; 
                    } ;
                    if(-not $pssEXOP){
                        $smsg = "Connecting to: $($exServer.FQDN)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $smsg = "_connect-ExOP w`n$(($pltCXOP|out-string).trim())" ;
                        $smsg += "`nServer $($exServer.FQDN)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        if($NoTest){
                            $ExPSS =$ExPSS = _connect-ExOP @pltCXOP -Server $exServer
                        } else {
                            TRY{
                                $smsg = "Testing Connection: $($exServer.FQDN)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                If(test-connection $exServer.FQDN -count 1 -ea 0) {
                                    $smsg = "confirmed pingable..." ;
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                } else {
                                    $smsg = "Unable to Ping $($exServer.FQDN)" ; ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                } ;
                                $smsg = "Testing WinRm: $($exServer.FQDN)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                $winrm=Test-WSMan @pltTW -ComputerName $exServer.FQDN ;
                                if($winrm){
                                    $ExPSS = _connect-ExOP @pltCXOP -Server $exServer;
                                } else {
                                    $smsg = "Unable to Test-WSMan $($exServer.FQDN) (skipping)" ; ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                } ;
                            }CATCH{
                                $errMsg="Server: $($exServer.FQDN)] $($_.Exception.Message)" ;
                                Write-Error -Message $errMsg ;
                                continue ;
                            } ;
                        };
                    } else {
                        $smsg = "$((get-date).ToString('HH:mm:ss')):Accepting first valid connection w`n$(($pssEXOP | ft -a $prpPSS|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $ExPSS = $pssEXOP ; 
                        break ; 
                    }  ;
                } ;
            }CATCH{
                Write-Error $_ ;
            } ;
        } ;
        END{
            if(-not $ExPSS){
                $smsg = "NO SUCCESSFUL CONNECTION WAS MADE, WITH THE SPECIFIED INPUTS!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $smsg = "(returning `$false to the pipeline...)" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                return $false
            } else{
                if($ExPSS.State -eq "Opened" -AND $ExPSS.Availability -eq "Available"){
                    if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                        $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ;
                        throw $smsg ;
                        $smsg | write-warning  ;
                    } else {
                        $smsg = "(connected to EXOP.Org:$($orgName))" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    return $ExPSS
                } ;
            } ; 
        } ;
    } ;
#} ; 
#endregion CONNECT_EXCHANGESERVERTDO ; #*------^ END Connect-ExchangeServerTDO ^------