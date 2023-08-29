#*------v Function Connect-ExchangeServerTDO v------
#if(-not (get-command Connect-ExchangeServerTDO -ea 0)){
    Function Connect-ExchangeServerTDO {
        <#
            .SYNOPSIS
            Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
            will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellREmote (REMS) connect to each server, 
            stopping at the first successful connection.
            .NOTES
            Version     : 3.0.0
            Author      : Todd Kadrie
            Website     : http://www.toddomation.com
            Twitter     : @tostka / http://twitter.com/tostka
            CreatedDate : 2015-09-03
            FileName    : Connect-ExchangeServerTDO.ps1
            License     : (none-asserted)
            Copyright   : (none-asserted)
            Github      : https://github.com/tostka/verb-Ex2010
            Tags        : Powershell, ActiveDirectory, Exchange, Discovery
            AddedCredit : Brian Farnsworth
            AddedWebsite: https://codeandkeep.com/
            AddedTwitter: URL
            REVISIONS
            * 4:02 PM 8/28/2023 debuged, updated CBH, renamed connect-ExchangeSErver -> Connect-ExchangeServerTDO (avoid name clashes, pretty common verb-noun combo).
            * 12:36 PM 8/24/2023 init

            .DESCRIPTION
            Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
            will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellREmote (REMS) connect to each server, 
            stopping at the first successful connection.

            Relies upon/requires get-ADExchangeServerTDO(), to return a descriptive summary of the Exchange server(s) revision etc, for connectivity logic.
            Supports Exchange 2010 through 2019, as implemented.
        
            Intent, as contrasted with verb-EXOP/Ex2010 is to have no local module dependancies, when running EXOP into other connected orgs, where syncing profile & supporting modules code can be problematic. 
            This uses native ADSI calls, which are supported by Windows itself, without need for external ActiveDirectory module etc.

            The particular approach inspired by BF's demo func that accompanied his take on get-adExchangeServer(), which I hybrided with my own existing code for cred-less connectivity. 
            I added get-OrganizationConfig testing, for connection pre/post confirmation, along with Exchange Server revision code for continutional handling of new-pssession remote powershell EMS connections.
            Also shifted connection code into _connect-EXOP() internal func.
            As this doesn't rely on local module presnece, it doesn't have to do the usual local remote/local invocation detection you'd do for non-dehydrated on-server EMS (more consistent this way, anyway; 
            there are only a few cmdlet outputs I'm aware of, that have fundementally broken returns dehydrated, and require local non-remote EMS use to function.

            My core usage would be to paste the function into the BEGIN{} block for a given remote org process, to function as a stricly local ad-hoc function.
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
            .PARAMETER TenOrg
            Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
            .INPUTS
            None. Does not accepted piped input.(.NET types, can add description)
            .OUTPUTS
            [system.object] Returns a system object containing a successful PSSession
            System.Boolean
            [| get-member the output to see what .NET obj TypeName is returned, to use here]
            System.Array of System.Object's
            .EXAMPLE
            PS> $PSSession = Connect-ExchangeServerTDO -siteName SITENAME -RoleNames @('HUB','CAS') -verbose 
            Demo's connecting to a functional Hub or CAS server in the SITENAME site with verbose outputs, the `PSSession variable will contain information about the successful connection.
            .LINK
            https://github.com/tostka/verb-XXX
            .LINK
            https://bitbucket.org/tostka/powershell/
            .LINK
            https://github.com/tostka/verb-Ex2010
            .LINK
            https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/
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
            [Parameter(Mandatory=$FALSE,HelpMessage="Tenant Tag (3-letter abbrebiation - defaults to global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']")]
                [ValidateNotNullOrEmpty()]
                #[ValidatePattern("^\w{3}$")]
                [string]$TenOrg = $global:o365_TenOrgDefault
        ) ; 
        BEGIN{
            $Verbose = ($VerbosePreference -eq 'Continue') ;
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
                $pltNPSS=@{ConnectionURI="http://$($Server.FQDN)/powershell"; ConfigurationName='Microsoft.Exchange' ; name='Exchange2010'} ;  
                if($ExVersNum -ge 15){
                    write-verbose "EXOP.15+:Adding -Authentication Kerberos" ; 
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
                $ExIPSS = Import-PSSession $ExPSS -allowclobber # -CommandName get-transportserver ;
                $ExPSS | write-output ; 
                $ExPSS= $ExIPSS = $null ; 
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
            #if($PSBoundParameters.ContainsKey('siteName')){
            # only hits if explicit param, if default the value, you need to test for value, not param
            if($SiteName){
                $pltGADX.Add('siteName',$siteName) ; 
            } ; 
            #RoleNames
            #if($PSBoundParameters.ContainsKey('RoleNames')){
            # only hits if explicit param, if default the value, you need to test for value, not param
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
                    $pltCXOP.Add('Credential',$pltCXOP.Credential) ; 
                } ; 
            
                foreach($exServer in $exchServers){
                    if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig' -ea SilentlyContinue){} else {
                        $smsg = "(mangled ExOP conn: disconnect/reconnect...)" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $pssEXOP = Get-PSSession |  where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ( $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND ( $_.runspace.ConnectionInfo.Port -eq '80') } ; 
                        if($pssEXOP.State -ne "Opened" -OR $pssEXOP.Availability -ne "Available"){ 
                            $pssEXOP | remove-pssession ; $pssEXOP = $null ; 
                        } ; 
                    } ; 
                    if(-not $pssEXOP){    
                        $smsg = "Connecting to: $($exServer.FQDN)" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
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
                                    #$ExPSS = _connect-ExOP -Server $exServer -verbose:$($VerbosePreference -eq "Continue") ; 
                                    $ExPSS = _connect-ExOP @pltCXOP -Server $exServer;
                                
                                } else {
                                    $smsg = "Unable to Test-WSMan $($exServer.FQDN) (skipping)" ; ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                } ; 
                            }CATCH{
                                $errMsg="Server: $($exServer.FQDN)] $($_.Exception.Message)" ; 
                                Write-Error -Message $errMsg ; 
                                continue ; 
                            } ; 
                        };
                        if($ExPSS.State -eq "Opened" -AND $ExPSS.Availability -eq "Available"){ 
                            if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                                $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ; 
                                throw $smsg ; 
                                $smsg | write-warning  ; 
                            } else {
                                $smsg = "(connected to EXOP.Org:$($orgName))" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            } ; 
                            return $ExPSS
                        } ; 
                    } ; 
                } ; 
            }CATCH{
                Write-Error $_ ; 
            } ; 
        } ;  # PROC-E
      END{
            if(-not $ExPSS){
                $smsg = "NO SUCCESSFUL CONNECTION WAS MADE, WITH THE SPECIFIED INPUTS!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "(returning `$false to the pipeline...)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                return $false 
            } 
      } ; 
    } ; 
#} ; 
#*------^ END Function Connect-ExchangeServerTDO ^------