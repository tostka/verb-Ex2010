# test-LocalExchangeInfoTDO.ps1

#region TEST_LOCALEXCHANGEINFOTDO ; #*------v test-LocalExchangeInfoTDO v------
#if (-not (get-item function:test-LocalExchangeInfoTDO -ea 0)) {
    function test-LocalExchangeInfoTDO {
    <#
    .SYNOPSIS
    test-LocalExchangeInfoTDO - Checks local server's status as an Exchange Server (checks for Exchange Services, Registry Keys, key roles, versions), without reliance on Exchange Mgmt Shell). Differs from vx10\get-xopServerAdminDisplayVersion(), in that it isn't intended to be run for remote server version verification, and avoids reliance on get-exchangeserver and other Exchange Mgmt Shell dependancies.
    .NOTES
    Version     : 0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 20250711-0423PM
    FileName    : test-LocalExchangeInfoTDO.ps1
    License     : MIT License
    Copyright   : (c) 2025 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell,Exchange,ExchangeServer,Install,Patch,Maintenance
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 2:58 PM 7/17/2025 updated CBH;  hybrid with prexisting vx10\test-LocalExchangeInfoTDO, combined best ideas from both; ren (again) to match existing: test-xopExchangeLocalInstallTDO -> test-LocalExchangeInfoTDO()
    * 4:09 PM 7/13/2025 add: cbh demo to test for down/disabled svcs state ;
     ren: test-xopExchangeLocalInstallTDO -> test-xopExchangeLocalInstallTDO, aliased orig ; add: all useful roles props 'isAdminTools','isCAS','isUM','isAdminTools','isCAS','isUM','isHub','isEdgeTransport' (tested in registry)
    * * 5:29 PM 7/12/2025 init; added support for version detect of Exchange Subcription Edition (identified as ExVers: ExSE, 15.2.2562+ (only differentiation from Ex2019 is that that vers is still 15.2.1748...)
    .DESCRIPTION
    test-LocalExchangeInfoTDO - Checks local server's status as an Exchange Server (checks for Exchange Services, Registry Keys, key roles, versions), without reliance on Exchange Mgmt Shell). Differs from vx10\get-xopServerAdminDisplayVersion(), in that it isn't intended to be run for remote server version verification, and avoids reliance on get-exchangeserver and other Exchange Mgmt Shell dependancies.

    Has the following potential properties that may be returned (only returns those populated/relevent to the local system):

    hasExServices = [boolean] ;
    ExServicesStatus = [msex & w3svc & clussvc services status] ;
    isLocalExchangeServer = [boolean] ;
    isAdminTools = [boolean] ;
    isCAS = [boolean] ;
    isUM = [boolean] ;
    isMbx = [boolean] ;
    isHub = [boolean] ;
    isEdgeTransport = [boolean] ;
    hasRoleWatermark = [boolean] ;
    isExSE = [boolean]  # Exchange Subscription Edition identifier.
    isEx2019 = [boolean]
    isEx2016 = [boolean]
    isEx2013 = [boolean]
    isEx2010 = [boolean]
    isEx2007 = [boolean]
    isEx2003 = [boolean]
    isEx2000 = [boolean]
    ExVers = [string]  'ExS','Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000'

    ## return on a typical Exchange 2016 Mailbox server (with services stopped/disabled)
    
    ```powershell
    isLocalExchangeServer : True
    hasExServices         : True
    ExServicesStatus      : {@{ServiceName=MSExchangeADTopology; DisplayName=Microsoft Exchange Active Directory Topology; Status=Stopped; StartType=Disabled}, @{ServiceName=MSExchangeAntispamUpdate; 
                            DisplayName=Microsoft Exchange Anti-spam Update; Status=Stopped; StartType=Automatic}, @{ServiceName=MSExchangeCompliance; DisplayName=Microsoft Exchange Compliance Service; 
                            Status=Stopped; StartType=Automatic}, @{ServiceName=MSExchangeDagMgmt; DisplayName=Microsoft Exchange DAG Management; Status=Stopped; StartType=Automatic}...}
    ExVers                : Ex2016
    isAdminTools          : True
    isCAS                 : True
    isEx2016              : True
    isHub                 : True
    isMbx                 : True
    isUM                  : True

     
    ```
    .INPUTS
    None, no piped input.
    .OUTPUTS
    System.Object summary of Exchange server descriptors, and service statuses.
    .EXAMPLE
    PS> $ExLocalStatus = test-LocalExchangeInfoTDO ;
    PS> $ExLocalStatus ;

        isLocalExchangeServer : True
        hasExServices         : True
        ExServicesStatus      : {@{ServiceName=MSExchangeADTopology; DisplayName=Microsoft Exchange Active Directory Topology; Status=Stopped; StartType=Disabled}, @{ServiceName=MSExchangeAntispamUpdate; 
                                DisplayName=Microsoft Exchange Anti-spam Update; Status=Stopped; StartType=Automatic}, @{ServiceName=MSExchangeCompliance; DisplayName=Microsoft Exchange Compliance Service; 
                                Status=Stopped; StartType=Automatic}, @{ServiceName=MSExchangeDagMgmt; DisplayName=Microsoft Exchange DAG Management; Status=Stopped; StartType=Automatic}...}
        ExVers                : Ex2016
        isAdminTools          : True
        isCAS                 : True
        isEx2016              : True
        isHub                 : True
        isMbx                 : True
        isUM                  : True

    Typical Exchange 2016 return information (Mailbox role server, with services stopped & disabled)
    .EXAMPLE
    PS> $rgxStatusKey = 'Automatic|Disabled' ;
    PS> $rgxXopKeySvcs = 'MSExchangeADTopology|MSExchangeFrontEndTransport|MSExchangeTransport|MSExchangeRPC|MSExchangeIS|MSExchangeEdgeCredential|W3SVC' ;
    PS> $ExLocalStatus = test-LocalExchangeInfoTDO ;
    PS> $exlocalstatus.ExServicesStatus |?{$_.StartType -match 'Automatic|Disabled'} | ?{$_.status -eq 'Stopped'} |?{$_.servicename -match 'MSExchangeADTopology|MSExchangeFrontEndTransport|MSExchangeTransport|MSExchangeRPC|MSExchangeIS|W3SVC'}
    PS> if($ExLocalStatus.ExServicesStatus |?{$_.StartType -match $rgxStatusKey} | ?{$_.status -eq 'Stopped'} |?{$_.servicename -match $rgxXopKeySvcs}){
    PS>     $smsg = "LOCAL SERVER IS *SERVICE-DISABLED/DOWN*!" ;
    PS>     $smsg += "`nENABLE SERVIVCES AND BRING BACK ONLINE BEFORE RUNNING THIS SCRIPT!" ;
    PS>     WRITE-WARNING $smsg ;
    PS>     throw $smsg ;
    PS>     break ;
    PS> } ;
    Demo testing returned status for running key service state.
    .LINK
    https://github.org/tostka/verb-ex2010/
    #>
        [alias('get-xopExchangeLocalVersionTDO', 'get-xopExchangeLocalVersion')]
        PARAM(
            [Parameter(HelpMessage = "Switch to force Watermark test status visible (silent unless fail, otherwise)")]
            [switch]$showWatermark
        ) ;
        BEGIN {
            $rgxExSvcNames = '^MSEx'
            $rgxExSvcNamesFull = '^(MSEx|W3SVC|ClusSvc)'
            $RegistryPath = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*" ;
            $rgxRoleSuBkEYS = '(AdminTools|ClientAccessRole|UnifiedMessagingRole|MailboxRole|FrontendTransportRole|CafeRole|EdgeTransportRole)$' ;
            $ValueNameToCheck = "Watermark" ; #
        }
        PROCESS {
            #$isLocalExchangeServer = $IsEdgeTransport = $isEx2019 =  $isEx2016 =  $isEx2013 =  $isEx2010 =  $isEx2007 =  $isEx2003 =  $isEx2000 = $false ;
            if ($host.version.major -ge 3) { $hSummary = [ordered]@{Dummy = $null ; } }
            else { $hSummary = @{Dummy = $null ; } } ;
            if ($hSummary.keys -contains 'dummy') { $hSummary.remove('Dummy') };
            $fieldsBoolean = 'isLocalExchangeServer', 'hasExServices' ; $fieldsBoolean | % { $hSummary.add($_, $false) } ;
            $fieldsnull = 'ExServicesStatus', 'isAdminTools', 'isCAS', 'isUM', 'isMbx', 'isHub', 'isEdgeTransport', 'hasRoleWatermark', 'isExSE', 'isEx2019', 'isEx2016', 'isEx2013', 'isEx2010', 'isEx2007', 'isEx2003', 'isEx2000', 'ExVers'  | sort ; $fieldsnull | % { $hSummary.add($_, $null) } ;
            <# creates equiv to hashtable:
            $hSummary=[ordered]@{
                hasExServices = $false ;
                ExServicesStatus = $null ;
                isLocalExchangeServer = $false ;
                isAdminTools = $null ;
                isCAS = $null ;
                isUM = $null ;
                isMbx = $null ;
                isHub = $null ;
                isEdgeTransport = $null ;
                hasRoleWatermark = $null ;
                isExSE = $null ;
                isEx2019 = $null ;
                isEx2016 = $null ;
                isEx2013 = $null ;
                isEx2010 = $null ;
                isEx2007 = $null ;
                isEx2003 = $null ;
                isEx2000 = $null ;
                ExVers = $null ;
            }
            #>
            if (get-service | ? { $_.ServiceName -match $rgxExSvcNames }) {
                $hSummary.hasExServices = $true ;
                $hSummary.ExServicesStatus = get-service | ? { $_.ServiceName -match $rgxExSvcNamesFull } | select-object servicename, displayname, status, starttype
            } else { $hSummary.hasExServices = $false } ;
            if ($env:ExchangeInstalled) {
                $hSummary.isLocalExchangeServer = $true ;
            } elseif ($hSummary.hasExServices -AND ($hklmPath = (resolve-path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\Setup").path)) {
                $hSummary.isLocalExchangeServer = $true ;
                switch -regex ($hklmPath) {
                    '\\v14\\' { $isEx2010 = $true ; $hSummary.ExVers = 'Ex2010' ; write-verbose "Ex2010" ; }
                    '\\v15\\' { write-verbose "\v115\Setup == Ex2016/Ex2019" }
                    default {
                        $smsg = "Unable to manually resolve $($hklmPath) to a known version path!" ;
                        write-warning $smsg ;
                        throw $smsg ;
                    }
                } ;
            } else {
                write-verbose "hSummary.isLocalExchangeServer:$false" ;
                $hSummary.isLocalExchangeServer = $false ;
            } ;
            if ($hSummary.isLocalExchangeServer) {
                if (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\AdminTools') {
                    $smsg = "Local Installed:AdminTools"
                    write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
                    $hSummary.isAdminTools = $true
                }
                if (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\ClientAccessRole') {
                    $smsg = "Local Installed:ClientAccessRole"
                    write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)"
                    $hSummary.isCAS = $true
                }
                if (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\UnifiedMessagingRole') {
                    $smsg = "Local Installed:UnifiedMessagingRole"
                    write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                    $hSummary.isUM = $true
                }
                if (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\MailboxRole') {
                    $smsg = "Local Installed:MailboxRole"
                    write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)";
                    $hSummary.isMbx = $true
                }
                if (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\FrontendTransportRole') {
                    $smsg = "Local Installed:FrontendTransportRole"
                    write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                    $hSummary.isHub = $true
                }
                if ((get-service MSExchangeEdgeCredential -ea 0) -AND (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\EdgeTransportRole')) {
                    $smsg = "Local Installed:EdgeTransportRole"
                    write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                    $hSummary.isEdgeTransport = $true
                } ;
            } ;
            # version detect:
            if ($hSummary.isLocalExchangeServer) {
                write-verbose "Checking local discovered Exsetup.exe FileversionInfo" ;
                if ($FileversionInfo = Get-Command Exsetup.exe | ForEach-Object { $_.FileversionInfo } ) {
                    write-verbose "`$FileversionInfo:Exsetup.exe`n$(($FileversionInfo | ft -a |out-string).trim())" ;
                    [version]$ExsetupRev = (@($FileversionInfo.FileMajorPart, $FileversionInfo.FileMinorPart, $FileversionInfo.FileBuildPart, $FileversionInfo.FilePrivatePart) -join '.')
                    #$ExsetupProduct = $BuildToProductName[$ExsetupRev.tostring()]
                    #write-verbose "`$ExsetupProduct:$($ExsetupProduct)" ;
                } elseif ($FileversionInfo = (get-item "$($env:ExchangeInstallPath)\Bin\Setup.exe" -ea 0).VersionInfo.FileVersionRaw ) {
                    write-verbose "`$FileversionInfo:Setup.exe`n$(($FileversionInfo | ft -a |out-string).trim())" ;
                    [version]$ExsetupRev = (@($FileversionInfo.FileMajorPart, $FileversionInfo.FileMinorPart, $FileversionInfo.FileBuildPart, $FileversionInfo.FilePrivatePart) -join '.')
                } else {
                    throw "$($Server.name):Unable to remote retrieve: Get-Command Exsetup.exe | ForEach-Object { $_.FileversionInfo}"
                } ;
                switch -regex ([string](@($FileversionInfo.FileMajorPart, $FileversionInfo.FileMinorPart) -join '.')) {
                    '15\.2' {
                        # SE only diffs from ex2019, in the 15.2.25*+ vs 15.2.221-1748
                        if ($FileversionInfo.FileBuildPart -ge 2562) {
                            $hSummary.isExSE = $true ; $hSummary.ExVers = 'ExSE'
                        } else {
                            $hSummary.isEx2019 = $true ; $hSummary.ExVers = 'Ex2019'
                        } ;
                    }
                    '15\.1' { $hSummary.isEx2016 = $true ; $hSummary.ExVers = 'Ex2016' }
                    '15\.0' { $hSummary.isEx2013 = $true ; $hSummary.ExVers = 'Ex2013' }
                    '14\..*' { $hSummary.isEx2010 = $true ; $hSummary.ExVers = 'Ex2010' }
                    '8\..*' { $hSummary.isEx2007 = $true ; $hSummary.ExVers = 'Ex2007' }
                    '6\.5' { $hSummary.isEx2003 = $true ; $hSummary.ExVers = 'Ex2003' }
                    '6|6\.0' { $hSummary.isEx2000 = $true ; $hSummary.ExVers = 'Ex2000' } ;
                    default {
                        $smsg = "UNRECOGNIZED ExVersNum.Major.Minor string:$(@($FileversionInfo.FileMajorPart,$FileversionInfo.FileMinorPart) -join '.')! ABORTING!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent }
                        else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        THROW $SMSG ;
                        BREAK ;
                    }
                } ;
            } ;
            <# this tests the version on the binpath folder
            if ($hSummary.isLocalExchangeServer) {
                if ($vers = (get-item "$($env:ExchangeInstallPath)\Bin\Setup.exe" -ea 0).VersionInfo.FileVersionRaw ) {} else {
                    if ($binPath = (resolve-path  "$($env:ProgramFiles)\Microsoft\Exchange Server\V1*\Bin\Setup.exe" -ea 0).path) { } else {
                        (get-psdrive -PSProvider FileSystem | ? { $_ -match '[D-Z]' }  | select -expand name) | foreach-object {
                            $drv = $_ ;
                            if ($rp = resolve-path  "$($drv)$($env:ProgramFiles.substring(1,($env:ProgramFiles.length-1)))\Microsoft\Exchange Server\V1*\Bin\Setup.exe" -ea 0) {
                                $binPath = $rp.path;
                                if ($host.version.major -gt 2) { break } else { write-verbose "PSv2 breaks entire script w break, instead of branching out of local loop" } ;
                            } ;
                        };
                    } ;
                    if ($binPath) {
                        if ( ($vers = (get-item $binPath).VersionInfo.FileVersionRaw) -OR ($vers = (get-item $binPath).VersionInfo.FileVersion) ) {
                        } else {
                            $smsg = "Unable to manually resolve an `$env:ExchangeInstallPath equiv, on any local drive" ;
                            write-warning $smsg ;
                            throw $smsg ;
                        }
                    } ;
                } ;
            } ;
            #>
            # add a watermark test
            if ($hSummary.isLocalExchangeServer) {
                if($wmarks = test-xopExchangeInstallWatermarkTDO){
                    $smsg = "FOUND INSTALL WATERMARKS! w`n$(($wmarks|out-string).trim())" ;
                    WRITE-WARNING  "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                } else {
                    $smsg = "No Watermark values found in `$RegistryPath" ;
                    if ($showWatermark) {
                        write-host -foregroundcolor green $smsg ;
                    } else {
                        write-verbose $smsg ;
                    } ;
                } ;
            } ;
            if ($hSummary.isLocalExchangeServer) {
                $smsg = @("`$hSummary.ExVers: $($hSummary.ExVers)") ;
                $smsg += @("`$$((gv "is$($hSummary.ExVers)" -ea 0).name): $((gv "is$($hSummary.ExVers)"  -ea 0).value)") ;
                if ($hSummary.IsEdgeTransport) { $smsg += @("`$hSummary.IsEdgeTransport: $($hSummary.IsEdgeTransport)") } else { $smsg += @(" (non-Edge)") } ;
                $smsg = ($smsg -join ' ') ;
                write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
            } else {
                $smsg = "(non-Local ExchangeServer (`$hSummary.isLocalExchangeServer:$([boolean]$hSummary.isLocalExchangeServer )))" ;
                if ($verbose) {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                    else { write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            } ;
        } ;
        END {
            # clean out null/empty value props (return only populated props)
            $mts = $hSummary.GetEnumerator() | ? { ($null -eq $_.value) -OR ($_.value -eq '') } ; $mts | foreach-object { $hSummary.remove($_.Name) } ; remove-variable mts -ea 0 ;
            [pscustomobject]$hSummary | write-output
        }
    } ;
#} ;
#endregion TEST_LOCALEXCHANGEINFOTDO ; #*------^ END test-LocalExchangeInfoTDO ^------