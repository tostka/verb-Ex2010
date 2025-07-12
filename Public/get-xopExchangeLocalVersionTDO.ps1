# get-xopExchangeLocalVersionTDO.ps1
#region get_xopExchangeLocalVersionTDO ; #*------v get-xopExchangeLocalVersionTDO v------
function get-xopExchangeLocalVersionTDO{
    <#
    .SYNOPSIS
    get-xopExchangeLocalVersionTDO - Checks local server's status as an Exchange Server (checks for Exchange Services, Registry Keys, key roles, versions), without reliance on Exchange Mgmt Shell). Differs from vx10\get-xopServerAdminDisplayVersion(), in that it isn't intended to be run for remotely server version verification, and avoids reliance on get-exchangeserver and other Exchange Mgmt Shell dependancies.
    .NOTES
    Version     : 0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 20250711-0423PM
    FileName    : get-xopExchangeLocalVersionTDO.ps1
    License     : MIT License
    Copyright   : (c) 2025 Todd Kadrie
    Github      : https://github.com/tostka/Network
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * * 5:29 PM 7/12/2025 init; added support for version detect of Exchange Subcription Edition (identified as ExVers: ExSE, 15.2.2562+ (only differentiation from Ex2019 is that that vers is still 15.2.1748...)
    .DESCRIPTION
    get-xopExchangeLocalVersionTDO - Checks local server's status as an Exchange Server (checks for Exchange Services, Registry Keys, key roles, versions), without reliance on Exchange Mgmt Shell). Differs from vx10\get-xopServerAdminDisplayVersion(), in that it isn't intended to be run for remotely server version verification, and avoids reliance on get-exchangeserver and other Exchange Mgmt Shell dependancies.     

    Has the following potential properties that may be returned (only returns those populated/relevent to the local system):

    hasExServices = [boolean] ;
    ExServicesStatus = [msex & w3svc & clussvc services status] ; 
    isLocalExchangeServer = [boolean] ; 
    isEdgeTransport = [boolean]
    isExSE = [boolean]  # Exchange Subscription Edition identifier.
    isEx2019 = [boolean]
    isEx2016 = [boolean]
    isEx2013 = [boolean]
    isEx2010 = [boolean]
    isEx2007 = [boolean]
    isEx2003 = [boolean]
    isEx2000 = [boolean]
    ExVers = [string] 'ExSE|Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000'

    .INPUTS
    None, no piped input.
    .OUTPUTS
    System.Object summary of Exchange server descriptors, and service statuses.
    .EXAMPLE
    PS> $ExLocalStatus = get-xopExchangeLocalVersionTDO ; 
    PS> $ExLocalStatus ; 

        $ExLocalStatus
        hasExServices         : True
        ExServicesStatus      : {@{ServiceName=MSExchangeADTopology; DisplayName=Microsoft Exchange Active Directory Topology; Status=Stopped; StartType=Disabled}, @{ServiceName=MSExchangeAntispamUpdate;
                                DisplayName=Microsoft Exchange Anti-spam Update; Status=Stopped; StartType=Automatic}, @{ServiceName=MSExchangeCompliance; DisplayName=Microsoft Exchange Compliance Service;
                                Status=Stopped; StartType=Automatic}, @{ServiceName=MSExchangeDagMgmt; DisplayName=Microsoft Exchange DAG Management; Status=Stopped; StartType=Automatic}...}
        isLocalExchangeServer : True
        isEx2016              : True
        ExVers                : Ex2016

    Typical Exchange 2016 return information
    .LINK
    https://github.org/tostka/powershellBB/
    #>
    [CmdletBinding()]
    [alias('get_xopExchangeLocalVersion')]
    PARAM() ;
    BEGIN{
        $rgxExSvcNames = '^MSEx'
        $rgxExSvcNamesFull = '^(MSEx|W3SVC|ClusSvc)'
    }
    PROCESS{
        $oRpt=[ordered]@{
            hasExServices = $false ;
            ExServicesStatus = $null ; 
            isLocalExchangeServer = $false ; 
            isEdgeTransport = $false
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
        # test for XOP environment
        if(get-service | ?{$_.ServiceName -match $rgxExSvcNames}){
            $oRpt.hasExServices = $true ;                
            #$oRpt.ExServicesStatus = get-service |?{$_.ServiceName -match $rgxExSvcNamesFull} | ft -a servicename,displayname,status,starttype
            $oRpt.ExServicesStatus = get-service |?{$_.ServiceName -match $rgxExSvcNamesFull} | select-object servicename,displayname,status,starttype
        } else {$oRpt.hasExServices = $false } ; 
        if(($oRpt.isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup')) -or
                ($oRpt.isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup')))
        {
        
            if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole') -or
                    (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'))
            {
                $smsg = "We are on Exchange Edge Transport Server"
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $oRpt.isEdgeTransport = $true
            }
        } ; 
        if($oRpt.isLocalExchangeServer){
            if($FileversionInfo = Get-Command Exsetup.exe | ForEach-Object { $_.FileversionInfo } ){
                write-verbose "`$FileversionInfo:`n$(($FileversionInfo | ft -a |out-string).trim())" ;
                [version]$ExsetupRev = (@($FileversionInfo.FileMajorPart,$FileversionInfo.FileMinorPart,$FileversionInfo.FileBuildPart,$FileversionInfo.FilePrivatePart) -join '.')
                #$ExsetupProduct = $BuildToProductName[$ExsetupRev.tostring()]
                #write-verbose "`$ExsetupProduct:$($ExsetupProduct)" ;
            } else {
                throw "$($Server.name):Unable to remote retrieve: Get-Command Exsetup.exe | ForEach-Object { $_.FileversionInfo}"
            } ; 
            switch -regex ([string](@($FileversionInfo.FileMajorPart,$FileversionInfo.FileMinorPart) -join '.')) {
                
                '15.2' { 
                    # SE only diffs from ex2019, in the 15.2.25*+ vs 15.2.221-1748
                    if($FileversionInfo.FileBuildPart -ge 2562){
                        $oRpt.isExSE = $true ; $oRpt.ExVers = 'ExSE' 
                    }else{
                        $oRpt.isEx2019 = $true ; $oRpt.ExVers = 'Ex2019' 
                    } ; 
                }
                '15.1' { $oRpt.isEx2016 = $true ; $oRpt.ExVers = 'Ex2016'}
                '15.0' { $oRpt.isEx2013 = $true ; $oRpt.ExVers = 'Ex2013'}
                '14.*' { $oRpt.isEx2010 = $true ; $oRpt.ExVers = 'Ex2010'}
                '8.*' { $oRpt.isEx2007 = $true ; $oRpt.ExVers = 'Ex2007'}
                '6.5' { $oRpt.isEx2003 = $true ; $oRpt.ExVers = 'Ex2003'}
                '6' {$oRpt.isEx2000 = $true ; $oRpt.ExVers = 'Ex2000'} ;
                default {
                    $smsg = "UNRECOGNIZED ExVersNum.Major.Minor string:$(@($FileversionInfo.FileMajorPart,$FileversionInfo.FileMinorPart) -join '.')! ABORTING!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    THROW $SMSG ;
                    BREAK ;
                }
            } ;        
        } ;            
    }
    END{
        # echo status variables
        #get-variable |?{$_.name -match '^(isEx|hasEx|isLocalExchange|IsEdgeTransport|ExVers)'} | ft -a | Out-Default | write-host -foregroundcolor yellow  ;             
        $mts = $oRpt.GetEnumerator() |?{ ($_.value -eq $null) -OR ($_.value -eq '')} ; $mts |foreach-object{$oRpt.remove($_.Name)} ; remove-variable mts -ea 0 ; 
        [pscustomobject]$oRpt | write-output  ; 
        # issue: although you can evaluate the $true & string values, you can't enumerate and expand the ExServicesStatus
        # try reurning the raw hash instead of cobj
        #$oRpt | write-output  ; 
        # ack! issue was that I had an ft -a in the grab
    }
} ; 
#endregion get_xopExchangeLocalVersionTDO ; #*------^ END get-xopExchangeLocalVersionTDO ^------