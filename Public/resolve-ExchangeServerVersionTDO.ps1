# resolve-ExchangeServerVersionTDO.ps1
#*------v resolve-ExchangeServerVersionTDO.ps1 v------
function resolve-ExchangeServerVersionTDO {
    <#
    .SYNOPSIS
    resolve-ExchangeServerVersionTDO - Resolves the ExchangeVersion details from a returned get-ExchangeServer, whether local undehydrated ('Microsoft.Exchange.Data.Directory.Management.ExchangeServer') or remote EMS ('System.Management.Automation.PSObject')
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2024-05-22
    FileName    : resolve-ExchangeServerVersionTDO
    License     : (none asserted)
    Copyright   : (none asserted)
    Github      : https://github.com/tostka/verb-dev
    Tags        : Powershell,ExchangeServer,Version
    AddedCredit : Bruno Lopes (brunokktro )
    AddedWebsite: https://www.linkedin.com/in/blopesinfo
    AddedTwitter: @brunokktro / https://twitter.com/brunokktro
    REVISIONS
    * 1:22 PM 5/22/2024init
    .DESCRIPTION
    resolve-ExchangeServerVersionTDO - Resolves the ExchangeVersion details from a returned get-ExchangeServer, whether local undehydrated ('Microsoft.Exchange.Data.Directory.Management.ExchangeServer') or remote EMS ('System.Management.Automation.PSObject')
    Returns a  PSCustomObject to the pipleine with the following properties:

        isEx2019             : [boolean]
        isEx2016             : [boolean]
        isEx2007             : [boolean]
        isEx2003             : [boolean]
        isEx2000             : [boolean]
        ExVers               : [string] 'Ex2010'
        ExchangeMajorVersion : [string] '14.3'
        isEx2013             : [boolean]
        isEx2010             : [boolean]

    Extends on sample code by brunokktro's Get-ExchangeEnvironmentReport.ps1

    .PARAMETER ExchangeServer
    Object returned by a get-ExchangeServer command
    .OUTPUT
    PSCustomObject version summary.
    .EXAMPLE
    PS> write-verbose 'Resolve the local ExchangeServer object to version description, and assign to `$returned' ;     
    PS> $returned = resolve-ExchangeServerVersionTDO -ExchangeServer (get-exchangeserver $env:computername) 
    PS> write-verbose "Expand returned populated properties into local variables" ; 
    PS> $returned.psobject.properties | ?{$_.value} | %{ set-variable -Name $_.name -Value $_.value -verbose } ; 
        
        VERBOSE: Performing the operation "Set variable" on target "Name: ExVers Value: Ex2010".
        VERBOSE: Performing the operation "Set variable" on target "Name: ExchangeMajorVersion Value: 14.3".
        VERBOSE: Performing the operation "Set variable" on target "Name: isEx2010 Value: True".

    Demo retrieving get-exchangeserver, assigning to output, processing it for version info, and expanding the populated returned values to local variables. 
    .LINK
    https://github.com/brunokktro/ExchangeServer/blob/master/Get-ExchangeEnvironmentReport.ps1
    .LINK
    https://github.com/tostka/verb-Ex2010
    #>
    [CmdletBinding()]
    #[Alias('rvExVers')]
    PARAM(
        [Parameter(Mandatory = $true,Position=0,HelpMessage="Object returned by a get-ExchangeServer command[-ExchangeServer `$ExObject]")]
            [array]$ExchangeServer
    ) ;
    BEGIN {
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        $verbose = $($VerbosePreference -eq "Continue")
        $sBnr="#*======v $($CmdletName): v======" ;
        write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
    }
    PROCESS {
        foreach ($item in $ExchangeServer){
            
            if($host.version.major -ge 3){$oReport=[ordered]@{Dummy = $null ;} }
            else {$oReport = New-Object Collections.Specialized.OrderedDictionary} ;
            If($oReport.Contains("Dummy")){$oReport.remove("Dummy")} ;
            #$oReport.add('sessionid',$sid) ; # add a static pre-stocked value
            # then loop out the commonly-valued/typed entries
            
            $fieldsnull = 'isEx2019','isEx2016','isEx2007','isEx2003','isEx2000','ExVers','ExchangeMajorVersion' ; $fieldsnull | % { $oReport.add($_,$null) } ;
            #$fieldsArray = 'key4','key5','key6' ; $fieldsArray | % { $oReport.add($_,@() ) } ;
            #$oReport.add('',) ; # explicit variable
            #$oReport.add('','') ; # explicit value
            #$oReport.key1 = 'value' ;  # assign value 

            # this may be undehydrated $RemoteExchangePath, or REMS dehydrated, w fundementally diff properties
            switch($item.admindisplayversion.gettype().fullname){

                'Microsoft.Exchange.Data.ServerVersion'{
                    #    '6.0'  = @{Long = 'Exchange 2000'; Short = 'E2000' }
                         #  '6.5'  = @{Long = 'Exchange 2003'; Short = 'E2003' }
            #               '8'    = @{Long = 'Exchange 2007'; Short = 'E2007' }
            #               '14'   = @{Long = 'Exchange 2010'; Short = 'E2010' } # Ex2010 version.Minor == SP#
            #               '15'   = @{Long = 'Exchange 2013'; Short = 'E2013' } # Ex2010 version.Minor == SP#
            #               '15.1' = @{Long = 'Exchange 2016'; Short = 'E2016' } 
            #               '15.2' = @{Long = 'Exchange 2019'; Short = 'E2019' } #2019-05-17 TST Exchange Server 2019 added
                    #
                    if ($item.AdminDisplayVersion.Major -eq 6) {
                        # 6(.0) == Ex2000  ; 6.5 == Ex2003 
                        $oReport.ExchangeMajorVersion = [double]('{0}.{1}' -f $item.AdminDisplayVersion.Major, $item.AdminDisplayVersion.Minor)
                        $ExchangeSPLevel = $item.AdminDisplayVersion.FilePatchLevelDescription.Replace('Service Pack ', '')
                    } elseif ($item.AdminDisplayVersion.Major -eq 15 -and $item.AdminDisplayVersion.Minor -ge 1) {
                        # 15.1 == Ex2016 ; 15.2 == Ex2019
                        $oReport.ExchangeMajorVersion = [double]('{0}.{1}' -f $item.AdminDisplayVersion.Major, $item.AdminDisplayVersion.Minor)
                        $ExchangeSPLevel = 0
                    } else {
                        # 8(.0) == Ex2007 ; 14(.0) == Ex2010 ; 15(.0) == Ex2013 
                        $oReport.ExchangeMajorVersion = $item.AdminDisplayVersion.Major ; 
                        $ExchangeSPLevel = $item.AdminDisplayVersion.Minor ; 
                    } ; 

                    $oReport.isEx2000 = $oReport.isEx2003 = $oReport.isEx2007 = $oReport.isEx2010 = $oReport.isEx2013 = $oReport.isEx2016 = $oReport.isEx2019 = $false ; 
                    $oReport.ExVers = $null ; 
                    switch ([string]$oReport.ExchangeMajorVersion) {
                        '15.2' { $oReport.isEx2019 = $true ; $oReport.ExVers = 'Ex2019' }
                        '15.1' { $oReport.isEx2016 = $true ; $oReport.ExVers = 'Ex2016'}
                        '15' { $oReport.isEx2013 = $true ; $oReport.ExVers = 'Ex2013'}
                        '14' { $oReport.isEx2010 = $true ; $oReport.ExVers = 'Ex2010'}
                        '8' { $oReport.isEx2007 = $true ; $oReport.ExVers = 'Ex2007'}  
                        '6.5' { $oReport.isEx2003 = $true ; $oReport.ExVers = 'Ex2003'} 
                        '6' {$oReport.isEx2000 = $true ; $oReport.ExVers = 'Ex2000'} ;
                        default { 
                            $smsg = "UNRECOGNIZED ExchangeServer.AdminDisplayVersion.Major.Minor string:$($oReport.ExchangeMajorVersion)! ABORTING!" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            THROW $SMSG ; 
                            BREAK ; 
                        }
                    } ; 
                }
                'System.String'{
                    $oReport.ExVers = $oReport.isEx2000 = $oReport.isEx2003 = $oReport.isEx2007 = $oReport.isEx2010 = $oReport.isEx2013 = $oReport.isEx2016 = $oReport.isEx2019 = $false ; 
                    if([double]$ExVersNum = [regex]::match($item.AdminDisplayVersion,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
                        switch -regex ([string]$ExVersNum) {
                            '15.2' { $oReport.isEx2019 = $true ; $oReport.ExVers = 'Ex2019' }
                            '15.1' { $oReport.isEx2016 = $true ; $oReport.ExVers = 'Ex2016'}
                            '15.0' { $oReport.isEx2013 = $true ; $oReport.ExVers = 'Ex2013'}
                            '14.*' { $oReport.isEx2010 = $true ; $oReport.ExVers = 'Ex2010'}
                            '8.*' { $oReport.isEx2007 = $true ; $oReport.ExVers = 'Ex2007'}
                            '6.5' { $oReport.isEx2003 = $true ; $oReport.ExVers = 'Ex2003'}
                            '6' {$oReport.isEx2000 = $true ; $oReport.ExVers = 'Ex2000'} ;
                            default {
                                $smsg = "UNRECOGNIZED ExchangeServer.AdminDisplayVersion string:$($item.AdminDisplayVersion)! ABORTING!" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                THROW $SMSG ;
                                BREAK ;
                            }
                        } ; 
                        $smsg = "Need `$oReport.ExchangeMajorVersion as well (emulating output of non-dehydrated)" 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $oReport.ExchangeMajorVersion = $ExVersNum
                    }else {
                        $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$item.version:$($item.version)!" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        throw $smsg ; 
                        break ; 
                    } ;
                } ;
                default {
                    # $item.admindisplayversion.gettype().fullname
                    $smsg = "Unable to detect `$item.admindisplayversion.gettype():$($item.admindisplayversion.gettype().fullname)!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    throw $smsg ; 
                    break ; 
                };  
            }
            $smsg = "(returning  results for $($item.name) to pipeline)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

            New-Object -TypeName PsObject -Property $oReport | write-output ; 
        } ; 
    } # PROC-E
    END{
        write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
    }
}; 
#*------^ resolve-ExchangeServerVersionTDO.ps1 ^------
