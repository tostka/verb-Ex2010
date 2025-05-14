# test-LocalExchangeInfoTDO.ps1

#region TEST_EXCHANGEINFO ; #*------v test-LocalExchangeInfoTDO v------
#if(-not (get-item function:test-LocalExchangeInfoTDO -ea 0)){
    function test-LocalExchangeInfoTDO {
        
        <#
        .SYNOPSIS
        test-LocalExchangeInfoTDO.ps1 - Checks local environment for evidence of a local Exchangeserver install, the version installed, and wether Edge Role. Returns a summary object to the pipeline.
        .NOTES
        Version     : 1.0.0
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2025-05-09
        FileName    : test-LocalExchangeInfoTDO.ps1
        License     : MIT License
        Copyright   : (c) 2025 Todd Kadrie
        Github      : https://github.com/tostka/verb-Network
        Tags        : Powershell
        AddedCredit : Fabian Bader
        AddedWebsite: https://cloudbrothers.info/en/
        AddedTwitter: 
        REVISION
        * 12:53 PM 5/13/2025 swaped w-h -> w-v
        * 10:10 AM 5/12/2025 added -whatif:$false -confirm:$false to nested set-variable cmds - SSP prevents set-vari updates, just like any other action verb cmdlet.
        * 1:44 PM 5/9/2025 init
        .DESCRIPTION
        test-LocalExchangeInfoTDO.ps1 - Checks local environment for evidence of a local Exchangeserver install, the version installed, and wether Edge Role. Returns a summary object to the pipeline.
            
        Returns a psCustomObject summarizing the environment findings, in re: local Exchange server fingerprints:

            Name                           Value
            ----                           -----
            isLocalExchangeServer          True
            IsEdgeTransport                False
            ExVers                         Ex2010
            isEx2019                       False
            isEx2016                       False
            isEx2013                       False
            isEx2010                       True
            isEx2007                       False
            isEx2003                       False
            isEx2000                       False

        .INPUTS
        Does not accept piped input
        .OUTPUTS
        System.Management.Automation.PSCustomObject environment summary object with following properties:

            isLocalExchangeServer          [boolean]
            IsEdgeTransport                [boolean]
            ExVers                         [version]
            isEx2019                       [boolean]
            isEx2016                       [boolean]
            isEx2013                       [boolean]
            isEx2010                       [boolean]
            isEx2007                       [boolean]
            isEx2003                       [boolean]
            isEx2000                       [boolean] 

        .EXAMPLE
        PS> $results = test-LocalExchangeInfoTDO ; 
        PS> $results ; 

            Name                           Value
            ----                           -----
            isLocalExchangeServer          True
            IsEdgeTransport                False
            ExVers                         Ex2010
            isEx2019                       False
            isEx2016                       False
            isEx2013                       False
            isEx2010                       True
            isEx2007                       False
            isEx2003                       False
            isEx2000                       False
        
        PS> write-verbose "Expand returned NoteProperty properties into matching local variables" ; 
        PS> if($host.version.major -gt 2){
        PS>     $results.PsObject.Properties | ?{$_.membertype -eq 'NoteProperty'} | %{set-variable -name $_.name -value $_.value -verbose -whatif:$false -Confirm:$false ;} ;
        PS> }else{
        PS>     write-verbose "Psv2 lacks the above expansion capability; just create simpler variable set" ; 
        PS>     $ExVers = $results.ExVers ; $isLocalExchangeServer = $results.isLocalExchangeServer ; $IsEdgeTransport = $results.IsEdgeTransport ;
        PS> } ;

            VERBOSE: Performing the operation "Set variable" on target "Name: isEx2003 Value: False".
            VERBOSE: Performing the operation "Set variable" on target "Name: isEx2013 Value: False".
            VERBOSE: Performing the operation "Set variable" on target "Name: isEx2010 Value: False".
            VERBOSE: Performing the operation "Set variable" on target "Name: isEx2019 Value: False".
            VERBOSE: Performing the operation "Set variable" on target "Name: isEx2000 Value: False".
            VERBOSE: Performing the operation "Set variable" on target "Name: IsEdgeTransport Value: False".
            VERBOSE: Performing the operation "Set variable" on target "Name: isEx2016 Value: True".
            VERBOSE: Performing the operation "Set variable" on target "Name: isLocalExchangeServer Value: True".
            VERBOSE: Performing the operation "Set variable" on target "Name: isEx2007 Value: False".
            VERBOSE: Performing the operation "Set variable" on target "Name: ExVers Value: Ex2016".

        Demo pass with follow-on expansion of return pscustomobject into matching individual variables (or, on PSv2, which support for syntax above, expansion , the simpler $ExVers, $isLocalExchangeserver & $IsEdgeTransport variables).
        .LINK
        https://github.com/tostka/verb-Ex2010
        #>
        [CmdletBinding()]
        #[Alias('')]
        PARAM ()    
        PROCESS {
            #$isLocalExchangeServer = $IsEdgeTransport = $isEx2019 =  $isEx2016 =  $isEx2013 =  $isEx2010 =  $isEx2007 =  $isEx2003 =  $isEx2000 = $false ; 
            if($host.version.major -ge 3){$hSummary=[ordered]@{Dummy = $null ;} }
            else {$hSummary = $hSummary = @{Dummy = $null ;} } ;
            if($hSummary.keys -contains 'dummy'){$hSummary.remove('Dummy') };
            $fieldsBoolean = 'isLocalExchangeServer','IsEdgeTransport','isEx2019','isEx2016','isEx2010','isEx2007','isEx2003','isEx2000' | sort ; $fieldsBoolean | % { $hSummary.add($_,$false) } ;
            $fieldsnull = 'ExVers'  | sort ; $fieldsnull | % { $hSummary.add($_,$null) } ;
            <# creates equiv to hashtable:
            $hSummary = @{
                isLocalExchangeServer = $false ; 
                IsEdgeTransport = $false ; 
                ExVers = $null ;  
                isEx2019 = $false ; 
                isEx2016 = $false ; 
                isEx2013 = $false ; 
                isEx2010 = $false ; 
                isEx2007 = $false ; 
                isEx2003 = $false ; 
                isEx2000 = $false ; 
            } ; 
            #>
            if($env:ExchangeInstalled){
                $hSummary.isLocalExchangeServer = $true ;
            } elseif((get-service MSEx* -ea 0) -AND  ($hklmPath = (resolve-path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v1*\Setup").path)){
                $hSummary.isLocalExchangeServer = $true ;
                switch -regex ($hklmPath){
                    '\\v14\\'{$isEx2010 = $true ; $hSummary.ExVers = 'Ex2010' ; write-verbose "Ex2010" ; }
                    '\\v15\\'{write-verbose "\v115\Setup == Ex2016/Ex2019"}
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
            if($hSummary.isLocalExchangeServer){
                if((get-service MSExchangeEdgeCredential -ea 0) -AND (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v1*\EdgeTransportRole')){$hSummary.IsEdgeTransport = $true} ;
                if($vers = (get-item "$($env:ExchangeInstallPath)\Bin\Setup.exe" -ea 0).VersionInfo.FileVersionRaw ){} else {
                    if($binPath = (resolve-path  "$($env:ProgramFiles)\Microsoft\Exchange Server\V1*\Bin\Setup.exe" -ea 0).path){ } else {
                        (get-psdrive -PSProvider FileSystem |?{$_ -match '[D-Z]'}  | select -expand name)|foreach-object{
                            $drv = $_ ;
                            if($rp = resolve-path  "$($drv)$($env:ProgramFiles.substring(1,($env:ProgramFiles.length-1)))\Microsoft\Exchange Server\V1*\Bin\Setup.exe" -ea 0){
                                $binPath = $rp.path;
                                if($host.version.major -gt 2){break} else {write-verbose "PSv2 breaks entire script w break, instead of branching out of local loop" } ;
                            } ;
                        };
                    } ;
                    if($binPath){
                        if( ($vers = (get-item $binPath).VersionInfo.FileVersionRaw) -OR ($vers = (get-item $binPath).VersionInfo.FileVersion) ){
                        }else {
                            $smsg = "Unable to manually resolve an `$env:ExchangeInstallPath equiv, on any local drive" ;
                            write-warning $smsg ;
                            throw $smsg ;
                        }
                    } ;
                } ;
            } ;
            if($hSummary.isLocalExchangeServer){
                if($vers){
                    switch -regex ($vers){
                        '15\.2' { $hSummary.isEx2019 = $true ; $hSummary.ExVers = 'Ex2019' }
                        '15\.1' { $hSummary.isEx2016 = $true ; $hSummary.ExVers = 'Ex2016'}
                        '15\.0' { $hSummary.isEx2013 = $true ; $hSummary.ExVers = 'Ex2013'}
                        '14\..*' { $hSummary.isEx2010 = $true ; $hSummary.ExVers = 'Ex2010'}
                        '8\..*' { $hSummary.isEx2007 = $true ; $hSummary.ExVers = 'Ex2007'}
                        '6\.5' { $hSummary.isEx2003 = $true ; $hSummary.ExVers = 'Ex2003'}
                        '6|6\.0' {$hSummary.isEx2000 = $true ; $hSummary.ExVers = 'Ex2000'} ;
                        default{ throw "[$($vers.tostring())]: Unrecognized version!" } ;
                    } ;
                }else {
                    throw "Empty `$vers resolved ExchangeVersion string variable!"
                } ; 
                $smsg = @("`$hSummary.ExVers: $($hSummary.ExVers)") ;
                $smsg += @("`$$((gv "is$($hSummary.ExVers)" -ea 0).name): $((gv "is$($hSummary.ExVers)"  -ea 0).value)") ;
                if($hSummary.IsEdgeTransport){ $smsg += @("`$hSummary.IsEdgeTransport: $($hSummary.IsEdgeTransport)") } else { $smsg += @(" (non-Edge)")} ;
                $smsg = ($smsg -join ' ') ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } else {
                $smsg = "(non-Local ExchangeServer (`$hSummary.isLocalExchangeServer:$([boolean]$hSummary.isLocalExchangeServer )))" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } ;
        } ; 
        END {
            [pscustomobject]$hSummary | write-output 
        }
    } ;
#} ; 
#endregion TEST_EXCHANGEINFO ; #*------^ END test-LocalExchangeInfoTDO ^------