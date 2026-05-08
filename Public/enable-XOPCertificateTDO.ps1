#region ENABLE_XOPCERTIFICATETDO ; #*------v FUNCTION enable-XOPCertificateTDO v------
function enable-XOPCertificateTDO{
            <#                                                                                                                                                                                                     <#
            .SYNOPSIS
            enable-XOPCertificateTDO.ps1 - Enable previously-imported Exchange Certificate and bind to services. 
            .NOTES
            Version     : 0.0.1
            Author      : Todd Kadrie
            Website     : http://www.toddomation.com
            Twitter     : @tostka / http://twitter.com/tostka
            CreatedDate : 2025-07-08
            FileName    : enable-XOPCertificateTDO
            License     : MIT License
            Copyright   : (c) 2025 Todd Kadrie
            Github      : https://github.com/tostka/PowershellBB
            Tags        : Powershell,Certificate,TrustChain,Import,Maintenance
            AddedCredit : REFERENCE
            AddedWebsite: URL
            AddedTwitter: URL
            REVISIONS
            * 2:35 PM 12/2/2025 spliced in re-import connect-xopLocalManagementShell code
            * 3:22 PM 7/24/2025 switched return to pscustomobj ; rem -whatif support ; logic upgrades write-myX compat ; added missing connect-xopLocalManagementShell code; added hybrid write-myX cmdlet support
            * 5:06 PM 7/12/2025 substantial rewrite, added ex-eval, version, & import detect code, spanning ex2010,16 & 19 (no SE support yet, need version numbers)
            * 5:11 PM 7/11/2025 init; wrote/added set-CertificatesInCaHierarchytdo(); flipped processing to straight loop in CA hier order, switched on extension. 

            .DESCRIPTION
            enable-XOPCertificateTDO.ps1 - Enable previously-imported Exchange Certificate and bind to services. 
            .PARAMETER CNName
            Target Certificate SubjectName[-CNName 'CN=GLOBAL.AD.TORO.COM
            .PARAMETER Computername
            Target Computer name[-Computername 'servername']
            .PARAMETER CertStoreLocation
            Destination Certificate Hive (My|LocalMachine|Both)[-CertStoreLocation 'LocalMachine']
            .PARAMETER TargetCertStore
            Optional override Destination Certificate Store name ('My':Certificates|'CA':IA's|'Root':Root CAs|'REQUEST':Pending CSRs). Default behavior is to identify Certificate type and auto-steer to the proper Store[-TargetCertStore 'CA']
            .PARAMETER SourcePath
            Optional path to a directory containing specified cert files[-SourcePath \\machine\share\]
            .PARAMETER noTranscript
            Switch to suppress internal transcript (e.g. broad ongoing transcript running)[-noTranscript]
            .PARAMETER Change
            Change Number[-Change 123456]
            .INPUTS
            None. Does not accepted piped input.(.NET types, can add description)
            .OUTPUTS
            Returns a customobject with a summary of the enabled certificate.
            .EXAMPLE
            PS> enable-XOPCertificateTDO.ps1 -CertificateFileNames 'PS_Trusted_Root_Cert.cer','ToddSelfII.pfx' -Computername $env:computername -CertStoreLocation LocalMachine -whatif
            .EXAMPLE
            PS> enable-XOPCertificateTDO.ps1 -CertificateFileNames @('new_entrust_bundle 8-20241010-0931AM.crt','mymail-toro-com-(11-10-2025).pfx') -Computername $env:computername -CertStoreLocation LocalMachine -verbose -whatif:
            TOR prod current cert install
            .EXAMPLE
            PS> enable-XOPCertificateTDO.ps1 -CertificateFileNames @('new_entrust_bundle 8-20241010-0931AM.crt','mymail-torolab-com-(10-10-2025).pfx') -Computername $env:computername -CertStoreLocation LocalMachine -verbose -whatif:
            TOL lab  current cert install
            .LINK
            https://github.com/tostka/powershellbb/
            #>
            [CmdletBinding()]
            PARAM(    
                [Parameter(Mandatory=$true,HelpMessage="Target Certificate SubjectName[-CNName 'CN=GLOBAL.AD.TORO.COM']")]
                    $CNName,
                [Parameter(Mandatory=$true,HelpMessage="Target Certificate NotBefore value[-NotBefore '7/23/2024 3:34:19 PM']")]
                    [ValidateNotNullOrEmpty()]
                    [datetime]$NotBefore,    
                [Parameter(HelpMessage="Change Number[-Change 123456]")]
                    [string] $Change,
                [Parameter(HelpMessage = "Switch to suppress internal transcript (e.g. broad ongoing transcript running)[-noTranscript]")]
                    [switch] $noTranscript
                #[Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
                #    [switch] $whatif
            )
            BEGIN{    
                if(gcm Write-MyOutput -ea 0){
                    $noTranscript = $TRUE 
                    $smsg = "Write-MyOutput detected: disabling transcribing" ;
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;
                } 
                #region BANNER ; #*------v BANNER v------
                $sBnr="#*======v $(${CmdletName}): v======" ;
                $smsg = $sBnr ;
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
                #endregion BANNER ; #*------^ END BANNER ^------

                $smsg = "`$PSBoundParameters w`n$(($PSBoundParameters|out-string).trim())" ; 
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
                if($CNName -AND $NotBefore){} else {
                    $smsg = "`$CNName ($($CNName)) -OR `$NotBefore ($($NotBefore)) are not BOTH populated!" 
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;  
                    throw $smsg ;
                    break ; 
                } ;
                
                #region CONNECT_XOPLOCAL ; #*------v connect-XopLocal v------
                $tcmdlet = 'get-exchangecertificate' ;
                $cmd = $null; $cmd = get-command $tcmdlet -erroraction 0 ;
                if(-not $cmd){
                    if($xopconn = connect-XopLocalManagementShell){
                        if($ExPSS = get-pssession | ? { $_.ComputerName -match "^$($env:computername)" -AND $_.ConfigurationName -eq 'Microsoft.Exchange' } | sort id -Descending | select -first 1 ){
                            TRY{
                                $cmd = $null; $cmd = get-command $tcmdlet -erroraction 0 ;
                                if(-not $cmd){
                                    $smsg = "Missing $($cmdlet): re-importing PSSession..." ;
                                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                    } ;
                                    $ExIPSS = Import-PSSession $ExPSS -allowclobber -ea STOP ;
                                } ;
                                $cmd = $null; $cmd = get-command 'Get-OrganizationConfig' -erroraction stop ;
                                $smsg = "Connected to: $($expss.computername)" ;
                                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                } ;
                            } CATCH {
                                $ErrTrapd=$Error[0] ;
                                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                } ;
                                BREAK ;
                            } ;
                        } ;
                    } else {
                        $smsg = "NOT CONNECTED!"
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        BREAK ;
                    } ;
                } ;
                #endregion CONNECT_XOPLOCAL ; #*------^ END connect-XopLocal ^------
                $rgxStatusKey = 'Automatic|Disabled' ; 
                $rgxXopKeySvcs = 'MSExchangeADTopology|MSExchangeFrontEndTransport|MSExchangeTransport|MSExchangeRPC|MSExchangeIS|W3SVC' ; 
                #$exlocalstatus.ExServicesStatus |?{$_.StartType -match 'Automatic|Disabled'} | ?{$_.status -eq 'Stopped'} |?{$_.servicename -match 'MSExchangeADTopology|MSExchangeFrontEndTransport|MSExchangeTransport|MSExchangeRPC|MSExchangeIS|W3SVC'}
                <#if($ExLocalStatus.ExServicesStatus |?{$_.StartType -match $rgxStatusKey} | ?{$_.status -eq 'Stopped'} |?{$_.servicename -match $rgxXopKeySvcs}){
                    $smsg = "LOCAL SERVER IS *SERVICE-DISABLED/DOWN*!" ; 
                    $smsg += "`nENABLE SERVIVCES AND BRING BACK ONLINE BEFORE RUNNING THIS SCRIPT!" ;
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    throw $smsg ; 
                    break ; 
                }
                #>
                if(-not $noTranscript){
                    $transcript = "c:\scripts\logs\$($Change)-$($env:COMPUTERNAME)-ImportExCertPFX-$(get-date -format 'yyyyMMdd-HHmmtt')" ; 
                    if(-not(test-path (split-path $transcript))){
                        $smsg = "(creating missing $(split-path $transcript))" ; 
                        if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ;
                        mkdir -path (split-path $transcript) -verbose ; 
                    }; 
                    if($whatif){
                        $transcript += "-WHATIF" ; 
                    } else { 
                        $transcript += "-EXEC" ; 
                    } ; 
                    $transcript+="-log-trans.txt" ;             
                    TRY{$stopresults = stop-transcript} CATCH {} ; 
                    TRY{$startresults = start-transcript $transcript ; 
                        $smsg = $startresults ; 
                        if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ;
                    } CATCH {
                        $smsg = "host doesn't support transcription"
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                    } ; 
                }else {
                    $smsg = "(-NoTranscript: skipping internal transcription)" ; 
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;
                }
                  
                $tcmdlet = 'get-exchangecertificate' ;
                $cmd = $null; $cmd = get-command $tcmdlet -erroraction 0 ;
                if(-not $cmd){
                    if($xopconn = connect-XopLocalManagementShell){
                        if($ExPSS = get-pssession | ? { $_.ComputerName -match "^$($env:computername)" -AND $_.ConfigurationName -eq 'Microsoft.Exchange' } | sort id -Descending | select -first 1 ){
                            TRY{
                                $cmd = $null; $cmd = get-command $tcmdlet -erroraction 0 ;
                                if(-not $cmd){
                                    $smsg = "Missing $($cmdlet): re-importing PSSession..." ;
                                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                    } ;
                                    $ExIPSS = Import-PSSession $ExPSS -allowclobber -ea STOP ;
                                } ;
                                $cmd = $null; $cmd = get-command 'Get-OrganizationConfig' -erroraction stop ;
                                $smsg = "Connected to: $($expss.computername)" ;
                                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                } ;
                            } CATCH {
                                $ErrTrapd=$Error[0] ;
                                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                } ;
                                BREAK ;
                            } ;
                        } ;
                    } else {
                        $smsg = "NOT CONNECTED!"
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        BREAK ;
                    } ;
                } ;              
            };  # BEG-E
            PROCESS{
                TRY{             
                    $cert=get-childitem cert:\LocalMachine\My\ -EA STOP|?{$_.Subject -like $CNName -AND $_.NotBefore -ge (get-date $NotBefore)} ;
                    #$cert | select subject,friendlyname,not*,DNSName,pspath ;
                    $smsg = "Matched Cert:`n$(($cert | select subject,friendlyname,not*,DNSName,pspath|out-string).trim())" ; 
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;
                    if($cert -isnot [system.array]){
                        if((get-exchangeserver $env:computername -EA STOP).IsClientAccessServer ){
                            $smsg = "The next command will prompt:" ;
                            $smsg += "`nConfirm Overwrite the existing default SMTP certificate?" ;
                            $smsg += "`nCurrent certificate: [thumb will match the *OLD* cert]" ;
                            $smsg += "`nReplace it with certificate: [new cert thumb]?`n" ;
                            $smsg += "ANSWER Y(YES) OVERWRITE!`n" ;                            
                            if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                        }elseif((get-exchangeserver $env:computername -EA STOP).IsEdgeServer){
                            $smsg = "The next command will prompt:" ;
                            $smsg += "`nConfirm Overwrite the existing default SMTP certificate?" ;
                            $smsg += "`nCurrent certificate: [thumb will match the *SELF-SIGNED* cert that runs edgsync and mail transfer]" ;
                            $smsg += "`nReplace it with certificate: [new cert thumb]?`n" ;
                            $smsg += "ANSWER N (NO!) OVERWRITE! TO AVOID DAMAGING SYNC & TRANSFER!`n" ;                            
                            if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                        }  ;
                        
                        TRY{
                            Enable-ExchangeCertificate -server $env:computername -Thumbprint $($cert.Thumbprint) -Services POP,IMAP,SMTP,IIS -EA STOP -whatif:$($whatif) ;
                        } CATCH {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                            if(gcm Write-MyError -ea 0){Write-MyError $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error}
                                else{ write-ERROR "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                        } ; 
                        
                        $smsg = "Confirming:" ;       
                        if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ;
                        TRY{                 
                            $bcert=get-exchangecertificate -Server $env:COMPUTERNAME -ErrorAction STOP |
                                    ?{($_.Services -match "IIS") -AND ($_.Services -match "SMTP") -AND ($_.Services -match "IMAP") -AND ($_.Services -match "POP")}  ;
                            $retObj = [ordered]@{
                                Subject    =  $bcert.Subject
                                Services   =  $bcert.Services
                                NotAfter   =    $bcert.NotAfter
                                NotBefore  =    $bcert.NotBefore
                                Thumbprint =    $bcert.Thumbprint
                            } ; 
                            $smsg = "`n=MATCHED CERT ON $($env:COMPUTERNAME):`n$(($bcert | fl Subject,Services,not*,thumb*|out-string).trim())" ;
                            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            } ;
                            if($bcert -isnot [system.array]){
                                $tCERT=gci cert:\LocalMachine\My\$($bcert.Thumbprint) ;
                                $retObj.add('SANS',"`n$((($tCert.Extensions | Where-Object {$_.Oid.FriendlyName -eq "subject alternative name"}).Format(1)|out-string).trim())")
                                $smsg = "==$($tcert.Subject) SANS" ;
                                $smsg += "`n$((($tCert.Extensions | Where-Object {$_.Oid.FriendlyName -eq "subject alternative name"}).Format(1)|out-string).trim())" ;
                            } else {
                                $smsg = "multiple certs returned! $($bcert.Subject| out-string)" 
                                if(gcm Write-MyError -ea 0){Write-MyError $smsg } else {
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error}
                                    else{ write-ERROR "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                } ;                                
                            };
                        } CATCH {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                            if(gcm Write-MyError -ea 0){Write-MyError $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error}
                                else{ write-ERROR "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                        } ; 
                        $smsg = "Returning certificate summary to pipeline" ;
                        if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                            if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        } ;
                        [pscustomobject]$retObj | write-output ;
                    } else {
                         $smsg = "multiple certs matched!"  ; 
                         if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;  
                        throw $smsg ;
                    } ;
                } CATCH {     $ErrTrapd=$Error[0] ;
                    $smsg = "`n$(($ErrTrapd  | fl * -Force|out-string).trim())" ;
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    Break ;
                } ;
   
            } ;  # PROC-E
            END{
                if(-not $noTranscript){
                    TRY{
                        $smsg = $stopresults = stop-transcript 
                        if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                    } CATCH {
                        $smsg = "host doesn't support transcription"
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                    } ; 
                } ; 
                $smsg = "$($sBnr.replace('=v','=^').replace('v=','^='))" ;
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
            };  # END-E   
        }
#endregion ENABLE_XOPCERTIFICATETDO ; #*------^ END FUNCTION enable-XOPCertificateTDO  ^------

# SIG # Begin signature block
# MIIELgYJKoZIhvcNAQcCoIIEHzCCBBsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUw260BRWU1wHA0o7ktacqtid3
# RUqgggI4MIICNDCCAaGgAwIBAgIQWsnStFUuSIVNR8uhNSlE6TAJBgUrDgMCHQUA
# MCwxKjAoBgNVBAMTIVBvd2VyU2hlbGwgTG9jYWwgQ2VydGlmaWNhdGUgUm9vdDAe
# Fw0xNDEyMjkxNzA3MzNaFw0zOTEyMzEyMzU5NTlaMBUxEzARBgNVBAMTClRvZGRT
# ZWxmSUkwgZ8wDQYJKoZIhvcNAQEBBQADgY0AMIGJAoGBALqRVt7uNweTkZZ+16QG
# a+NnFYNRPPa8Bnm071ohGe27jNWKPVUbDfd0OY2sqCBQCEFVb5pqcIECRRnlhN5H
# +EEJmm2x9AU0uS7IHxHeUo8fkW4vm49adkat5gAoOZOwbuNntBOAJy9LCyNs4F1I
# KKphP3TyDwe8XqsEVwB2m9FPAgMBAAGjdjB0MBMGA1UdJQQMMAoGCCsGAQUFBwMD
# MF0GA1UdAQRWMFSAEL95r+Rh65kgqZl+tgchMuKhLjAsMSowKAYDVQQDEyFQb3dl
# clNoZWxsIExvY2FsIENlcnRpZmljYXRlIFJvb3SCEGwiXbeZNci7Rxiz/r43gVsw
# CQYFKw4DAh0FAAOBgQB6ECSnXHUs7/bCr6Z556K6IDJNWsccjcV89fHA/zKMX0w0
# 6NefCtxas/QHUA9mS87HRHLzKjFqweA3BnQ5lr5mPDlho8U90Nvtpj58G9I5SPUg
# CspNr5jEHOL5EdJFBIv3zI2jQ8TPbFGC0Cz72+4oYzSxWpftNX41MmEsZkMaADGC
# AWAwggFcAgEBMEAwLDEqMCgGA1UEAxMhUG93ZXJTaGVsbCBMb2NhbCBDZXJ0aWZp
# Y2F0ZSBSb290AhBaydK0VS5IhU1Hy6E1KUTpMAkGBSsOAwIaBQCgeDAYBgorBgEE
# AYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwG
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQ8Bjr7
# 21tm3SkAuQ6SgxyPhzMR5zANBgkqhkiG9w0BAQEFAASBgC+1Cb64ITIO0W1/aHVK
# kXGkrMSILes+p5OwTgpReEoWj+MLNorlHL5B/BPEd2onDLif2d9i6EPR4CiONpob
# t0u/rHe2ZHOcLO9dch/UUcFm1CXH+OdqL/YPIFJhEfTqRbPbe75QLk4tENi0Kt+0
# EZqRzqW2yPcWTbX8afZWNto/
# SIG # End signature block
