#region IMPORT_XOPCERTIFICATEPFXTDO ; #*------v FUNCTION import-xopCertificatePfxTDO v------
function import-xopCertificatePfxTDO{
        <#
        .SYNOPSIS
        import-xopCertificatePfxTDO - Imports Exchange certificates - has logic to auto-locate & find per UserDomain. Returns certificate summary properties to pipeline
        .NOTES
        Version     : 0.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2025-07-24
        FileName    : import-xopCertificatePfxTDO.ps1
        License     : MIT License
        Copyright   : (c) 2025 Todd Kadrie
        Github      : https://github.com/tostka/verb-XXX
        Tags        : Powershell,Certificate,TrustChain,Import,Maintenance,ExchangeServer,Exchange
        AddedCredit : REFERENCE
        AddedWebsite: URL
        AddedTwitter: URL
        REVISIONS
        * 2:11 PM 8/12/2025 added trycatch on the import attempt, trying to capture the TOL 01T rpc error that left certs non returnable via get-exchangecertificate.
        * 2:47 PM 7/24/2025 getting an unresolvable alias on orig name: ren import-xopCertificateTDO -> import-xopCertificatePfxTDO 
        updated path resolution logic ren import-xopCertificate -> import-xopCertificateTDO, alias orig
        * 11:30 AM 7/11/2025 init
        .DESCRIPTION
        import-xopCertificatePfxTDO - Imports Exchange certificates - has logic to auto-locate & find per UserDomain
        .PARAMETER Computername
        Target Computer name[-Computername 'servername']"
        .PARAMETER CertStoreLocation
        "Destination Certificate Hive (CurrentUser|LocalMachine|Both)[-CertStoreLocation 'LocalMachine']
        .PARAMETER TargetCertStore
        Destination Certificate Store name ('My':Certificates|'CA':IA's|'Root':Root CAs|'REQUEST':Pending CSRs)[-TargetCertStore 'CA']
        .PARAMETER Path
        "Path to a target Certificate .cer/.crt file [-Path \\machine\share\]
        .PARAMETER whatIf
        Whatif Flag  [-whatIf]
        .INPUTS
        None. Does not accepted piped input.
        .OUTPUTS
        System.Security.Cryptography.X509Certificates.X509Certificate
        .EXAMPLE
        PS> $results = import-xopCertificatePfxTDO -whatif -verbose
        PS> $results 

            Subject      : CN=SUB.DOMAIN.com, O=The Toro Company, L=Bloomington, S=Minnesota, C=US
            Services     :
            NotAfter     : 11/10/2025 7:54:40 AM
            NotBefore    : 10/10/2024 8:54:41 AM
            Thumbprint   : DF6F6ADD7A304FB8A894C679C240FB91B6843ED4
            FriendlyName :

        Run with whatif & verbose
        .EXAMPLE
        PS> $pltLegCER=@{
        PS>     computername = $env:computername ;
        PS>     CertStoreLocation = 'LocalMachine' ;
        PS>     TargetCertStore = 'Root' ;
        PS>     Path = 'c:\path-to\cert.cer' ;
        PS>     whatIf = $($whatif)
        PS> } ;
        PS> $smsg += "`nusing import-xopCertificatePfxTDO`n$(($pltLegCER|out-string).trim())" ;
        PS> write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
        PS> $ret = import-xopCertificatePfxTDO @pltLegCER ;
        Typical splatted call
        .EXAMPLE
        PS> $result = import-xopCertificatePfxTDO -computername $env:computername -CertStoreLocation 'LocalMachine' -TargetCertStore Root -Path d:\cab\new_entrust_bundle 8-20241010-0931AM.crt -whatIf:$true ;
        Parameter call        
        .LINK
        https://github.com/tostka/verb-Network
        #>
        [CmdletBinding()]
        #[Alias('import-xopCertificate')]
        Param (
            [Parameter(Mandatory=$false,HelpMessage="Path to target PFX file")]
                [string]$FilePath,
            [Parameter(Mandatory=$false,HelpMessage="Securestring Password for PFX format certificates[-Password `$securestringpw")]
                [System.Security.SecureString]$Password,
            [Parameter(Mandatory=$false,HelpMessage="Exchange version string (ExSE|Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000)[-ExVers Ex2016]")]
                [ValidateSet('ExSE','Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000')]
                [string]$ExVers = 'Ex2016'
        )
        BEGIN{
            if(test-path $FilePath -PathType Container -ea 0){
                # fed a folder it tries to discover for the domain in the folder                
                switch ($env:USERDOMAIN){
                    'TORO'{
                        $FilePath = get-childitem -path "$($FilePath)\*.pfx" -ea STOP |?{$_.name -match 'mymail-toro-com-'} | select -expand fullname ;
                    }
                    'TORO-LAB'{
                        $FilePath = get-childitem -path "$($FilePath)\*.pfx" -ea STOP |?{$_.name -match 'mymail-torolab-com-'} | select -expand fullname ; 
                    }
                    DEFAULT{
                        $smsg = "UNCONFIGURED/UNRECOGNIZED `$ENV:USERDOMAIN: $($env:USERDOMAIN)`nSKIPPING CERT CHAIN OF TRUST IMPORT" ; 
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                    }
                }
                if(test-path $FilePath -PathType Leaf){
                    $smsg = "Using $($env:USERDOMAIN) -discovered certificate PFX file:$($FilePath)" ;
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;                    
                }else {
                    $smsg = "UNABLE TO RESOLVE -FILEPATH $($FilePath) TO A SUITABLE PFX FILE!" ; 
                    if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                }
            } ; 
        }
        PROCESS{
            foreach($CertificatePath in $FilePath){                
                if($CertificatePath = gci $CertificatePath){
                    if($CertificatePath.name.tolower() -match '\.pfx$' -and -not $Password){
                        $smsg = "Next dialog will prompt for PFX credential,`nfor:$($CertificatePath.name)`nenter anything ('dummy') for Username, it will not be used" ; 
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        $Password = (get-credential -Credential dummy).password ; 
                    } ; 
                    # demo cmdline: Import-ExchangeCertificate -FileData ([Byte[]]$(Get-Content -Path "D:\cab\mymail-torolab-com-(10-10-2025).pfx" -Encoding Byte -ReadCount 0)) -Password (Read-Host "Enter Certificate Password" -AsSecureString)
                    $pltIpExCert=@{
                        #FileData=([Byte[]]$(Get-Content -Path $CertificatePath.fullname -Encoding byte -ReadCount 0 -erroraction 'STOP')) ;
                        #Password=$Password ;
                        erroraction = 'STOP' ;
                        whatif = $($whatif) ;
                    } ;
                    if($Password){
                        $pltIpExCert.add('Password',$Password) ; 
                    }
                    # note UNC sources are disabled now!
                    # Import-ExchangeCertificate -Server Mailbox01 -FileData ([System.IO.File]::ReadAllBytes('\\FileServer01\Data\Exported Fabrikam Cert.pfx')) -Password (Get-Credential).password
                    switch -regex ($ExVers){
                        'Ex2010'{
                            $pltIpExCert.add('FileData',([Byte[]]$(Get-Content -Path $CertificatePath.fullname -Encoding byte -ReadCount 0 -erroraction 'STOP'))) ;  
                        }
                        'Ex2013'{
                            $pltIpExCert.add('FileName',$CertificatePath.fullname) ;   
                        }
                        'ExSE|Ex2019|Ex2016'{
                            #$pltIpExCert.add('FileData',([System.IO.File]::ReadAllBytes($CertificatePath.fullname))) ;   
                            $pltIpExCert.add('FileData',([System.IO.File]::ReadAllBytes($CertificatePath.fullname))) ;   
                        }
                        default{
                            $smsg = "UNCONFIGURED `$ExVers:$($ExVers)!" ; 
                            $smsg += "`nThis script has not been configured (yet) to support ImportExchangeCertificate syntax on the specified version" ; 
                            if(gcm Write-MyError -ea 0){Write-MyError $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error} else{ write-ERROR "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                            break ; 
                        }
                    }
                    $smsg = "Import-ExchangeCertificate  w`n$(($pltIpExCert|out-string).trim())" ; 
                    if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;
                    TRY{
                        $results = Import-ExchangeCertificate @pltIpExCert ;                 
                    } CATCH [InvalidOperationException]{
                        $ErrTrapd=$Error[0] ;
                        $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                        $smsg += "`n$($ErrTrapd.Exception.Message)" ;
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                        $smsg += "`n$($ErrTrapd.Exception.Message)" ;
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } CATCH[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                    } ;
                    if(-not ($whatif) -and $results.Thumbprint){
                        if (-not $whatif -AND ($rout = Get-ExchangeCertificate -server $env:computername -thumb $results.Thumbprint)) {
                            #$rout | select Subject, Services, not*, thumb*, friend* | write-output ;
                            # return the cert not summary
                            $rout | write-output ; 
                        } elseif (-not $whatif) {
                            $false | write-output ;
                        } else {
                            $smsg = "-whatif:$($whatif): skipped import/confirmation" ;
                            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            } ;
                        } ;
                    }else{
                        $smsg = "Import-ExchangeCertificate RETURNED NO POPULATED THUMBPRINT!(BUG?)"
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        $false | write-output ;
                    }
                    if(-not $whatif -AND ($rout = Get-ExchangeCertificate -server $env:computername -thumb $results.Thumbprint)){
                        $rout | select Subject,Services,not*,thumb*,friend* | write-output ;
                    } elseif(-not $whatif){
                        $false | write-output ; 
                    }else {                        
                        $smsg = "-whatif:$($whatif): skipped import/confirmation" ; 
                        if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ;
                    } ; 
                } else { 
                   $smsg = "Unable to locate cert `$pfx!:$($CertificatePath)" ; 
                   if(gcm Write-MyError -ea 0){Write-MyError $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error}
                        else{ write-ERROR "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } ;
            }  # loop-E       
        } ;  # PROC-E
    }
#endregion IMPORT_XOPCERTIFICATEPFXTDO ; #*------^ END FUNCTION import-xopCertificatePfxTDO  ^------

# SIG # Begin signature block
# MIIELgYJKoZIhvcNAQcCoIIEHzCCBBsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU433NpSIIPT5zF3awK7+X7p46
# BWygggI4MIICNDCCAaGgAwIBAgIQWsnStFUuSIVNR8uhNSlE6TAJBgUrDgMCHQUA
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
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQH0Pjo
# ZHWCGO2puy3xfEIRWfNIvDANBgkqhkiG9w0BAQEFAASBgGVTI+uZLXMuJpxaNrHY
# VHCgN2xH/dacuGemDpZXgvu9dotU96rNmCMGD25UC/RiV9Iio8xCC0Xp2Szvpk4V
# d+ImZl0UdWmXEeXbFI8NcZx0GLlChPXDVti7QIlPrjwNIrCsq6R7grC7/y+BH8mR
# Qlsao5eu9sGSTHs1FP6OXfDO
# SIG # End signature block
