#region TEST_XOPCERTICATEBINDING ; #*------v FUNCTION test-xopCerticateBinding v------
function test-xopCerticateBinding{
        <#
        .SYNOPSIS
        test-xopCerticateBinding - Use get-ExchangeCertificate (EMS) to retrieve current Services-bound certificate, and evaluate that it meets the necessary validity checks for production use (Subject matches production host names; Status: Valid; NotAfter & NotBefore are in window for use, IISServices has a binding, PrivateKeyExportable:`$true, and Issuer is the expected Public vendor for public authentication
        .NOTES
        Version     : 0.0.3
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2025-11-13
        FileName    : test-xopCerticateBinding.ps1
        License     : MIT License
        Copyright   : (c) 2025 Todd Kadrie
        Github      : https://github.com/tostka/verb-XXX
        Tags        : Powershell,Exchange,ExchangeServer,Certificate,Validation,Maintenance
        AddedCredit : REFERENCE
        AddedWebsite: URL
        AddedTwitter: URL
        REVISIONS
        * 11:05 AM 11/19/2025 add break on dep-fails (wasn't exiting) ; add test for broken signing on ex16 - doesn't return Thumbprint or other properties; convert to func(), add to xopBL        * 1:20 PM 11/13/2025 init, works in production.
        .DESCRIPTION
        test-xopCerticateBinding - Use get-ExchangeCertificate (EMS) to retrieve current Services-bound certificate, and evaluate that it meets the necessary validity checks for production use (Subject matches production host names; Status: Valid; NotAfter & NotBefore are in window for use, IISServices has a binding, PrivateKeyExportable:`$true, and Issuer is the expected Public vendor for public authentication
        .INPUTS
        None. Does not accepted piped input.(.NET types, can add description)
        .OUTPUTS
        Returns psjCustomObject reflecting status of each configured test
        .EXAMPLE
        PS> test-xopCerticateBinding -verbose 

        =MATCHED CERT ON AAAAAnnn:

        Subject              : CN=AAAAAA.AAAA.com, O=The Toro Company, L=Bloomington, S=Minnesota, C=US
        Status               : Valid
        Services             : IMAP, POP, IIS, SMTP
        NotAfter             : 11/2/2026 5:59:59 PM
        NotBefore            : 11/2/2025 6:00:00 PM
        Thumbprint           : 6F029E0D20605DBTRIM7016B787C4503FD0002C0
        IisServices          : {IIS://AAAAAnnn/W3SVC/1}
        PrivateKeyExportable : True
        PublicKeySize        : 2048
        FriendlyName         : AAAAAA.AAAA.com (11-02-2026)
        Issuer               : CN=DigiCert Global G2 TLS RSA SHA256 2020 CA1, O=DigiCert Inc, C=US

        =CN=AAAAAA.AAAA.com, O=AAA AAAA AAAAAAA, L=AAAAAAAAAAA, S=AAAAAAAAA, C=US 
        SANS
        DNS Name=AAAAAA.AAAA.com
        DNS Name=autodiscover.AAAA.com

        Thumbprint                                Services   Subject                                                                                                  
        ----------                                --------   -------                                                                                                  
        6F029E0D20605DBTRIM7016B787C4503FD0002C0  IP.WS.     CN=AAAAAA.AAAA.com, O=AAA AAAA AAAAAAA, L=AAAAAAAAAAA, S=AAAAAAAAA, C=US                                 

        VERBOSE: Tests to be Run:
        subject match CN=AAAAAA\.
        status eq Valid
        NotAfter gt NOW
        NotBefore lt NOW
        IISServices eq TRUE
        PrivateKeyExportable eq TRUE
        issuer match CN=DigiCert

        subject	[Y] MATCH	CN=AAAAAA\. 	:	CN=AAAAAA.AAAA.com, O=AAA AAAA AAAAAAA, L=AAAAAAAAAAA, S=AAAAAAAAA, C=US
        status	[Y] -EQ	Valid 	:	Valid
        NotAfter	[Y] -GT	11/13/2025 13:30:31 	:	11/02/2026 17:59:59
        NotBefore	[Y] -LT	11/13/2025 13:30:31 	:	11/02/2025 18:00:00
        IISServices	[Y] -EQ	True 	:	Microsoft.Exchange.Management.SystemConfigurationTasks.IisService
        PrivateKeyExportable	[Y] -EQ	True 	:	True
        issuer	[Y] MATCH	CN=DigiCert 	:	CN=DigiCert Global G2 TLS RSA SHA256 2020 CA1, O=DigiCert Inc, C=US

        $TestResults:
        subject	[Y] MATCH	CN=AAAAAA\. 	:	CN=AAAAAA.AAAA.com, O=AAA AAAA AAAAAAA, L=AAAAAAAAAAA, S=AAAAAAAAA, C=US
        status	[Y] -EQ	Valid 	:	Valid
        NotAfter	[Y] -GT	11/13/2025 13:30:31 	:	11/02/2026 17:59:59
        NotBefore	[Y] -LT	11/13/2025 13:30:31 	:	11/02/2025 18:00:00
        IISServices	[Y] -EQ	True 	:	Microsoft.Exchange.Management.SystemConfigurationTasks.IisService
        PrivateKeyExportable	[Y] -EQ	True 	:	True
        issuer	[Y] MATCH	CN=DigiCert 	:	CN=DigiCert Global G2 TLS RSA SHA256 2020 CA1, O=DigiCert Inc, C=US

        .LINK
        https://github.com/tostka/powershellbb/
        #>
        [CmdletBinding()]
        PARAM()
        BEGIN{
            # PKI cert object: doesn't include Status and other key values; must obtain via get-exchangecertificate cmdlet.
            #$CNName = 'CN=mymail.toro.com*' ;
            #$NotBefore = '11/2/2025 6:00:00 PM' ;
            #$cert=get-childitem cert:\LocalMachine\My\ |?{$_.Subject -like $CNName -AND $_.NotBefore -ge (get-date $NotBefore)} ;
            #$cert | select subject,friendlyname,not*,DNSName,pspath,thumbprint ;
            if(-not (get-command get-exchangecertificate)){
                $smsg = "MISSING EMS: get-exchangecertificate cmdlet" ; 
                $smsg += "`nPre-connect to Exchange Management Shell *locally* on an Exchange Server to run this check" ; 
                $smsg += "`n(must be local EMS: Remote EMS lacks Status and other key values on return)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                throw $smsg ; 
                break ; 
            }
            $cert=get-exchangecertificate -Server $env:LOCALHOST |?{($_.Services -match "IIS") -AND ($_.Services -match "SMTP") -AND ($_.Services -match "IMAP") -AND ($_.Services -match "POP")}  ;"`n=MATCHED CERT ON $($env:COMPUTERNAME):" ;$cert | fl Subject,Status,Services,not*,thumb*,Iis*,PrivateKeyEx*,PublicKeyS*,Friend*,Issuer ;if($cert -isnot [system.array]){  $tCERT=gci cert:\LocalMachine\My\$($cert.Thumbprint) ;  "=$($tcert.Subject) SANS" ;  ($tCert.Extensions | Where-Object {$_.Oid.FriendlyName -eq "subject alternative name"}).Format(1) ;} else {    throw "multiple certs returned! $($cert.Subject| out-string)" ;};
            # 10:41 AM 11/19/2025 add test for broken signing on ex16 - doesn't return Thumbprint or other properties
            if(($cert | gm | ?{$_.MemberType -eq 'property'} | select -expand name) -contains 'Thumbprint'){
                $smsg = "(confirmed get-exchangecertificate returns a populated Thumbprint at mininum: no Fed Cert blank return issue)" ; 
                if($VerbosePreference -eq "Continue"){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            }else{
                $smsg = "BROKEN EMS: get-exchangecertificate cmdlet!" ; 
                $smsg += "`nSigning-Behavior: FAILS to return properly populated Thumbprint, Status, Subject, Services or other normal propertiesr" ; 
                $smsg += "`n(requires rebuild of Federation cert, with CAB'd post-run of HCW to repair to function)" ; 
                $smsg += "`nTHIS TEST WILL NOT FUNCTION UNTIL GET-EXCHANGECERTIFICATE IS FIXED AND FUNCTIONAL!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                THROW $SMSG ; 
                BREAK ; 
            }
            write-host "`n" ;    
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):matched cert:`n$(($cert|out-string).trim())" ; 

            $TestsValid = @(
                'subject;match;CN=mymail\.',
                'status;eq;Valid',
                'NotAfter;gt;NOW',
                'NotBefore;lt;NOW',
                'IISServices;eq;TRUE',
                'PrivateKeyExportable;eq;TRUE',
                'issuer;match;CN=DigiCert'
            ) ;
            $smsg = "Tests to be Run:`n$(($testsvalid | %{$_.replace(';',' ')}|out-string).trim())" ; 
            write-verbose $smsg ; 
            $TestResults = @() ;
        } # BEG-E
        PROCESS{
            foreach($test in $TestsValid){
                $prop,$comp,$valu = $null ;
                write-verbose $test ;
                $prop,$comp,$valu = $test.split(';') ;
                write-verbose $(@($prop,$comp,$valu) -join ',')
                switch ($valu){
                    'NOW'{$valu = (get-date ) }
                    'TRUE'{$valu = $true}
                    'FALSE'{$valu = $false}
                    default{
                        write-verbose "passthru usable valu:$($valu)" ;
                    } ;
                } ;
                switch ($comp){
                    'match'{
                        if(($cert | select -expand $prop) -MATCH $valu){
                            $TestResults += "$($prop)`t[Y] MATCH`t$($valu) `t:`t$($cert | select -expand $prop)" ;
                        } else{
                            $TestResults += "$($prop)`t*[X]* $($valu) `t:`t$($cert | select -expand $prop)" ;
                        };
                    }
                    'gt'{
                        if(($cert | select -expand $prop) -GT $valu){
                            $TestResults += "$($prop)`t[Y] -GT`t$($valu) `t:`t$($cert | select -expand $prop)" ;
                        } else{
                            $TestResults += "$($prop)`t*[X]* $($valu) `t:`t$($cert | select -expand $prop)" ;
                        };
                    }
                    'lt'{
                        if(($cert | select -expand $prop) -LT $valu){
                            $TestResults += "$($prop)`t[Y] -LT`t$($valu) `t:`t$($cert | select -expand $prop)" ;
                        } else{
                            $TestResults += "$($prop)`t*[X]* $($valu) `t:`t$($cert | select -expand $prop)" ;
                        };
                    }
                    'eq'{
                        if($valu -eq $true){
                            if(($cert | select -expand $prop) | ?{$_ -eq $valu -or $_ -ne $null}){
                                $TestResults += "$($prop)`t[Y] -EQ`t$($valu) `t:`t$($cert | select -expand $prop)" ;
                            } else{
                                $TestResults += "$($prop)`t*[X]* $($valu) `t:`t$($cert | select -expand $prop)" ;
                            };
                        }elseif($valu -eq $false){
                            if(($cert | select -expand $prop) | ?{$_ -eq $valu -or $_ -eq $null}){
                                $TestResults += "$($prop)`t[Y] -EQ`t$($valu) `t:`t$($cert | select -expand $prop)" ;
                            } else{
                                $TestResults += "$($prop)`t*[X]* $($valu) `t:`t$($cert | select -expand $prop)" ;
                            };
                        }else{
                          if(($cert | select -expand $prop) -EQ $valu){
                                $TestResults += "$($prop)`t[Y] -EQ`t$($valu) `t:`t$($cert | select -expand $prop)" ;
                            } else{
                                $TestResults += "$($prop)`t*[X]* $($valu) `t:`t$($cert | select -expand $prop)" ;
                            };
                        }
               
                    }
                    default{
                        throw "unrecognized `$comp! comparision operator!:$($comp)" ;
                    }
                } ;
            } ; # loop-E
        }  # PROC-E
        END{
            $TestResults | write-output ;
            WRITE-HOST "`n`$TestResults:`n$(($TestResults|out-string).trim())`n" ; 
        } ; 
    }
#endregion TEST_XOPCERTICATEBINDING ; #*------^ END FUNCTION test-xopCerticateBinding  ^------

# SIG # Begin signature block
# MIIELgYJKoZIhvcNAQcCoIIEHzCCBBsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUuh+A3TtLxFatRIes0g3XC2n/
# 1umgggI4MIICNDCCAaGgAwIBAgIQWsnStFUuSIVNR8uhNSlE6TAJBgUrDgMCHQUA
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
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBTt1mUI
# OF+FcTkQuveBme0fPOd6fzANBgkqhkiG9w0BAQEFAASBgKax0fTFbF3rYj48lCL9
# ZTZVoGSiW8DK2ocPsHWll3ivvumXfv0cL8bgiaW7DBIT3BMz+wvW9w78IW1sJFxT
# TGvYzxA0tpvnLoX+wc4l2KlBpu3gPIu2vbSZ+ovmWbobQB/pWF9E8qlbhGertidJ
# uYoPoSeiI48xC9dGm2K7N9Cl
# SIG # End signature block
