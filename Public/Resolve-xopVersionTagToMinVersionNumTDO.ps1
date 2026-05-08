#region RESOLVE_XOPVERSIONTAGTOMINVERSIONNUMTDO ; #*------v FUNCTION Resolve-xopVersionTagToMinVersionNumTDO v------
Function Resolve-xopVersionTagToMinVersionNumTDO {
        <#
        .SYNOPSIS
        Resolve-xopVersionTagToMinVersionNumTDO - Resolves Exchange Server Major Server Revision tag (EXSE|EX2019|EX2016|EX2013|EX2010|EX2007|EX2003|EX2000|EX55|EX50|EX40_SE) to Min/RTM SemanticVersion BuildNumber for filtering server lists to return specific revisions 
        .NOTES
        Version     : 0.0.
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 20250711-0423PM
        FileName    : Resolve-xopVersionTagToMinVersionNumTDO.ps1
        License     : (none asserted)
        Copyright   : (none asserted)
        Github      : https://github.com/tostka/verb-ex2010
        Tags        : Powershell,Exchange,ExchangeServer,Install,Patch,Maintenance
        AddedCredit : 
        AddedWebsite: 
        AddedTwitter: URL
        REVISIONS
        * 3:06 PM 11/26/2025 init version, simplified major version version of Resolve-xopBuildSemVersToTextNameTDO, returns solely the build tag, not further details

        .DESCRIPTION
        Resolve-xopVersionTagToMinVersionNumTDO - Resolves Exchange Server Major Server Revision tag (EXSE|EX2019|EX2016|EX2013|EX2010|EX2007|EX2003|EX2000|EX55|EX50|EX40_SE) to Min/RTM SemanticVersion BuildNumber for filtering server lists to return specific revisions 
        
        Returns the following variant Version Tags:

            Tag     | Server Major Release Version
            ------- | ------------------------------------
            EXSE    | Exchange Server Subscription Edition
            EX2019  | Exchange Server 2019
            EX2016  | Exchange Server 2016
            EX2013  | Exchange Server 2013
            EX2010  | Exchange Server 2010
            EX2007  | Exchange Server 2007
            EX2003  | Exchange Server 2003
            EX2000  | Exchange Server 2000
            EX55    | Exchange Server 5.5
            EX50    | Exchange Server 5.0
            EX40_SE | Exchange Server 4.0 SE

        Supports Exchange Server 4.0 SE through Exchange Server Subscription Edition. 

        Simple check of semver against RTM/Preview etc initial release versions ('Breakpoints'): 

            Breakpoints for major versions (1st/lowest BuildNumberShort for the revision level):
            NickName           BuildNumberShort
            --------           ----------------
            EXSE_RTM           15.2.2562.17
            EX2019_Preview     15.2.196.0
            EX2016_RTM         15.1.225.42
            EX2016_Preview     15.1.225.16
            EX2013_RTM         15.0.516.32
            EX2010_RTM         14.0.639.21
            EX2007_RTM         8.0.685.25
            EX2003             6.5.6944
            EX2000             6.0.4417
            EX55               5.5.1960
            EX50               5.0.1457
            EX40_SE            4.0.837

        .PARAMETER Version
        .INPUTS
        None, no piped input.
        .OUTPUTS
        System.String Exchange Server Major Version Tag
        .EXAMPLE
        PS> $SemVers = Resolve-xopVersionTagToMinVersionNumTDO -VersionTag Ex2016
        PS> $SemVers ; 

            Major  Minor  Build  Revision
            -----  -----  -----  --------
            15     1      225    16      

        Demo resolving Ex2016 minimum revision semantic version number
        .EXAMPLE
        PS> $SemVers = Resolve-xopVersionTagToMinVersionNumTDO -VersionTag Ex2019 ; 
        PS> $SemVers ; 

            Major  Minor  Build  Revision
            -----  -----  -----  --------
            15     2      196    0  

        Demo resolving Ex2019 minimum revision semantic version number

        .LINK
        https://github.com/tostka/verb-ex2010        
        .LINK
        https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates
        #>
        [CmdletBinding()]
        #[alias('Get-DetectedFileVersion')]
        PARAM(
            [Parameter(Mandatory=$TRUE,HelpMessage = "Exchange Server Major Server Revision tag (EXSE|EX2019|EX2016|EX2013|EX2010|EX2007|EX2003|EX2000|EX55|EX50|EX40_SE)[-Version 'EX2016']")]                
                [alias('FileVersion')]
                [string]$VersionTag        
        ) ;         
        PROCESS {
            # when updating $BuildToProductName table (below), also record date of last update here (echos to console, for awareness on results)
            [datetime]$lastBuildTableUpedate = '2025-11-26' ; 
            $BuildTableUpedateUrl = 'https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates'
            #'https://docs.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-date' ; 
            #Creating the hash table with build numbers and cumulative updates
            # updated as of 9:56 AM 3/26/2025 to curr https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates?view=exchserver-2019
            # also using unmodified MS Build names, from the chart (changing just burns time)
            $smsg = "NOTE:`$BuildToProductName table was last updated on $($lastBuildTableUpedate.ToShortDateString())" ; 
            $smsg += "`n(update from:$($BuildTableUpedateUrl))" ;
            write-host -foregroundcolor yellow $smsg ; 
          
            if($VersionTag){
                switch($VersionTag){
                    'EXSE' {$MinVersion = [version]'15.2.196.0' }
                    'EX2019' {$MinVersion = [version]'15.2.196.0' }
                    'EX2016' {$MinVersion = [version]'15.1.225.16' }
                    'EX2013' {$MinVersion = [version]'15.0.516.32' }
                    'EX2010' {$MinVersion = [version]'14.0.639.21' }
                    'EX2007' {$MinVersion = [version]'8.0.685.25' }
                    'EX2003' {$MinVersion = [version]'6.5.6944' }
                    'EX2000' {$MinVersion = [version]'6.0.4417' }
                    'EX55' {$MinVersion = [version]'5.5.1960' }
                    'EX50' {$MinVersion = [version]'5.0.1457' }
                    'EX40_SE' {$MinVersion = [version]'4.0.837' }
                    default{
                        $smsg = "Unrecognized -VersionTag: $($VersionTag)" ; 
                        write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ; 
                    }
                }; 
                if($MinVersion){
                    $MinVersion | write-output ;
                } 
               
            } ; 
        };  # PROC-E        
    }
#endregion RESOLVE_XOPVERSIONTAGTOMINVERSIONNUMTDO ; #*------^ END FUNCTION Resolve-xopVersionTagToMinVersionNumTDO  ^------

# SIG # Begin signature block
# MIIELgYJKoZIhvcNAQcCoIIEHzCCBBsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUEEBN9NzdV1Lk84UJnybdzlMJ
# Im6gggI4MIICNDCCAaGgAwIBAgIQWsnStFUuSIVNR8uhNSlE6TAJBgUrDgMCHQUA
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
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBSI/XBg
# UV9ctFdONBpv2NdqsrJ2ZzANBgkqhkiG9w0BAQEFAASBgI+/MlhBFRz92/3dYQDo
# ytbL1xKNLSBR8qJGKnQ8zKBQbZhNlfJjsG8jzSHNertG/JzmWH2NDvtw1nD88mo+
# Gq2xpjyo22Zs+K7fYoTtybLqldaRKzECmAaiZ4BoFTq2cbT7eyWAYczJvybpsMtV
# od94JJnpJiqKhRTeq81WEeli
# SIG # End signature block
