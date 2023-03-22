#rebuild-module.ps1

<#
.SYNOPSIS
rebuild-module.ps1 - Rebuild verb-Ex2010 & publish to localrepo
.NOTES
Version     : 1.0.0
Author      : Todd Kadrie
Website     : http://www.toddomation.com
Twitter     : @tostka / http://twitter.com/tostka
CreatedDate : 2020-03-17
FileName    : rebuild-module.ps1
License     : MIT License
Copyright   : (c) 2020 Todd Kadrie
Github      : https://github.com/tostka
Tags        : Powershell
REVISIONS
* 8:49 AM 3/17/2020 init
.DESCRIPTION
rebuild-module.ps1 - Rebuild verb-Ex2010 & publish to localrepo
.PARAMETER Whatif
Parameter to run a Test no-change pass [-Whatif switch]
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output
.EXAMPLE
.\rebuild-module.ps1 -whatif ; 
Rebuild pass with -whatif
.EXAMPLE
.\rebuild-module.ps1
Non-whatif rebuild
.LINK
https://github.com/tostka
#>
[CmdletBinding()]
PARAM(
    [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
    [switch] $whatIf
) ;
$Verbose = ($VerbosePreference -eq 'Continue') ; 
write-verbose -verbose:$verbose "`$PSBoundParameters:`n$(($PSBoundParameters|out-string).trim())" ; 
pushd ; 
cd c:\usr\work\ps\scripts\
.\process-NewModule.ps1 -ModuleName "verb-Ex2010" -ModDirPath "C:\sc\verb-Ex2010" -Repository $localPSRepo -Merge -RunTest -showdebug -whatif:$($whatif) ;
popd ; 

# SIG # Begin signature block
# MIIELgYJKoZIhvcNAQcCoIIEHzCCBBsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUlUYfR248GBUAqHx6Vcq2sJ6k
# JoagggI4MIICNDCCAaGgAwIBAgIQWsnStFUuSIVNR8uhNSlE6TAJBgUrDgMCHQUA
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
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQgs0vb
# tvdHJzp48EkTIIxQQCZeMjANBgkqhkiG9w0BAQEFAASBgIh+xTDgzzjq9fGdbmmK
# AiyGr8iG6yGEhiieTxUB8d2X99X+JM7DeXqBy8ACQuAnmqzZ8hRDLCRcwpI7jHIj
# 2UeKdDT+9tKUyruSzel7VG2qe+UAqj/F/vQ38mKyYUeIK6dDCdyflfkaJFn6tEJk
# T9Bajvg2Miy4OVQ/rJBZvMVT
# SIG # End signature block
