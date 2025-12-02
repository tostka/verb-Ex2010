# Convert-xopAdminDisplayVersionTDO.ps1

    #region CONVERT_XOPADMINDISPLAYVERSIONTDO ; #*------v Convert-xopAdminDisplayVersionTDO v------
    function Convert-xopAdminDisplayVersionTDO {
            <#
            .SYNOPSIS
            Convert-xopAdminDisplayVersionTDO - Convert Exchange Server AdminDisplayVersion (as returned by EMS get-ExchangeServer) to Semantic Version (n.n.n.n)
            .NOTES
            Version     : 0.0.1
            Author      : Todd Kadrie
            Website     : http://www.toddomation.com
            Twitter     : @tostka / http://twitter.com/tostka
            CreatedDate : 2025-12-02
            FileName    : Convert-xopAdminDisplayVersionTDO.ps1
            License     : MIT License
            Copyright   : (c) 2025 Todd Kadrie
            Github      : https://github.com/tostka/verb-ex2010
            Tags        : Powershell,Exchange,ExchangeServer,Install,Patch,Maintenance
            AddedCredit : 
            AddedWebsite: https://www.google.com/search?client=firefox-b-1-d&q=powershell+convert+exchange+admindisplayversion+to+semantic+version
            AddedTwitter: URL
            REVISIONS
            * 12:10 PM 12/2/2025 revised to named caps;  init version

            .DESCRIPTION
            Convert-xopAdminDisplayVersionTDO - Convert Exchange Server AdminDisplayVersion 'Version [major].[minor] (Build [buildmajor].[buildminor])' (as returned by EMS get-ExchangeServer) to equiv Semantic Version (major.minor.buildmajor.build.minor)

            Simple job: 
            - the AdminDisplayVersion string: 'Version 15.1 (Build 2507.6)'
            ...represents following SemVersion components: 
                'Version [major].[minor] (Build [buildMajor].[buildMinor])'
            ... which just need to be regex parsed and lined up into equiv semversion:
                [major].[minor].[buildMajor].[buildMinor]
                == 15.1.2507.6

            .PARAMETER AdminDisplayVersion
            .INPUTS
            Accepts piped input.
            .OUTPUTS
            System.Version Semantic Version object
            .EXAMPLE
            PS> $VersionNum = Convert-xopAdminDisplayVersionTDO -AdminDisplayVersion ((get-exchangeserver Server1).AdminDisplayVersion) ; 
            PS> $VersionNum ; 

                15.1.2507.6

            PS> $VersInfo = Resolve-xopBuildSemVersToTextNameTDO -Version $VersionNum ; 
            PS> $VersInfo ; 

                ProductName      : Exchange Server 2016 CU23 (2022H1)
                ReleaseDate      : 4/20/2022
                BuildNumberShort : 15.1.2507.6
                BuildNumberLong  : 15.01.2507.006
                PatchBasis       : Exchange Server 2016 CU23
                NickName         : EX2016_CU23_2022H1
                IsInstallable    : TRUE

            Demo resolving get-exchangeserver AdminDisplayVersion to Semantic Version, and then resolving that version to Exchange Version info through vx10\Resolve-xopBuildSemVersToTextNameTDO.
            .EXAMPLE
            PS> $VersionNum = Convert-xopAdminDisplayVersionTDO -AdminDisplayVersion 'Version 15.1 (Build 2507.6)' -verbose ; 
            PS> $VersionNum ; 
            Demo resolving an AdminDisplayVersion static string value to Semantic Version string
            .LINK
            https://github.com/tostka/verb-ex2010
            #>
        [CmdletBinding()]
        [Alias('Convert-AdminDisplayVersion')]
        PARAM(
            [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
            [string]$AdminDisplayVersion
        )
        PROCESS {
            # Example AdminDisplayVersion: "Version 15.1 (Build 1913.5)"
             <# # prior position caps vers
             if ($AdminDisplayVersion -match 'Version (\d+)\.(\d+)\s+\(Build (\d+)\.(\d+)\)') {
                $major = $matches[1]
                $minor = $matches[2]
                $buildMajor = $matches[3]
                $buildMinor = $matches[4]                
                $semanticVersion = "$major.$minor.$buildMajor.$buildMinor"
            #>
            # flip to named caps rgx
            [regex]$rx = "Version\s(?<major>\d+)\.(?<minor>\d+)\s\(Build\s(?<buildmajor>\d+)\.(?<buildminor>\d+)\)" ;
            if ($AdminDisplayVersion -match $rx){               
                [version]$semanticVersion = "$($matches.major).$($matches.minor).$($matches.buildmajor).$($matches.buildminor)"
                Write-Output $semanticVersion
            } else {
                $smsg = "Could not parse AdminDisplayVersion: '$AdminDisplayVersion'"
                if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            }
        } ;  # PROC-E
    } ; 
    #endregion CONVERT_XOPADMINDISPLAYVERSIONTDO ; #*------^ END Convert-xopAdminDisplayVersionTDO ^------
    