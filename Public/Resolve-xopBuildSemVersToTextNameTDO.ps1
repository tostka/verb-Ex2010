# Resolve-xopBuildSemVersToTextNameTDO.ps1

    #region RESOLVE_XOPBUILDSEMVERSTOTEXTNAMETDO ; #*------v Resolve-xopBuildSemVersToTextNameTDO v------
    Function Resolve-xopBuildSemVersToTextNameTDO {
        <#
        .SYNOPSIS
        Resolve-xopBuildSemVersToTextNameTDO - Resolves Exchange Server SemanticVersion BuildNumber to MS Build/Release information details
        .NOTES
        Version     : 0.0.
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 20250711-0423PM
        FileName    : Resolve-xopBuildSemVersToTextNameTDO.ps1
        License     : (none asserted)
        Copyright   : (none asserted)
        Github      : https://github.com/tostka/verb-ex2010
        Tags        : Powershell,Exchange,ExchangeServer,Install,Patch,Maintenance
        AddedCredit : Michel de Rooij / michel@eightwone.com
        AddedWebsite: http://eightwone.com
        AddedTwitter: URL
        REVISIONS
        * 9:24 AM 10/22/2025 FIXED REGION TAGS
        * 10:22 AM 9/25/2025 ADD: -AllVersions to return the raw versions table ; added IsInstallable field, to tag installable base RTM/SP/CU releases; 
            CBH: add fields specs for -AllVersions; 
            add: demos for postfiltering -AllVersions returned data to find installable 
        * 4:07 PM 9/24/2025 flip return from [hashtable] to [pscustomobject]; 
             add alias & param alias for install-Exchange15-TTC.ps1\Get-FileVersion() emulation; pull -LongBuildNumber: the shift to [system.version]$Version obviates differences: 
            it autoflattens zero-pads to short format when .tostring()'d. So all compares are effectively to the shortversionnumber, regardless of input format.
        * 4:25 PM 9/23/2025 flipped -Version [string]->[version], does it's own type validation; updated table added NickName in format: EX[vers]_[SpCU]_[HuSu]
        * 4:02 PM 9/22/2025 init; Updated BuildToProductName indexed hash to specs posted as of 04/25/2025 at https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates
            Maxes reflected in this script, as of that time:
             - Exchange Server SE RTM Sep25HU 	September 8, 2025 	15.2.2562.27 	15.02.2562.027
             - Exchange Server 2019 CU15 Sep25HU 	September 8, 2025 	15.2.1748.37 	15.02.1748.037
             - Exchange Server 2016 CU23 Sep25HU 	September 8, 2025 	15.1.2507.59 	15.01.2507.059
             - Exchange Server 2013 CU23 Mar23SU 	March 14, 2023 	15.0.1497.48 	15.00.1497.048
             - Update Rollup 32 for Exchange Server 2010 SP3 	March 2, 2021 	14.3.513.0 	14.03.0513.000
             - Update Rollup 23 for Exchange Server 2007 SP3 	March 21, 2017 	8.3.517.0 	8.03.0517.000
             - Exchange Server 2003 post-SP2 	August 2008 	6.5.7654.4
             - Exchange 2000 Server post-SP3 	August 2008 	6.0.6620.7
             - Exchange Server version 5.5 SP4 	November 1, 2000 	5.5.2653
             - Exchange Server 5.0 SP2 	February 19, 1998 	5.0.1460
             - Exchange Server 4.0 SP5 	May 5, 1998 	4.0.996

        .DESCRIPTION
        Resolve-xopBuildSemVersToTextNameTDO - Resolves Exchange Server SemanticVersion BuildNumber to MS Build/Release information details
        
        -AllVersions fields (and those in use in source table):
        - Most, aside from PatchBasis, NickName & 'IsInstallable' are the field names and data directly lifted from the source MS table above
        - ProductName | the "official" MS name for the RTM/SP/CU/SU/HU release (ReleaseToManuf,ServicePack,CumulativeUpdate,ServiceUpdate,HotfixUpdate)
        - ReleaseDate | documented release date
        - BuildNumberShort | SemanticVersion string with 0-padding removed
        - BuildNumberLong |  SemanticVersion string with 0-padding intact
        - PatchBasis | The most recent full IsInstallable RTM/SP/CU prior to a given HU,SU 'patch'
        - NickName | converted from ProductName: 1) subst 'EX' for 'Exchange\sServer', 2) Move 'Update\sRollup\s\d+' prefix to suffix, 3) subst '_' for '\s', 4) remove parenthesis, 5) replace 'Standard\sEdition' w 'SE'
        - IsInstallable | represents an RTM, SP or CU release in setup.exe-installable ISO or CAB availablility: e.g. the baseline installs that get patched with later CU/SU/HU updates to reach fully patched status.

        ## To Update `$xopBuilds table in this function to current published specs:

        1. Review https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates for updates
        2. Copy updates across releases (each is in a separate subtable, listed in descending release order) into the xopTable
            - Simplest accomplished by using Excel to do formatting and editing and appending the entries into proper placese
            - Be sure to construct NickName, note the PatchBasis, and specify IsInstallable status (if it's an RTM, SP or CU)
            - Then File > Save As a csv unicode encoded. 
            - then import and reparse the csv into markdowntable format, used as the data source in the xopBuilds inputs below:
            ``````powershell
            $builds = import-csv -path c:\pathto\UpdatedBuilds.csv ; 
            $builds | verb-io\convertto-Markdowntable ; 
            ```
            - copy the displayed console markdown table to clipboard, and paste into the $xopBuilds hashtable between @" and "@.
            Update the $lastBuildTableUpedate below to reflect the date of update: Note, 
            the date displayed at the top of the table, as of 9/25/2025, is *not* the 
            latest update time, it's the original time the article was posted.  

            There is no visible timestamp on the page to indicate last content update, simply you will find later revs at the tops of the tables on the page. 

        .PARAMETER Version
        Exchange Version in Semantic Version Number format (n.n.n.n)[-Version '8.0.708.3']
        .PARAMETER AllVersions
        Switch to return all versions information to pipeline[-AllVersions]
        .INPUTS
        None, no piped input.
        .OUTPUTS
        System.Object summary of Exchange server build specifications.
        .EXAMPLE
        PS> $VersInfo = Resolve-xopBuildSemVersToTextNameTDO -Version '8.0.708.3' ; 
        PS> $VersInfo ; 
        
            ProductName      : Update Rollup 1 for Exchange Server 2007
            ReleaseDate      : 4/17/2007
            BuildNumberShort : 8.0.708.3
            BuildNumberLong  : 8.00.0708.003
            PatchBasis       : Exchange Server 2007 RTM
            NickName         : EX2007_UR1
            IsInstallable    : 

        Demo resolving a Semantic Version string to specific release/build details
        .EXAMPLE
        PS> $VersInfo = Resolve-xopBuildSemVersToTextNameTDO -Version '8.0.708.3' ; 
        PS> $VersInfo.ProductName ;

            Update Rollup 1 for Exchange Server 2007

        Demo resolving semversion to equivelent ProductName string property of the Build returned
        .EXAMPLE
        PS> $SourcePath = 'D:\cab\ExchangeServer2016-x64-CU23-ISO\unpacked' ; 
        PS> $ExSetupVersion = (Get-Command "$($SourcePath)\Setup\ServerRoles\Common\ExSetup.exe").FileVersionInfo.ProductVersion ; 
        PS> $VersInfo = Resolve-xopBuildSemVersToTextNameTDO -Version $ExSetupVersion ; 
        PS> $VersInfo | fl * ; 

            ProductName      : Exchange Server 2016 CU23 (2022H1)
            ReleaseDate      : 4/20/2022
            BuildNumberShort : 15.1.2507.6
            BuildNumberLong  : 15.01.2507.006
            PatchBasis       : Exchange Server 2016 CU23
            NickName         : EX2016_CU23_2022H1
            IsInstallable    : TRUE

        Demo resolving details of local CAB source ExSetup.exe build details.
        .EXAMPLE
        PS> $ExSetupVersion = (Get-Command ExSetup.exe).FileVersionInfo.ProductVersion ; 
        PS> $VersInfo = Resolve-xopBuildSemVersToTextNameTDO -Version $ExSetupVersion ; 
        PS> $VersInfo | fl * ; 

            ProductName      : Update Rollup 30 for Exchange Server 2010 SP3
            ProductName      : Exchange Server 2016 CU23 May25HU
            ReleaseDate      : 5/29/2025
            BuildNumberShort : 15.1.2507.57
            BuildNumberLong  : 15.01.2507.057
            PatchBasis       : Exchange Server 2016 CU23
            NickName         : EX2016_CU23_May25HU
            IsInstallable    : 

        Demo resolving details of local installed Exchange Server bin build details to determine installed CU/SP/RTM base build (by resolving installed Exchanage ExSetup.exe's presence in the local Path evari). Note: This reflects latest patch level (SP/CU/SU/HU)        
        .EXAMPLE
        PS> $ExBinVersion = (Get-Command "$($env:ExchangeInstallPath)Bin\Microsoft.Exchange.Directory.TopologyService.exe").FileVersionInfo.ProductVersion ; 
        PS> $VersInfo = Resolve-xopBuildSemVersToTextNameTDO -Version $ExBinVersion ; 
        PS> $VersInfo | fl * ; 

            ProductName      : Exchange Server 2016 CU23 May25HU
            ReleaseDate      : 5/29/2025
            BuildNumberShort : 15.1.2507.57
            BuildNumberLong  : 15.01.2507.057
            PatchBasis       : Exchange Server 2016 CU23
            NickName         : EX2016_CU23_May25HU
            IsInstallable    : 

        Demo resolving installed ADTopology service .exe to obtain current installed patch revision (will match the much easier tolocate bin dir ExSetup.exe)
        .EXAMPLE
        PS> $VersTable = Resolve-xopBuildSemVersToTextNameTDO -AllVersions
        PS> $VersTable |?{$_.isinstallable -AND $_.NickName -match '^EX2019'} | ft -a 

            -AllVersions: Returning full builds table to pipeline (for post-filtering)

            ProductName                        ReleaseDate BuildNumberShort BuildNumberLong PatchBasis                NickName           IsInstallable
            -----------                        ----------- ---------------- --------------- ----------                --------           -------------
            Exchange Server 2019 CU15 (2025H1) 2/10/2025   15.2.1748.10     15.02.1748.010  Exchange Server 2019 CU15 EX2019_CU15_2025H1 TRUE         
            Exchange Server 2019 CU14 (2024H1) 2/13/2024   15.2.1544.4      15.02.1544.004  Exchange Server 2019 CU14 EX2019_CU14_2024H1 TRUE         
            Exchange Server 2019 CU13 (2023H1) 5/3/2023    15.2.1258.12     15.02.1258.012  Exchange Server 2019 CU13 EX2019_CU13_2023H1 TRUE         
            Exchange Server 2019 CU12 (2022H1) 4/20/2022   15.2.1118.7      15.02.1118.007  Exchange Server 2019 CU12 EX2019_CU12_2022H1 TRUE         
            Exchange Server 2019 CU11          9/28/2021   15.2.986.5       15.02.0986.005  Exchange Server 2019 CU11 EX2019_CU11        TRUE         
            Exchange Server 2019 CU10          6/29/2021   15.2.922.7       15.02.0922.007  Exchange Server 2019 CU10 EX2019_CU10        TRUE         
            Exchange Server 2019 CU9           3/16/2021   15.2.858.5       15.02.0858.005  Exchange Server 2019 CU9  EX2019_CU9         TRUE         
            Exchange Server 2019 CU8           12/15/2020  15.2.792.3       15.02.0792.003  Exchange Server 2019 CU8  EX2019_CU8         TRUE         
            Exchange Server 2019 CU7           9/15/2020   15.2.721.2       15.02.0721.002  Exchange Server 2019 CU7  EX2019_CU7         TRUE         
            Exchange Server 2019 CU6           6/16/2020   15.2.659.4       15.02.0659.004  Exchange Server 2019 CU6  EX2019_CU6         TRUE         
            Exchange Server 2019 CU5           3/17/2020   15.2.595.3       15.02.0595.003  Exchange Server 2019 CU5  EX2019_CU5         TRUE         
            Exchange Server 2019 CU4           12/17/2019  15.2.529.5       15.02.0529.005  Exchange Server 2019 CU4  EX2019_CU4         TRUE         
            Exchange Server 2019 CU3           9/17/2019   15.2.464.5       15.02.0464.005  Exchange Server 2019 CU3  EX2019_CU3         TRUE         
            Exchange Server 2019 CU2           6/18/2019   15.2.397.3       15.02.0397.003  Exchange Server 2019 CU2  EX2019_CU2         TRUE         
            Exchange Server 2019 CU1           2/12/2019   15.2.330.5       15.02.0330.005  Exchange Server 2019 CU1  EX2019_CU1         TRUE         
            Exchange Server 2019 RTM           10/22/2018  15.2.221.12      15.02.0221.012  Exchange Server 2019 RTM  EX2019_RTM         TRUE      

        Demo returning AllVersions table, and post-filtering for installable Exchange 2019 versions 
        .EXAMPLE
        PS> $VersTable = Resolve-xopBuildSemVersToTextNameTDO -AllVersions
        PS> $VersTable |?{$_.NickName -match '^EX2019' -AND $_.isinstallable -ne $true} | select -first 1 

            ProductName      : Exchange Server 2019 CU15 Sep25HU
            ReleaseDate      : 9/8/2025
            BuildNumberShort : 15.2.1748.37
            BuildNumberLong  : 15.02.1748.037
            PatchBasis       : Exchange Server 2019 CU15
            NickName         : EX2019_CU15_Sep25HU
            IsInstallable    : 

        Demo returning AllVersions table, and post-filtering for most recent Exchange 2019 patch versions         
        .LINK
        https://github.com/tostka/verb-ex2010        
        .LINK
        https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates
        #>
        [CmdletBinding()]
        #[alias('Get-DetectedFileVersion')]
        PARAM(
            [Parameter(Mandatory=$false,HelpMessage = "Exchange Version in Semantic Version Number format (n.n.n.n)[-Version '8.0.708.3']")]                
                [alias('FileVersion')]
                [version]$Version,
            [Parameter(Mandatory=$false,HelpMessage = "Switch to return all versions information to pipeline[-AllVersions]")]                
                [switch]$AllVersions
        ) ;        
        BEGIN {
            if(-not ($Version -OR $AllVersions)){$smsg = "Neither -Version or -AllVersions specified`nPlease specify one or the other" } ; 
            if($Version -AND $AllVersions){$smsg = "BOTH -Version & -AllVersions specified`nPlease specify one or the other" } ; 
            # XXXXXX 4:15 PM 9/24/2025
            if($smsg){
                write-warning $smsg ; 
                throw ; 
                break ; 
            } ; 

            # when updating $BuildToProductName table (below), also record date of last update here (echos to console, for awareness on results)
            [datetime]$lastBuildTableUpedate = '9/22/2025' ; 
            $BuildTableUpedateUrl = 'https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates'
            #'https://docs.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-date' ; 
            #Creating the hash table with build numbers and cumulative updates
            # updated as of 9:56 AM 3/26/2025 to curr https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates?view=exchserver-2019
            # also using unmodified MS Build names, from the chart (changing just burns time)
            $smsg = "NOTE:`$BuildToProductName table was last updated on $($lastBuildTableUpedate.ToShortDateString())" ; 
            $smsg += "`n(update from:$($BuildTableUpedateUrl))" ;
            write-host -foregroundcolor yellow $smsg ; 

            #region CONVERTFROM_MARKDOWNTABLE ; #*------v convertfrom-markdowntable v------
            if(-not (gcm convertfrom-markdowntable -ea 0)){
                                                                                                                                                                                                                                                                                                                                                        function convertFrom-MarkdownTable{
                <#
                .SYNOPSIS
                convertFrom-MarkdownTable.ps1 - Converts a Markdown table to a PowerShell object.
                .NOTES
                Version     : 1.0.3
                Author      : Todd Kadrie
                Website     : http://www.toddomation.com
                Twitter     : @tostka / http://twitter.com/tostka
                CreatedDate : 2021-06-21
                FileName    : convertFrom-MarkdownTable.ps1
                License     : MIT License
                Copyright   : (c) 2024 Todd Kadrie
                Github      : https://github.com/tostka/verb-io
                Tags        : Powershell,Markdown,Input,Conversion
                REVISION
                * 9:33 AM 4/11/2025 add alias: cfmdt (reflects standard verbalias)
                * 12:33 PM 5/17/2024 fixed odd bug, was failing to trim trailing | on some rows, which caused convertfrom-csv to drop that column.
                * 9:04 AM 9/27/2023 cbh demo output tweaks (unindented, results in 1st line de-indent and rest indented.
                * 10:35 AM 2/21/2022 CBH example ps> adds 
                * 12:42 PM 6/22/2021 bug workaround: empty fields in source md table (|data||data|) cause later (export-csv|convertto-csv) to create a csv with *missing* delimiting comma on the problem field ;  added trim of each field content, and CBH example for creation of a csv from mdtable input; added aliases
                * 5:40 PM 6/21/2021 init
                .DESCRIPTION
                convertFrom-MarkdownTable.ps1 - Converts a Markdown table to a PowerShell object.
                Also supports convesion of variant 'border' md table syntax (e.g. each line wrapped in outter pipe | chars)
                Intent is as a simpler alternative to here-stringinputs for csv building. 
                .PARAMETER markdowntext
                Markdown-formated table to be converted into an object [-markdowntext 'title text']
                .INPUTS
                Accepts piped input.
                .OUTPUTS
                System.Object[]
               .EXAMPLE
               PS> $svcs = Get-Service Bits,Winrm | select status,name,displayname | 
                  convertTo-MarkdownTable -border | ConvertFrom-MarkDownTable ;  
               Convert Service listing to and back from MD table, demo's working around border md table syntax (outter pipe-wrapped lines)
              .EXAMPLE
               PS> $mdtable = @"
            |EmailAddress|DisplayName|Groups|Ticket|
            |---|---|---|---|
            |da.pope@vatican.org||CardinalDL@vatican.org|999999|
            |bozo@clown.com|Bozo Clown|SillyDL;SmartDL|000001|
            "@ ; 
                  $of = ".\out-csv-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ; 
                  $mdtable | convertfrom-markdowntable | export-csv -path $of -notype ;
                  cat $of ;

                    "EmailAddress","DisplayName","Groups","Ticket"
                    "da.pope@vatican.org","","CardinalDL@vatican.org","999999"
                    "bozo@clown.com","Bozo Clown","SillyDL;SmartDL","000001"

                Example simpler method for building csv input files fr mdtable syntax, without PSCustomObjects, hashes, or invoked object creation.
                .EXAMPLE
                PS> $mdtable | convertFrom-MarkdownTable | convertTo-MarkdownTable -border ; 
                Example to expand and dress up a simple md table, leveraging both convertfrom-mtd and convertto-mtd (which performs space padding to align pipe columns)
                .LINK
                https://github.com/tostka/verb-IO
                #>
                [CmdletBinding()]
                [alias('convertfrom-mdt','in-markdowntable','in-mdt','cfmdt')]    
                Param (
                    [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Markdown-formated table to be converted into an object [-markdowntext 'title text']")]
                    $markdowntext
                ) ;
                PROCESS {
                    $content = @() ; 
                    if(($markdowntext|measure).count -eq 1){$markdowntext  = $markdowntext -split '\n' } ;
                    # # bug, empty fields (||) when exported (export-csv|convertto-csv) -> broken csv (missing delimiting comma between 2 fields). 
                    # workaround : flip empty field => \s. The $object still comes out properly $null on the field, but the change cause export-csv of the resulting obj to no longer output a broken csv.(weird)
                    $markdowntext  = $markdowntext -replace '\|\|','| |' ; 
                    $content = $markdowntext  | ?{$_ -notmatch "--" } ;
                } ;  
                END {
                    # trim lead/trail '| from each line (borders) ; remove empty lines; foreach
                    #$PsObj = $content.trim('|')| where-object{$_} | ForEach-Object{ 
                    # 11:19 AM 5/17/2024 issue, it's not triming trailing '|' on "THROTTLE |The message was throttled.|problem|":
                    $PsObj = $content.trim('|').trimend('|')| where-object{$_} | ForEach-Object{ 
                        #$_.split('|').trim() -join '|' ; # split fields and trim leading/trailing spaces from each , then re-join with '|'
                        # still coming through with a surviving trailing |, though the leading border is gone (causes to drop trailing cell)
                        # filter populated, trim start/end spaces, and refilter pop'd then join result - that seems to have fixed the bug
                        ($_.split('|') | where-object{$_} | foreach-object{$_.trim()} |where-object{$_} )  -join '|' ; 
                    } | ConvertFrom-Csv -Delimiter '|'; # convert to object
                    $PsObj | write-output ; 
                } ; # END-E
            } ; 
            } ; 
            #endregion CONVERTFROM_MARKDOWNTABLE ; #*------^ END convertFrom-MarkdownTable ^------
            
            $xopBuilds = @"
ProductName                                                | ReleaseDate     | BuildNumberShort | BuildNumberLong | PatchBasis                           | NickName              | IsInstallable
---------------------------------------------------------- | --------------- | ---------------- | --------------- | ------------------------------------ | --------------------- | -------------
Exchange Server SE RTM Sep25HU                             | 9/8/2025        | 15.2.2562.27     | 15.02.2562.027  | Exchange Server SE RTM               | EXSE_RTM_Sep25HU      |
Exchange Server SE RTM Aug25SU                             | 8/12/2025       | 15.2.2562.20     | 15.02.2562.020  | Exchange Server SE RTM               | EXSE_RTM_Aug25SU      |
Exchange Server SE RTM                                     | 7/1/2025        | 15.2.2562.17     | 15.02.2562.017  | Exchange Server SE RTM               | EXSE_RTM              | TRUE
Exchange Server 2019 CU15 Sep25HU                          | 9/8/2025        | 15.2.1748.37     | 15.02.1748.037  | Exchange Server 2019 CU15            | EX2019_CU15_Sep25HU   |
Exchange Server 2019 CU15 Aug25SU                          | 8/12/2025       | 15.2.1748.36     | 15.02.1748.036  | Exchange Server 2019 CU15            | EX2019_CU15_Aug25SU   |
Exchange Server 2019 CU15 May25HU                          | 5/29/2025       | 15.2.1748.26     | 15.02.1748.026  | Exchange Server 2019 CU15            | EX2019_CU15_May25HU   |
Exchange Server 2019 CU15 Apr25HU                          | 4/18/2025       | 15.2.1748.24     | 15.02.1748.024  | Exchange Server 2019 CU15            | EX2019_CU15_Apr25HU   |
Exchange Server 2019 CU15 (2025H1)                         | 2/10/2025       | 15.2.1748.10     | 15.02.1748.010  | Exchange Server 2019 CU15            | EX2019_CU15_2025H1    | TRUE
Exchange Server 2019 CU14 Sep25HU                          | 9/8/2025        | 15.2.1544.34     | 15.02.1544.034  | Exchange Server 2019 CU14            | EX2019_CU14_Sep25HU   |
Exchange Server 2019 CU14 Aug25SU                          | 8/12/2025       | 15.2.1544.33     | 15.02.1544.033  | Exchange Server 2019 CU14            | EX2019_CU14_Aug25SU   |
Exchange Server 2019 CU14 May25HU                          | 5/29/2025       | 15.2.1544.27     | 15.02.1544.027  | Exchange Server 2019 CU14            | EX2019_CU14_May25HU   |
Exchange Server 2019 CU14 Apr25HU                          | 4/18/2025       | 15.2.1544.25     | 15.02.1544.025  | Exchange Server 2019 CU14            | EX2019_CU14_Apr25HU   |
Exchange Server 2019 CU14 Nov24SUv2                        | 11/27/2024      | 15.2.1544.14     | 15.02.1544.014  | Exchange Server 2019 CU14            | EX2019_CU14_Nov24SUv2 |
Exchange Server 2019 CU14 Nov24SU                          | 11/12/2024      | 15.2.1544.13     | 15.02.1544.013  | Exchange Server 2019 CU14            | EX2019_CU14_Nov24SU   |
Exchange Server 2019 CU14 Apr24HU                          | 4/23/2024       | 15.2.1544.11     | 15.02.1544.011  | Exchange Server 2019 CU14            | EX2019_CU14_Apr24HU   |
Exchange Server 2019 CU14 Mar24SU                          | 3/12/2024       | 15.2.1544.9      | 15.02.1544.009  | Exchange Server 2019 CU14            | EX2019_CU14_Mar24SU   |
Exchange Server 2019 CU14 (2024H1)                         | 2/13/2024       | 15.2.1544.4      | 15.02.1544.004  | Exchange Server 2019 CU14            | EX2019_CU14_2024H1    | TRUE
Exchange Server 2019 CU13 Nov24SUv2                        | 11/27/2024      | 15.2.1258.39     | 15.02.1258.039  | Exchange Server 2019 CU13            | EX2019_CU13_Nov24SUv2 |
Exchange Server 2019 CU13 Nov24SU                          | 11/12/2024      | 15.2.1258.38     | 15.02.1258.038  | Exchange Server 2019 CU13            | EX2019_CU13_Nov24SU   |
Exchange Server 2019 CU13 Apr24HU                          | 4/23/2024       | 15.2.1258.34     | 15.02.1258.034  | Exchange Server 2019 CU13            | EX2019_CU13_Apr24HU   |
Exchange Server 2019 CU13 Mar24SU                          | 3/12/2024       | 15.2.1258.32     | 15.02.1258.032  | Exchange Server 2019 CU13            | EX2019_CU13_Mar24SU   |
Exchange Server 2019 CU13 Nov23SU                          | 11/14/2023      | 15.2.1258.28     | 15.02.1258.028  | Exchange Server 2019 CU13            | EX2019_CU13_Nov23SU   |
Exchange Server 2019 CU13 Oct23SU                          | 10/10/2023      | 15.2.1258.27     | 15.02.1258.027  | Exchange Server 2019 CU13            | EX2019_CU13_Oct23SU   |
Exchange Server 2019 CU13 Aug23SUv2                        | 8/15/2023       | 15.2.1258.25     | 15.02.1258.025  | Exchange Server 2019 CU13            | EX2019_CU13_Aug23SUv2 |
Exchange Server 2019 CU13 Aug23SU                          | 8/8/2023        | 15.2.1258.23     | 15.02.1258.023  | Exchange Server 2019 CU13            | EX2019_CU13_Aug23SU   |
Exchange Server 2019 CU13 Jun23SU                          | 6/13/2023       | 15.2.1258.16     | 15.02.1258.016  | Exchange Server 2019 CU13            | EX2019_CU13_Jun23SU   |
Exchange Server 2019 CU13 (2023H1)                         | 5/3/2023        | 15.2.1258.12     | 15.02.1258.012  | Exchange Server 2019 CU13            | EX2019_CU13_2023H1    | TRUE
Exchange Server 2019 CU12 Nov23SU                          | 11/14/2023      | 15.2.1118.40     | 15.02.1118.040  | Exchange Server 2019 CU12            | EX2019_CU12_Nov23SU   |
Exchange Server 2019 CU12 Oct23SU                          | 10/10/2023      | 15.2.1118.39     | 15.02.1118.039  | Exchange Server 2019 CU12            | EX2019_CU12_Oct23SU   |
Exchange Server 2019 CU12 Aug23SUv2                        | 8/15/2023       | 15.2.1118.37     | 15.02.1118.037  | Exchange Server 2019 CU12            | EX2019_CU12_Aug23SUv2 |
Exchange Server 2019 CU12 Aug23SU                          | 8/8/2023        | 15.2.1118.36     | 15.02.1118.036  | Exchange Server 2019 CU12            | EX2019_CU12_Aug23SU   |
Exchange Server 2019 CU12 Jun23SU                          | 6/13/2023       | 15.2.1118.30     | 15.02.1118.030  | Exchange Server 2019 CU12            | EX2019_CU12_Jun23SU   |
Exchange Server 2019 CU12 Mar23SU                          | 3/14/2023       | 15.2.1118.26     | 15.02.1118.026  | Exchange Server 2019 CU12            | EX2019_CU12_Mar23SU   |
Exchange Server 2019 CU12 Feb23SU                          | 2/14/2023       | 15.2.1118.25     | 15.02.1118.025  | Exchange Server 2019 CU12            | EX2019_CU12_Feb23SU   |
Exchange Server 2019 CU12 Jan23SU                          | 1/10/2023       | 15.2.1118.21     | 15.02.1118.021  | Exchange Server 2019 CU12            | EX2019_CU12_Jan23SU   |
Exchange Server 2019 CU12 Nov22SU                          | 11/8/2022       | 15.2.1118.20     | 15.02.1118.020  | Exchange Server 2019 CU12            | EX2019_CU12_Nov22SU   |
Exchange Server 2019 CU12 Oct22SU                          | 10/11/2022      | 15.2.1118.15     | 15.02.1118.015  | Exchange Server 2019 CU12            | EX2019_CU12_Oct22SU   |
Exchange Server 2019 CU12 Aug22SU                          | 8/9/2022        | 15.2.1118.12     | 15.02.1118.012  | Exchange Server 2019 CU12            | EX2019_CU12_Aug22SU   |
Exchange Server 2019 CU12 May22SU                          | 5/10/2022       | 15.2.1118.9      | 15.02.1118.009  | Exchange Server 2019 CU12            | EX2019_CU12_May22SU   |
Exchange Server 2019 CU12 (2022H1)                         | 4/20/2022       | 15.2.1118.7      | 15.02.1118.007  | Exchange Server 2019 CU12            | EX2019_CU12_2022H1    | TRUE
Exchange Server 2019 CU11 Mar23SU                          | 3/14/2023       | 15.2.986.42      | 15.02.0986.042  | Exchange Server 2019 CU11            | EX2019_CU11_Mar23SU   |
Exchange Server 2019 CU11 Feb23SU                          | 2/14/2023       | 15.2.986.41      | 15.02.0986.041  | Exchange Server 2019 CU11            | EX2019_CU11_Feb23SU   |
Exchange Server 2019 CU11 Jan23SU                          | 1/10/2023       | 15.2.986.37      | 15.02.0986.037  | Exchange Server 2019 CU11            | EX2019_CU11_Jan23SU   |
Exchange Server 2019 CU11 Nov22SU                          | 11/8/2022       | 15.2.986.36      | 15.02.0986.036  | Exchange Server 2019 CU11            | EX2019_CU11_Nov22SU   |
Exchange Server 2019 CU11 Oct22SU                          | 10/11/2022      | 15.2.986.30      | 15.02.0986.030  | Exchange Server 2019 CU11            | EX2019_CU11_Oct22SU   |
Exchange Server 2019 CU11 Aug22SU                          | 8/9/2022        | 15.2.986.29      | 15.02.0986.029  | Exchange Server 2019 CU11            | EX2019_CU11_Aug22SU   |
Exchange Server 2019 CU11 May22SU                          | 5/10/2022       | 15.2.986.26      | 15.02.0986.026  | Exchange Server 2019 CU11            | EX2019_CU11_May22SU   |
Exchange Server 2019 CU11 Mar22SU                          | 3/8/2022        | 15.2.986.22      | 15.02.0986.022  | Exchange Server 2019 CU11            | EX2019_CU11_Mar22SU   |
Exchange Server 2019 CU11 Jan22SU                          | 1/11/2022       | 15.2.986.15      | 15.02.0986.015  | Exchange Server 2019 CU11            | EX2019_CU11_Jan22SU   |
Exchange Server 2019 CU11 Nov21SU                          | 11/9/2021       | 15.2.986.14      | 15.02.0986.014  | Exchange Server 2019 CU11            | EX2019_CU11_Nov21SU   |
Exchange Server 2019 CU11 Oct21SU                          | 10/12/2021      | 15.2.986.9       | 15.02.0986.009  | Exchange Server 2019 CU11            | EX2019_CU11_Oct21SU   |
Exchange Server 2019 CU11                                  | 9/28/2021       | 15.2.986.5       | 15.02.0986.005  | Exchange Server 2019 CU11            | EX2019_CU11           | TRUE
Exchange Server 2019 CU10 Mar22SU                          | 3/8/2022        | 15.2.922.27      | 15.02.0922.027  | Exchange Server 2019 CU10            | EX2019_CU10_Mar22SU   |
Exchange Server 2019 CU10 Jan22SU                          | 1/11/2022       | 15.2.922.20      | 15.02.0922.020  | Exchange Server 2019 CU10            | EX2019_CU10_Jan22SU   |
Exchange Server 2019 CU10 Nov21SU                          | 11/9/2021       | 15.2.922.19      | 15.02.0922.019  | Exchange Server 2019 CU10            | EX2019_CU10_Nov21SU   |
Exchange Server 2019 CU10 Oct21SU                          | 10/12/2021      | 15.2.922.14      | 15.02.0922.014  | Exchange Server 2019 CU10            | EX2019_CU10_Oct21SU   |
Exchange Server 2019 CU10 Jul21SU                          | 7/13/2021       | 15.2.922.13      | 15.02.0922.013  | Exchange Server 2019 CU10            | EX2019_CU10_Jul21SU   |
Exchange Server 2019 CU10                                  | 6/29/2021       | 15.2.922.7       | 15.02.0922.007  | Exchange Server 2019 CU10            | EX2019_CU10           | TRUE
Exchange Server 2019 CU9 Jul21SU                           | 7/13/2021       | 15.2.858.15      | 15.02.0858.015  | Exchange Server 2019 CU9             | EX2019_CU9_Jul21SU    |
Exchange Server 2019 CU9 May21SU                           | 5/11/2021       | 15.2.858.12      | 15.02.0858.012  | Exchange Server 2019 CU9             | EX2019_CU9_May21SU    |
Exchange Server 2019 CU9 Apr21SU                           | 4/13/2021       | 15.2.858.10      | 15.02.0858.010  | Exchange Server 2019 CU9             | EX2019_CU9_Apr21SU    |
Exchange Server 2019 CU9                                   | 3/16/2021       | 15.2.858.5       | 15.02.0858.005  | Exchange Server 2019 CU9             | EX2019_CU9            | TRUE
Exchange Server 2019 CU8 May21SU                           | 5/11/2021       | 15.2.792.15      | 15.02.0792.015  | Exchange Server 2019 CU8             | EX2019_CU8_May21SU    |
Exchange Server 2019 CU8 Apr21SU                           | 4/13/2021       | 15.2.792.13      | 15.02.0792.013  | Exchange Server 2019 CU8             | EX2019_CU8_Apr21SU    |
Exchange Server 2019 CU8 Mar21SU                           | 3/2/2021        | 15.2.792.10      | 15.02.0792.010  | Exchange Server 2019 CU8             | EX2019_CU8_Mar21SU    |
Exchange Server 2019 CU8                                   | 12/15/2020      | 15.2.792.3       | 15.02.0792.003  | Exchange Server 2019 CU8             | EX2019_CU8            | TRUE
Exchange Server 2019 CU7 Mar21SU                           | 3/2/2021        | 15.2.721.13      | 15.02.0721.013  | Exchange Server 2019 CU7             | EX2019_CU7_Mar21SU    |
Exchange Server 2019 CU7                                   | 9/15/2020       | 15.2.721.2       | 15.02.0721.002  | Exchange Server 2019 CU7             | EX2019_CU7            | TRUE
Exchange Server 2019 CU6 Mar21SU                           | 3/2/2021        | 15.2.659.12      | 15.02.0659.012  | Exchange Server 2019 CU6             | EX2019_CU6_Mar21SU    |
Exchange Server 2019 CU6                                   | 6/16/2020       | 15.2.659.4       | 15.02.0659.004  | Exchange Server 2019 CU6             | EX2019_CU6            | TRUE
Exchange Server 2019 CU5 Mar21SU                           | 3/2/2021        | 15.2.595.8       | 15.02.0595.008  | Exchange Server 2019 CU5             | EX2019_CU5_Mar21SU    |
Exchange Server 2019 CU5                                   | 3/17/2020       | 15.2.595.3       | 15.02.0595.003  | Exchange Server 2019 CU5             | EX2019_CU5            | TRUE
Exchange Server 2019 CU4 Mar21SU                           | 3/2/2021        | 15.2.529.13      | 15.02.0529.013  | Exchange Server 2019 CU4             | EX2019_CU4_Mar21SU    |
Exchange Server 2019 CU4                                   | 12/17/2019      | 15.2.529.5       | 15.02.0529.005  | Exchange Server 2019 CU4             | EX2019_CU4            | TRUE
Exchange Server 2019 CU3 Mar21SU                           | 3/2/2021        | 15.2.464.15      | 15.02.0464.015  | Exchange Server 2019 CU3             | EX2019_CU3_Mar21SU    |
Exchange Server 2019 CU3                                   | 9/17/2019       | 15.2.464.5       | 15.02.0464.005  | Exchange Server 2019 CU3             | EX2019_CU3            | TRUE
Exchange Server 2019 CU2 Mar21SU                           | 3/2/2021        | 15.2.397.11      | 15.02.0397.011  | Exchange Server 2019 CU2             | EX2019_CU2_Mar21SU    |
Exchange Server 2019 CU2                                   | 6/18/2019       | 15.2.397.3       | 15.02.0397.003  | Exchange Server 2019 CU2             | EX2019_CU2            | TRUE
Exchange Server 2019 CU1 Mar21SU                           | 3/2/2021        | 15.2.330.11      | 15.02.0330.011  | Exchange Server 2019 CU1             | EX2019_CU1_Mar21SU    |
Exchange Server 2019 CU1                                   | 2/12/2019       | 15.2.330.5       | 15.02.0330.005  | Exchange Server 2019 CU1             | EX2019_CU1            | TRUE
Exchange Server 2019 RTM Mar21SU                           | 3/2/2021        | 15.2.221.18      | 15.02.0221.018  | Exchange Server 2019 RTM             | EX2019_RTM_Mar21SU    |
Exchange Server 2019 RTM                                   | 10/22/2018      | 15.2.221.12      | 15.02.0221.012  | Exchange Server 2019 RTM             | EX2019_RTM            | TRUE
Exchange Server 2019 Preview                               | 7/24/2018       | 15.2.196.0       | 15.02.0196.000  | Exchange Server 2019 Preview         | EX2019_Preview        |
Exchange Server 2016 CU23 Sep25HU                          | 9/8/2025        | 15.1.2507.59     | 15.01.2507.059  | Exchange Server 2016 CU23            | EX2016_CU23_Sep25HU   |
Exchange Server 2016 CU23 Aug25SU                          | 8/12/2025       | 15.1.2507.58     | 15.01.2507.058  | Exchange Server 2016 CU23            | EX2016_CU23_Aug25SU   |
Exchange Server 2016 CU23 May25HU                          | 5/29/2025       | 15.1.2507.57     | 15.01.2507.057  | Exchange Server 2016 CU23            | EX2016_CU23_May25HU   |
Exchange Server 2016 CU23 Apr25HU                          | 4/18/2025       | 15.1.2507.55     | 15.01.2507.055  | Exchange Server 2016 CU23            | EX2016_CU23_Apr25HU   |
Exchange Server 2016 CU23 Nov24SUv2                        | 11/27/2024      | 15.1.2507.44     | 15.01.2507.044  | Exchange Server 2016 CU23            | EX2016_CU23_Nov24SUv2 |
Exchange Server 2016 CU23 Nov24SU                          | 11/12/2024      | 15.1.2507.43     | 15.01.2507.043  | Exchange Server 2016 CU23            | EX2016_CU23_Nov24SU   |
Exchange Server 2016 CU23 Apr24HU                          | 4/23/2024       | 15.1.2507.39     | 15.01.2507.039  | Exchange Server 2016 CU23            | EX2016_CU23_Apr24HU   |
Exchange Server 2016 CU23 Mar24SU                          | 3/12/2024       | 15.1.2507.37     | 15.01.2507.037  | Exchange Server 2016 CU23            | EX2016_CU23_Mar24SU   |
Exchange Server 2016 CU23 Nov23SU                          | 11/14/2023      | 15.1.2507.35     | 15.01.2507.035  | Exchange Server 2016 CU23            | EX2016_CU23_Nov23SU   |
Exchange Server 2016 CU23 Oct23SU                          | 10/10/2023      | 15.1.2507.34     | 15.01.2507.034  | Exchange Server 2016 CU23            | EX2016_CU23_Oct23SU   |
Exchange Server 2016 CU23 Aug23SUv2                        | 8/15/2023       | 15.1.2507.32     | 15.01.2507.032  | Exchange Server 2016 CU23            | EX2016_CU23_Aug23SUv2 |
Exchange Server 2016 CU23 Aug23SU                          | 8/8/2023        | 15.1.2507.31     | 15.01.2507.031  | Exchange Server 2016 CU23            | EX2016_CU23_Aug23SU   |
Exchange Server 2016 CU23 Jun23SU                          | 6/13/2023       | 15.1.2507.27     | 15.01.2507.027  | Exchange Server 2016 CU23            | EX2016_CU23_Jun23SU   |
Exchange Server 2016 CU23 Mar23SU                          | 3/14/2023       | 15.1.2507.23     | 15.01.2507.023  | Exchange Server 2016 CU23            | EX2016_CU23_Mar23SU   |
Exchange Server 2016 CU23 Feb23SU                          | 2/14/2023       | 15.1.2507.21     | 15.01.2507.021  | Exchange Server 2016 CU23            | EX2016_CU23_Feb23SU   |
Exchange Server 2016 CU23 Jan23SU                          | 1/10/2023       | 15.1.2507.17     | 15.01.2507.017  | Exchange Server 2016 CU23            | EX2016_CU23_Jan23SU   |
Exchange Server 2016 CU23 Nov22SU                          | 11/8/2022       | 15.1.2507.16     | 15.01.2507.016  | Exchange Server 2016 CU23            | EX2016_CU23_Nov22SU   |
Exchange Server 2016 CU23 Oct22SU                          | 10/11/2022      | 15.1.2507.13     | 15.01.2507.013  | Exchange Server 2016 CU23            | EX2016_CU23_Oct22SU   |
Exchange Server 2016 CU23 Aug22SU                          | 8/9/2022        | 15.1.2507.12     | 15.01.2507.012  | Exchange Server 2016 CU23            | EX2016_CU23_Aug22SU   |
Exchange Server 2016 CU23 May22SU                          | 5/10/2022       | 15.1.2507.9      | 15.01.2507.009  | Exchange Server 2016 CU23            | EX2016_CU23_May22SU   |
Exchange Server 2016 CU23 (2022H1)                         | 4/20/2022       | 15.1.2507.6      | 15.01.2507.006  | Exchange Server 2016 CU23            | EX2016_CU23_2022H1    | TRUE
Exchange Server 2016 CU22 Nov22SU                          | 11/8/2022       | 15.1.2375.37     | 15.01.2375.037  | Exchange Server 2016 CU22            | EX2016_CU22_Nov22SU   |
Exchange Server 2016 CU22 Oct22SU                          | 10/11/2022      | 15.1.2375.32     | 15.01.2375.032  | Exchange Server 2016 CU22            | EX2016_CU22_Oct22SU   |
Exchange Server 2016 CU22 Aug22SU                          | 8/9/2022        | 15.1.2375.31     | 15.01.2375.031  | Exchange Server 2016 CU22            | EX2016_CU22_Aug22SU   |
Exchange Server 2016 CU22 May22SU                          | 5/10/2022       | 15.1.2375.28     | 15.01.2375.028  | Exchange Server 2016 CU22            | EX2016_CU22_May22SU   |
Exchange Server 2016 CU22 Mar22SU                          | 3/8/2022        | 15.1.2375.24     | 15.01.2375.024  | Exchange Server 2016 CU22            | EX2016_CU22_Mar22SU   |
Exchange Server 2016 CU22 Jan22SU                          | 1/11/2022       | 15.1.2375.18     | 15.01.2375.018  | Exchange Server 2016 CU22            | EX2016_CU22_Jan22SU   |
Exchange Server 2016 CU22 Nov21SU                          | 11/9/2021       | 15.1.2375.17     | 15.01.2375.017  | Exchange Server 2016 CU22            | EX2016_CU22_Nov21SU   |
Exchange Server 2016 CU22 Oct21SU                          | 10/12/2021      | 15.1.2375.12     | 15.01.2375.012  | Exchange Server 2016 CU22            | EX2016_CU22_Oct21SU   |
Exchange Server 2016 CU22                                  | 9/28/2021       | 15.1.2375.7      | 15.01.2375.007  | Exchange Server 2016 CU22            | EX2016_CU22           | TRUE
Exchange Server 2016 CU21 Mar22SU                          | 3/8/2022        | 15.1.2308.27     | 15.01.2308.027  | Exchange Server 2016 CU21            | EX2016_CU21_Mar22SU   |
Exchange Server 2016 CU21 Jan22SU                          | 1/11/2022       | 15.1.2308.21     | 15.01.2308.021  | Exchange Server 2016 CU21            | EX2016_CU21_Jan22SU   |
Exchange Server 2016 CU21 Nov21SU                          | 11/9/2021       | 15.1.2308.20     | 15.01.2308.020  | Exchange Server 2016 CU21            | EX2016_CU21_Nov21SU   |
Exchange Server 2016 CU21 Oct21SU                          | 10/12/2021      | 15.1.2308.15     | 15.01.2308.015  | Exchange Server 2016 CU21            | EX2016_CU21_Oct21SU   |
Exchange Server 2016 CU21 Jul21SU                          | 7/13/2021       | 15.1.2308.14     | 15.01.2308.014  | Exchange Server 2016 CU21            | EX2016_CU21_Jul21SU   |
Exchange Server 2016 CU21                                  | 6/29/2021       | 15.1.2308.8      | 15.01.2308.008  | Exchange Server 2016 CU21            | EX2016_CU21           | TRUE
Exchange Server 2016 CU20 Jul21SU                          | 7/13/2021       | 15.1.2242.12     | 15.01.2242.012  | Exchange Server 2016 CU20            | EX2016_CU20_Jul21SU   |
Exchange Server 2016 CU20 May21SU                          | 5/11/2021       | 15.1.2242.10     | 15.01.2242.010  | Exchange Server 2016 CU20            | EX2016_CU20_May21SU   |
Exchange Server 2016 CU20 Apr21SU                          | 4/13/2021       | 15.1.2242.8      | 15.01.2242.008  | Exchange Server 2016 CU20            | EX2016_CU20_Apr21SU   |
Exchange Server 2016 CU20                                  | 3/16/2021       | 15.1.2242.4      | 15.01.2242.004  | Exchange Server 2016 CU20            | EX2016_CU20           | TRUE
Exchange Server 2016 CU19 May21SU                          | 5/11/2021       | 15.1.2176.14     | 15.01.2176.014  | Exchange Server 2016 CU19            | EX2016_CU19_May21SU   |
Exchange Server 2016 CU19 Apr21SU                          | 4/13/2021       | 15.1.2176.12     | 15.01.2176.012  | Exchange Server 2016 CU19            | EX2016_CU19_Apr21SU   |
Exchange Server 2016 CU19 Mar21SU                          | 3/2/2021        | 15.1.2176.9      | 15.01.2176.009  | Exchange Server 2016 CU19            | EX2016_CU19_Mar21SU   |
Exchange Server 2016 CU19                                  | 12/15/2020      | 15.1.2176.2      | 15.01.2176.002  | Exchange Server 2016 CU19            | EX2016_CU19           | TRUE
Exchange Server 2016 CU18 Mar21SU                          | 3/2/2021        | 15.1.2106.13     | 15.01.2106.013  | Exchange Server 2016 CU18            | EX2016_CU18_Mar21SU   |
Exchange Server 2016 CU18                                  | 9/15/2020       | 15.1.2106.2      | 15.01.2106.002  | Exchange Server 2016 CU18            | EX2016_CU18           | TRUE
Exchange Server 2016 CU17 Mar21SU                          | 3/2/2021        | 15.1.2044.13     | 15.01.2044.013  | Exchange Server 2016 CU17            | EX2016_CU17_Mar21SU   |
Exchange Server 2016 CU17                                  | 6/16/2020       | 15.1.2044.4      | 15.01.2044.004  | Exchange Server 2016 CU17            | EX2016_CU17           | TRUE
Exchange Server 2016 CU16 Mar21SU                          | 3/2/2021        | 15.1.1979.8      | 15.01.1979.008  | Exchange Server 2016 CU16            | EX2016_CU16_Mar21SU   |
Exchange Server 2016 CU16                                  | 3/17/2020       | 15.1.1979.3      | 15.01.1979.003  | Exchange Server 2016 CU16            | EX2016_CU16           | TRUE
Exchange Server 2016 CU15 Mar21SU                          | 3/2/2021        | 15.1.1913.12     | 15.01.1913.012  | Exchange Server 2016 CU15            | EX2016_CU15_Mar21SU   |
Exchange Server 2016 CU15                                  | 12/17/2019      | 15.1.1913.5      | 15.01.1913.005  | Exchange Server 2016 CU15            | EX2016_CU15           | TRUE
Exchange Server 2016 CU14 Mar21SU                          | 3/2/2021        | 15.1.1847.12     | 15.01.1847.012  | Exchange Server 2016 CU14            | EX2016_CU14_Mar21SU   |
Exchange Server 2016 CU14                                  | 9/17/2019       | 15.1.1847.3      | 15.01.1847.003  | Exchange Server 2016 CU14            | EX2016_CU14           | TRUE
Exchange Server 2016 CU13 Mar21SU                          | 3/2/2021        | 15.1.1779.8      | 15.01.1779.008  | Exchange Server 2016 CU13            | EX2016_CU13_Mar21SU   |
Exchange Server 2016 CU13                                  | 6/18/2019       | 15.1.1779.2      | 15.01.1779.002  | Exchange Server 2016 CU13            | EX2016_CU13           | TRUE
Exchange Server 2016 CU12 Mar21SU                          | 3/2/2021        | 15.1.1713.10     | 15.01.1713.010  | Exchange Server 2016 CU12            | EX2016_CU12_Mar21SU   |
Exchange Server 2016 CU12                                  | 2/12/2019       | 15.1.1713.5      | 15.01.1713.005  | Exchange Server 2016 CU12            | EX2016_CU12           | TRUE
Exchange Server 2016 CU11 Mar21SU                          | 3/2/2021        | 15.1.1591.18     | 15.01.1591.018  | Exchange Server 2016 CU11            | EX2016_CU11_Mar21SU   |
Exchange Server 2016 CU11                                  | 10/16/2018      | 15.1.1591.10     | 15.01.1591.010  | Exchange Server 2016 CU11            | EX2016_CU11           | TRUE
Exchange Server 2016 CU10 Mar21SU                          | 3/2/2021        | 15.1.1531.12     | 15.01.1531.012  | Exchange Server 2016 CU10            | EX2016_CU10_Mar21SU   |
Exchange Server 2016 CU10                                  | 6/19/2018       | 15.1.1531.3      | 15.01.1531.003  | Exchange Server 2016 CU10            | EX2016_CU10           | TRUE
Exchange Server 2016 CU9 Mar21SU                           | 3/2/2021        | 15.1.1466.16     | 15.01.1466.016  | Exchange Server 2016 CU9             | EX2016_CU9_Mar21SU    |
Exchange Server 2016 CU9                                   | 3/20/2018       | 15.1.1466.3      | 15.01.1466.003  | Exchange Server 2016 CU9             | EX2016_CU9            | TRUE
Exchange Server 2016 CU8 Mar21SU                           | 3/2/2021        | 15.1.1415.10     | 15.01.1415.010  | Exchange Server 2016 CU8             | EX2016_CU8_Mar21SU    |
Exchange Server 2016 CU8                                   | 12/19/2017      | 15.1.1415.2      | 15.01.1415.002  | Exchange Server 2016 CU8             | EX2016_CU8            | TRUE
Exchange Server 2016 CU7                                   | 9/19/2017       | 15.1.1261.35     | 15.01.1261.035  | Exchange Server 2016 CU7             | EX2016_CU7            | TRUE
Exchange Server 2016 CU6                                   | 6/27/2017       | 15.1.1034.26     | 15.01.1034.026  | Exchange Server 2016 CU6             | EX2016_CU6            | TRUE
Exchange Server 2016 CU5                                   | 3/21/2017       | 15.1.845.34      | 15.01.0845.034  | Exchange Server 2016 CU5             | EX2016_CU5            | TRUE
Exchange Server 2016 CU4                                   | 12/13/2016      | 15.1.669.32      | 15.01.0669.032  | Exchange Server 2016 CU4             | EX2016_CU4            | TRUE
Exchange Server 2016 CU3                                   | 9/20/2016       | 15.1.544.27      | 15.01.0544.027  | Exchange Server 2016 CU3             | EX2016_CU3            | TRUE
Exchange Server 2016 CU2                                   | 6/21/2016       | 15.1.466.34      | 15.01.0466.034  | Exchange Server 2016 CU2             | EX2016_CU2            | TRUE
Exchange Server 2016 CU1                                   | 3/15/2016       | 15.1.396.30      | 15.01.0396.030  | Exchange Server 2016 CU1             | EX2016_CU1            | TRUE
Exchange Server 2016 RTM                                   | 10/1/2015       | 15.1.225.42      | 15.01.0225.042  | Exchange Server 2016 RTM             | EX2016_RTM            | TRUE
Exchange Server 2016 Preview                               | 7/22/2015       | 15.1.225.16      | 15.01.0225.016  | Exchange Server 2016 Preview         | EX2016_Preview        |
Exchange Server 2013 CU23 Mar23SU                          | 3/14/2023       | 15.0.1497.48     | 15.00.1497.048  | Exchange Server 2013 CU23            | EX2013_CU23_Mar23SU   |
Exchange Server 2013 CU23 Feb23SU                          | 2/14/2023       | 15.0.1497.47     | 15.00.1497.047  | Exchange Server 2013 CU23            | EX2013_CU23_Feb23SU   |
Exchange Server 2013 CU23 Jan23SU                          | 1/10/2023       | 15.0.1497.45     | 15.00.1497.045  | Exchange Server 2013 CU23            | EX2013_CU23_Jan23SU   |
Exchange Server 2013 CU23 Nov22SU                          | 11/8/2022       | 15.0.1497.44     | 15.00.1497.044  | Exchange Server 2013 CU23            | EX2013_CU23_Nov22SU   |
Exchange Server 2013 CU23 Oct22SU                          | 10/11/2022      | 15.0.1497.42     | 15.00.1497.042  | Exchange Server 2013 CU23            | EX2013_CU23_Oct22SU   |
Exchange Server 2013 CU23 Aug22SU                          | 8/9/2022        | 15.0.1497.40     | 15.00.1497.040  | Exchange Server 2013 CU23            | EX2013_CU23_Aug22SU   |
Exchange Server 2013 CU23 May22SU                          | 5/10/2022       | 15.0.1497.36     | 15.00.1497.036  | Exchange Server 2013 CU23            | EX2013_CU23_May22SU   |
Exchange Server 2013 CU23 Mar22SU                          | 3/8/2022        | 15.0.1497.33     | 15.00.1497.033  | Exchange Server 2013 CU23            | EX2013_CU23_Mar22SU   |
Exchange Server 2013 CU23 Jan22SU                          | 1/11/2022       | 15.0.1497.28     | 15.00.1497.028  | Exchange Server 2013 CU23            | EX2013_CU23_Jan22SU   |
Exchange Server 2013 CU23 Nov21SU                          | 11/9/2021       | 15.0.1497.26     | 15.00.1497.026  | Exchange Server 2013 CU23            | EX2013_CU23_Nov21SU   |
Exchange Server 2013 CU23 Oct21SU                          | 10/12/2021      | 15.0.1497.24     | 15.00.1497.024  | Exchange Server 2013 CU23            | EX2013_CU23_Oct21SU   |
Exchange Server 2013 CU23 Jul21SU                          | 7/13/2021       | 15.0.1497.23     | 15.00.1497.023  | Exchange Server 2013 CU23            | EX2013_CU23_Jul21SU   |
Exchange Server 2013 CU23 May21SU                          | 5/11/2021       | 15.0.1497.18     | 15.00.1497.018  | Exchange Server 2013 CU23            | EX2013_CU23_May21SU   |
Exchange Server 2013 CU23 Apr21SU                          | 4/13/2021       | 15.0.1497.15     | 15.00.1497.015  | Exchange Server 2013 CU23            | EX2013_CU23_Apr21SU   |
Exchange Server 2013 CU23 Mar21SU                          | 3/2/2021        | 15.0.1497.12     | 15.00.1497.012  | Exchange Server 2013 CU23            | EX2013_CU23_Mar21SU   |
Exchange Server 2013 CU23                                  | 6/18/2019       | 15.0.1497.2      | 15.00.1497.002  | Exchange Server 2013 CU23            | EX2013_CU23           | TRUE
Exchange Server 2013 CU22 Mar21SU                          | 3/2/2021        | 15.0.1473.6      | 15.00.1473.006  | Exchange Server 2013 CU22            | EX2013_CU22_Mar21SU   |
Exchange Server 2013 CU22                                  | 2/12/2019       | 15.0.1473.3      | 15.00.1473.003  | Exchange Server 2013 CU22            | EX2013_CU22           | TRUE
Exchange Server 2013 CU21 Mar21SU                          | 3/2/2021        | 15.0.1395.12     | 15.00.1395.012  | Exchange Server 2013 CU21            | EX2013_CU21_Mar21SU   |
Exchange Server 2013 CU21                                  | 6/19/2018       | 15.0.1395.4      | 15.00.1395.004  | Exchange Server 2013 CU21            | EX2013_CU21           | TRUE
Exchange Server 2013 CU20                                  | 3/20/2018       | 15.0.1367.3      | 15.00.1367.003  | Exchange Server 2013 CU20            | EX2013_CU20           | TRUE
Exchange Server 2013 CU19                                  | 12/19/2017      | 15.0.1365.1      | 15.00.1365.001  | Exchange Server 2013 CU19            | EX2013_CU19           | TRUE
Exchange Server 2013 CU18                                  | 9/19/2017       | 15.0.1347.2      | 15.00.1347.002  | Exchange Server 2013 CU18            | EX2013_CU18           | TRUE
Exchange Server 2013 CU17                                  | 6/27/2017       | 15.0.1320.4      | 15.00.1320.004  | Exchange Server 2013 CU17            | EX2013_CU17           | TRUE
Exchange Server 2013 CU16                                  | 3/21/2017       | 15.0.1293.2      | 15.00.1293.002  | Exchange Server 2013 CU16            | EX2013_CU16           | TRUE
Exchange Server 2013 CU15                                  | 12/13/2016      | 15.0.1263.5      | 15.00.1263.005  | Exchange Server 2013 CU15            | EX2013_CU15           | TRUE
Exchange Server 2013 CU14                                  | 9/20/2016       | 15.0.1236.3      | 15.00.1236.003  | Exchange Server 2013 CU14            | EX2013_CU14           | TRUE
Exchange Server 2013 CU13                                  | 6/21/2016       | 15.0.1210.3      | 15.00.1210.003  | Exchange Server 2013 CU13            | EX2013_CU13           | TRUE
Exchange Server 2013 CU12                                  | 3/15/2016       | 15.0.1178.4      | 15.00.1178.004  | Exchange Server 2013 CU12            | EX2013_CU12           | TRUE
Exchange Server 2013 CU11                                  | 12/15/2015      | 15.0.1156.6      | 15.00.1156.006  | Exchange Server 2013 CU11            | EX2013_CU11           | TRUE
Exchange Server 2013 CU10                                  | 9/15/2015       | 15.0.1130.7      | 15.00.1130.007  | Exchange Server 2013 CU10            | EX2013_CU10           | TRUE
Exchange Server 2013 CU9                                   | 6/17/2015       | 15.0.1104.5      | 15.00.1104.005  | Exchange Server 2013 CU9             | EX2013_CU9            | TRUE
Exchange Server 2013 CU8                                   | 3/17/2015       | 15.0.1076.9      | 15.00.1076.009  | Exchange Server 2013 CU8             | EX2013_CU8            | TRUE
Exchange Server 2013 CU7                                   | 12/9/2014       | 15.0.1044.25     | 15.00.1044.025  | Exchange Server 2013 CU7             | EX2013_CU7            | TRUE
Exchange Server 2013 CU6                                   | 8/26/2014       | 15.0.995.29      | 15.00.0995.029  | Exchange Server 2013 CU6             | EX2013_CU6            | TRUE
Exchange Server 2013 CU5                                   | 5/27/2014       | 15.0.913.22      | 15.00.0913.022  | Exchange Server 2013 CU5             | EX2013_CU5            | TRUE
Exchange Server 2013 SP1 Mar21SU                           | 3/2/2021        | 15.0.847.64      | 15.00.0847.064  | Exchange Server 2013 SP1             | EX2013_SP1_Mar21SU    |
Exchange Server 2013 SP1                                   | 2/25/2014       | 15.0.847.32      | 15.00.0847.032  | Exchange Server 2013 SP1             | EX2013_SP1            | TRUE
Exchange Server 2013 CU3                                   | 11/25/2013      | 15.0.775.38      | 15.00.0775.038  | Exchange Server 2013 CU3             | EX2013_CU3            | TRUE
Exchange Server 2013 CU2                                   | 7/9/2013        | 15.0.712.24      | 15.00.0712.024  | Exchange Server 2013 CU2             | EX2013_CU2            | TRUE
Exchange Server 2013 CU1                                   | 4/2/2013        | 15.0.620.29      | 15.00.0620.029  | Exchange Server 2013 CU1             | EX2013_CU1            | TRUE
Exchange Server 2013 RTM                                   | 12/3/2012       | 15.0.516.32      | 15.00.0516.032  | Exchange Server 2013 RTM             | EX2013_RTM            | TRUE
Update Rollup 32 for Exchange Server 2010 SP3              | 3/2/2021        | 14.3.513.0       | 14.03.0513.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR32       |
Update Rollup 31 for Exchange Server 2010 SP3              | 12/1/2020       | 14.3.509.0       | 14.03.0509.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR31       |
Update Rollup 30 for Exchange Server 2010 SP3              | 2/11/2020       | 14.3.496.0       | 14.03.0496.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR30       |
Update Rollup 29 for Exchange Server 2010 SP3              | 7/9/2019        | 14.3.468.0       | 14.03.0468.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR29       |
Update Rollup 28 for Exchange Server 2010 SP3              | 6/7/2019        | 14.3.461.1       | 14.03.0461.001  | Exchange Server 2010 SP3             | EX2010_SP3_UR28       |
Update Rollup 27 for Exchange Server 2010 SP3              | 4/9/2019        | 14.3.452.0       | 14.03.0452.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR27       |
Update Rollup 26 for Exchange Server 2010 SP3              | 2/12/2019       | 14.3.442.0       | 14.03.0442.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR26       |
Update Rollup 25 for Exchange Server 2010 SP3              | 1/8/2019        | 14.3.435.0       | 14.03.0435.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR25       |
Update Rollup 24 for Exchange Server 2010 SP3              | 9/5/2018        | 14.3.419.0       | 14.03.0419.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR24       |
Update Rollup 23 for Exchange Server 2010 SP3              | 8/13/2018       | 14.3.417.1       | 14.03.0417.001  | Exchange Server 2010 SP3             | EX2010_SP3_UR23       |
Update Rollup 22 for Exchange Server 2010 SP3              | 6/19/2018       | 14.3.411.0       | 14.03.0411.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR22       |
Update Rollup 21 for Exchange Server 2010 SP3              | 5/7/2018        | 14.3.399.2       | 14.03.0399.002  | Exchange Server 2010 SP3             | EX2010_SP3_UR21       |
Update Rollup 20 for Exchange Server 2010 SP3              | 3/5/2018        | 14.3.389.1       | 14.03.0389.001  | Exchange Server 2010 SP3             | EX2010_SP3_UR20       |
Update Rollup 19 for Exchange Server 2010 SP3              | 12/19/2017      | 14.3.382.0       | 14.03.0382.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR19       |
Update Rollup 18 for Exchange Server 2010 SP3              | 7/11/2017       | 14.3.361.1       | 14.03.0361.001  | Exchange Server 2010 SP3             | EX2010_SP3_UR18       |
Update Rollup 17 for Exchange Server 2010 SP3              | 3/21/2017       | 14.3.352.0       | 14.03.0352.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR17       |
Update Rollup 16 for Exchange Server 2010 SP3              | 12/13/2016      | 14.3.336.0       | 14.03.0336.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR16       |
Update Rollup 15 for Exchange Server 2010 SP3              | 9/20/2016       | 14.3.319.2       | 14.03.0319.002  | Exchange Server 2010 SP3             | EX2010_SP3_UR15       |
Update Rollup 14 for Exchange Server 2010 SP3              | 6/21/2016       | 14.3.301.0       | 14.03.0301.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR14       |
Update Rollup 13 for Exchange Server 2010 SP3              | 3/15/2016       | 14.3.294.0       | 14.03.0294.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR13       |
Update Rollup 12 for Exchange Server 2010 SP3              | 12/15/2015      | 14.3.279.2       | 14.03.0279.002  | Exchange Server 2010 SP3             | EX2010_SP3_UR12       |
Update Rollup 11 for Exchange Server 2010 SP3              | 9/15/2015       | 14.3.266.2       | 14.03.0266.002  | Exchange Server 2010 SP3             | EX2010_SP3_UR11       |
Update Rollup 10 for Exchange Server 2010 SP3              | 6/17/2015       | 14.3.248.2       | 14.03.0248.002  | Exchange Server 2010 SP3             | EX2010_SP3_UR10       |
Update Rollup 9 for Exchange Server 2010 SP3               | 3/17/2015       | 14.3.235.1       | 14.03.0235.001  | Exchange Server 2010 SP3             | EX2010_SP3_UR9        |
Update Rollup 8 v2 for Exchange Server 2010 SP3            | 12/12/2014      | 14.3.224.2       | 14.03.0224.002  | Exchange Server 2010 SP3             | EX2010_SP3_UR8_v2     |
Update Rollup 8 v1 for Exchange Server 2010 SP3 (recalled) | 12/9/2014       | 14.3.224.1       | 14.03.0224.001  | Exchange Server 2010 SP3             | EX2010_SP3_UR8_v1     |
Update Rollup 7 for Exchange Server 2010 SP3               | 8/26/2014       | 14.3.210.2       | 14.03.0210.002  | Exchange Server 2010 SP3             | EX2010_SP3_UR7        |
Update Rollup 6 for Exchange Server 2010 SP3               | 5/27/2014       | 14.3.195.1       | 14.03.0195.001  | Exchange Server 2010 SP3             | EX2010_SP3_UR6        |
Update Rollup 5 for Exchange Server 2010 SP3               | 2/24/2014       | 14.3.181.6       | 14.03.0181.006  | Exchange Server 2010 SP3             | EX2010_SP3_UR5        |
Update Rollup 4 for Exchange Server 2010 SP3               | 12/9/2013       | 14.3.174.1       | 14.03.0174.001  | Exchange Server 2010 SP3             | EX2010_SP3_UR4        |
Update Rollup 3 for Exchange Server 2010 SP3               | 11/25/2013      | 14.3.169.1       | 14.03.0169.001  | Exchange Server 2010 SP3             | EX2010_SP3_UR3        |
Update Rollup 2 for Exchange Server 2010 SP3               | 8/8/2013        | 14.3.158.1       | 14.03.0158.001  | Exchange Server 2010 SP3             | EX2010_SP3_UR2        |
Update Rollup 1 for Exchange Server 2010 SP3               | 5/29/2013       | 14.3.146.0       | 14.03.0146.000  | Exchange Server 2010 SP3             | EX2010_SP3_UR1        |
Exchange Server 2010 SP3                                   | 2/12/2013       | 14.3.123.4       | 14.03.0123.004  | Exchange Server 2010 SP3             | EX2010_SP3            | TRUE
Update Rollup 8 for Exchange Server 2010 SP2               | 12/9/2013       | 14.2.390.3       | 14.02.0390.003  | Exchange Server 2010 SP2             | EX2010_SP2_UR8        |
Update Rollup 7 for Exchange Server 2010 SP2               | 8/3/2013        | 14.2.375.0       | 14.02.0375.000  | Exchange Server 2010 SP2             | EX2010_SP2_UR7        |
Update Rollup 6 Exchange Server 2010 SP2                   | 2/12/2013       | 14.2.342.3       | 14.02.0342.003  | Exchange Server 2010 SP2             | EX2010_SP2_UR6        |
Update Rollup 5 v2 for Exchange Server 2010 SP2            | 12/10/2012      | 14.2.328.10      | 14.02.0328.010  | Exchange Server 2010 SP2             | EX2010_SP2_UR5_v2     |
Update Rollup 5 for Exchange Server 2010 SP2               | 11/13/2012      | 14.3.328.5       | 14.03.0328.005  | Exchange Server 2010 SP2             | EX2010_SP2_UR5        |
Update Rollup 4 v2 for Exchange Server 2010 SP2            | 10/9/2012       | 14.2.318.4       | 14.02.0318.004  | Exchange Server 2010 SP2             | EX2010_SP2_UR4_v2     |
Update Rollup 4 for Exchange Server 2010 SP2               | 8/13/2012       | 14.2.318.2       | 14.02.0318.002  | Exchange Server 2010 SP2             | EX2010_SP2_UR4        |
Update Rollup 3 for Exchange Server 2010 SP2               | 5/29/2012       | 14.2.309.2       | 14.02.0309.002  | Exchange Server 2010 SP2             | EX2010_SP2_UR3        |
Update Rollup 2 for Exchange Server 2010 SP2               | 4/16/2012       | 14.2.298.4       | 14.02.0298.004  | Exchange Server 2010 SP2             | EX2010_SP2_UR2        |
Update Rollup 1 for Exchange Server 2010 SP2               | 2/13/2012       | 14.2.283.3       | 14.02.0283.003  | Exchange Server 2010 SP2             | EX2010_SP2_UR1        |
Exchange Server 2010 SP2                                   | 12/4/2011       | 14.2.247.5       | 14.02.0247.005  | Exchange Server 2010 SP2             | EX2010_SP2            | TRUE
Update Rollup 8 for Exchange Server 2010 SP1               | 12/10/2012      | 14.1.438.0       | 14.01.0438.000  | Exchange Server 2010 SP1             | EX2010_SP1_UR8        |
Update Rollup 7 v3 for Exchange Server 2010 SP1            | 11/13/2012      | 14.1.421.3       | 14.01.0421.003  | Exchange Server 2010 SP1             | EX2010_SP1_UR7_v3     |
Update Rollup 7 v2 for Exchange Server 2010 SP1            | 10/10/2012      | 14.1.421.2       | 14.01.0421.002  | Exchange Server 2010 SP1             | EX2010_SP1_UR7_v2     |
Update Rollup 7 for Exchange Server 2010 SP1               | 8/8/2012        | 14.1.421.0       | 14.01.0421.000  | Exchange Server 2010 SP1             | EX2010_SP1_UR7        |
Update Rollup 6 for Exchange Server 2010 SP1               | 10/27/2011      | 14.1.355.2       | 14.01.0355.002  | Exchange Server 2010 SP1             | EX2010_SP1_UR6        |
Update Rollup 5 for Exchange Server 2010 SP1               | 8/23/2011       | 14.1.339.1       | 14.01.0339.001  | Exchange Server 2010 SP1             | EX2010_SP1_UR5        |
Update Rollup 4 for Exchange Server 2010 SP1               | 7/27/2011       | 14.1.323.6       | 14.01.0323.006  | Exchange Server 2010 SP1             | EX2010_SP1_UR4        |
Update Rollup 3 for Exchange Server 2010 SP1               | 4/6/2011        | 14.1.289.7       | 14.01.0289.007  | Exchange Server 2010 SP1             | EX2010_SP1_UR3        |
Update Rollup 2 for Exchange Server 2010 SP1               | 12/9/2010       | 14.1.270.1       | 14.01.0270.001  | Exchange Server 2010 SP1             | EX2010_SP1_UR2        |
Update Rollup 1 for Exchange Server 2010 SP1               | 10/4/2010       | 14.1.255.2       | 14.01.0255.002  | Exchange Server 2010 SP1             | EX2010_SP1_UR1        |
Exchange Server 2010 SP1                                   | 8/23/2010       | 14.1.218.15      | 14.01.0218.015  | Exchange Server 2010 SP1             | EX2010_SP1            | TRUE
Update Rollup 5 for Exchange Server 2010                   | 12/13/2010      | 14.0.726.0       | 14.00.0726.000  | Exchange Server 2010 RTM             | EX2010_UR5            |
Update Rollup 4 for Exchange Server 2010                   | 6/10/2010       | 14.0.702.1       | 14.00.0702.001  | Exchange Server 2010 RTM             | EX2010_UR4            |
Update Rollup 3 for Exchange Server 2010                   | 4/13/2010       | 14.0.694.0       | 14.00.0694.000  | Exchange Server 2010 RTM             | EX2010_UR3            |
Update Rollup 2 for Exchange Server 2010                   | 3/4/2010        | 14.0.689.0       | 14.00.0689.000  | Exchange Server 2010 RTM             | EX2010_UR2            |
Update Rollup 1 for Exchange Server 2010                   | 12/9/2009       | 14.0.682.1       | 14.00.0682.001  | Exchange Server 2010 RTM             | EX2010_UR1            |
Exchange Server 2010 RTM                                   | 11/9/2009       | 14.0.639.21      | 14.00.0639.021  | Exchange Server 2010 RTM             | EX2010_RTM            | TRUE
Update Rollup 23 for Exchange Server 2007 SP3              | 3/21/2017       | 8.3.517.0        | 8.03.0517.000   | Exchange Server 2007 SP3             | EX2007_SP3_UR23       |
Update Rollup 22 for Exchange Server 2007 SP3              | 12/13/2016      | 8.3.502.0        | 8.03.0502.000   | Exchange Server 2007 SP3             | EX2007_SP3_UR22       |
Update Rollup 21 for Exchange Server 2007 SP3              | 9/20/2016       | 8.3.485.1        | 8.03.0485.001   | Exchange Server 2007 SP3             | EX2007_SP3_UR21       |
Update Rollup 20 for Exchange Server 2007 SP3              | 6/21/2016       | 8.3.468.0        | 8.03.0468.000   | Exchange Server 2007 SP3             | EX2007_SP3_UR20       |
Update Rollup 19 forExchange Server 2007 SP3               | 3/15/2016       | 8.3.459.0        | 8.03.0459.000   | Exchange Server 2007 SP3             | EX2007_SP3_UR19       |
Update Rollup 18 forExchange Server 2007 SP3               | December, 2015  | 8.3.445.0        | 8.03.0445.000   | Exchange Server 2007 SP3             | EX2007_SP3_UR18       |
Update Rollup 17 forExchange Server 2007 SP3               | 6/17/2015       | 8.3.417.1        | 8.03.0417.001   | Exchange Server 2007 SP3             | EX2007_SP3_UR17       |
Update Rollup 16 for Exchange Server 2007 SP3              | 3/17/2015       | 8.3.406.0        | 8.03.0406.000   | Exchange Server 2007 SP3             | EX2007_SP3_UR16       |
Update Rollup 15 for Exchange Server 2007 SP3              | 12/9/2014       | 8.3.389.2        | 8.03.0389.002   | Exchange Server 2007 SP3             | EX2007_SP3_UR15       |
Update Rollup 14 for Exchange Server 2007 SP3              | 8/26/2014       | 8.3.379.2        | 8.03.0379.002   | Exchange Server 2007 SP3             | EX2007_SP3_UR14       |
Update Rollup 13 for Exchange Server 2007 SP3              | 2/24/2014       | 8.3.348.2        | 8.03.0348.002   | Exchange Server 2007 SP3             | EX2007_SP3_UR13       |
Update Rollup 12 for Exchange Server 2007 SP3              | 12/9/2013       | 8.3.342.4        | 8.03.0342.004   | Exchange Server 2007 SP3             | EX2007_SP3_UR12       |
Update Rollup 11 for Exchange Server 2007 SP3              | 8/13/2013       | 8.3.327.1        | 8.03.0327.001   | Exchange Server 2007 SP3             | EX2007_SP3_UR11       |
Update Rollup 10 for Exchange Server 2007 SP3              | 2/11/2013       | 8.3.298.3        | 8.03.0298.003   | Exchange Server 2007 SP3             | EX2007_SP3_UR10       |
Update Rollup 9 for Exchange Server 2007 SP3               | 12/10/2012      | 8.3.297.2        | 8.03.0297.002   | Exchange Server 2007 SP3             | EX2007_SP3_UR9        |
Update Rollup 8-v3 for Exchange Server 2007 SP3            | 11/13/2012      | 8.3.279.6        | 8.03.0279.006   | Exchange Server 2007 SP3             | EX2007_SP3_UR8-v3     |
Update Rollup 8-v2 for Exchange Server 2007 SP3            | 10/9/2012       | 8.3.279.5        | 8.03.0279.005   | Exchange Server 2007 SP3             | EX2007_SP3_UR8-v2     |
Update Rollup 8 for Exchange Server 2007 SP3               | 8/13/2012       | 8.3.279.3        | 8.03.0279.003   | Exchange Server 2007 SP3             | EX2007_SP3_UR8        |
Update Rollup 7 for Exchange Server 2007 SP3               | 4/16/2012       | 8.3.264.0        | 8.03.0264.000   | Exchange Server 2007 SP3             | EX2007_SP3_UR7        |
Update Rollup 6 for Exchange Server 2007 SP3               | 1/26/2012       | 8.3.245.2        | 8.03.0245.002   | Exchange Server 2007 SP3             | EX2007_SP3_UR6        |
Update Rollup 5 for Exchange Server 2007 SP3               | 9/21/2011       | 8.3.213.1        | 8.03.0213.001   | Exchange Server 2007 SP3             | EX2007_SP3_UR5        |
Update Rollup 4 for Exchange Server 2007 SP3               | 5/28/2011       | 8.3.192.1        | 8.03.0192.001   | Exchange Server 2007 SP3             | EX2007_SP3_UR4        |
Update Rollup 3-v2 for Exchange Server 2007 SP3            | 3/30/2011       | 8.3.159.2        | 8.03.0159.002   | Exchange Server 2007 SP3             | EX2007_SP3_UR3-v2     |
Update Rollup 2 for Exchange Server 2007 SP3               | 12/10/2010      | 8.3.137.3        | 8.03.0137.003   | Exchange Server 2007 SP3             | EX2007_SP3_UR2        |
Update Rollup 1 for Exchange Server 2007 SP3               | 9/9/2010        | 8.3.106.2        | 8.03.0106.002   | Exchange Server 2007 SP3             | EX2007_SP3_UR1        |
Exchange Server 2007 SP3                                   | 6/7/2010        | 8.3.83.6         | 8.03.0083.006   | Exchange Server 2007 SP3             | EX2007_SP3            | TRUE
Update Rollup 5 for Exchange Server 2007 SP2               | 12/7/2010       | 8.2.305.3        | 8.02.0305.003   | Exchange Server 2007 SP2             | EX2007_SP2_UR5        |
Update Rollup 4 for Exchange Server 2007 SP2               | 4/9/2010        | 8.2.254.0        | 8.02.0254.000   | Exchange Server 2007 SP2             | EX2007_SP2_UR4        |
Update Rollup 3 for Exchange Server 2007 SP2               | 3/17/2010       | 8.2.247.2        | 8.02.0247.002   | Exchange Server 2007 SP2             | EX2007_SP2_UR3        |
Update Rollup 2 for Exchange Server 2007 SP2               | 1/22/2010       | 8.2.234.1        | 8.02.0234.001   | Exchange Server 2007 SP2             | EX2007_SP2_UR2        |
Update Rollup 1 for Exchange Server 2007 SP2               | 11/19/2009      | 8.2.217.3        | 8.02.0217.003   | Exchange Server 2007 SP2             | EX2007_SP2_UR1        |
Exchange Server 2007 SP2                                   | 8/24/2009       | 8.2.176.2        | 8.02.0176.002   | Exchange Server 2007 SP2             | EX2007_SP2            | TRUE
Update Rollup 10 for Exchange Server 2007 SP1              | 4/13/2010       | 8.1.436.0        | 8.01.0436.000   | Exchange Server 2007 SP1             | EX2007_SP1_UR10       |
Update Rollup 9 for Exchange Server 2007 SP1               | 7/16/2009       | 8.1.393.1        | 8.01.0393.001   | Exchange Server 2007 SP1             | EX2007_SP1_UR9        |
Update Rollup 8 for Exchange Server 2007 SP1               | 5/19/2009       | 8.1.375.2        | 8.01.0375.002   | Exchange Server 2007 SP1             | EX2007_SP1_UR8        |
Update Rollup 7 for Exchange Server 2007 SP1               | 3/18/2009       | 8.1.359.2        | 8.01.0359.002   | Exchange Server 2007 SP1             | EX2007_SP1_UR7        |
Update Rollup 6 for Exchange Server 2007 SP1               | 2/10/2009       | 8.1.340.1        | 8.01.0340.001   | Exchange Server 2007 SP1             | EX2007_SP1_UR6        |
Update Rollup 5 for Exchange Server 2007 SP1               | 11/20/2008      | 8.1.336.1        | 8.01.0336.01    | Exchange Server 2007 SP1             | EX2007_SP1_UR5        |
Update Rollup 4 for Exchange Server 2007 SP1               | 10/7/2008       | 8.1.311.3        | 8.01.0311.003   | Exchange Server 2007 SP1             | EX2007_SP1_UR4        |
Update Rollup 3 for Exchange Server 2007 SP1               | 7/8/2008        | 8.1.291.2        | 8.01.0291.002   | Exchange Server 2007 SP1             | EX2007_SP1_UR3        |
Update Rollup 2 for Exchange Server 2007 SP1               | 5/9/2008        | 8.1.278.2        | 8.01.0278.002   | Exchange Server 2007 SP1             | EX2007_SP1_UR2        |
Update Rollup 1 for Exchange Server 2007 SP1               | 2/28/2008       | 8.1.263.1        | 8.01.0263.001   | Exchange Server 2007 SP1             | EX2007_SP1_UR1        |
Exchange Server 2007 SP1                                   | 11/29/2007      | 8.1.240.6        | 8.01.0240.006   | Exchange Server 2007 SP1             | EX2007_SP1            | TRUE
Update Rollup 7 for Exchange Server 2007                   | 7/8/2008        | 8.0.813.0        | 8.00.0813.000   | Exchange Server 2007 RTM             | EX2007_UR7            |
Update Rollup 6 for Exchange Server 2007                   | 2/21/2008       | 8.0.783.2        | 8.00.0783.002   | Exchange Server 2007 RTM             | EX2007_UR6            |
Update Rollup 5 for Exchange Server 2007                   | 10/25/2007      | 8.0.754.0        | 8.00.0754.000   | Exchange Server 2007 RTM             | EX2007_UR5            |
Update Rollup 4 for Exchange Server 2007                   | 8/23/2007       | 8.0.744.0        | 8.00.0744.000   | Exchange Server 2007 RTM             | EX2007_UR4            |
Update Rollup 3 for Exchange Server 2007                   | 6/28/2007       | 8.0.730.1        | 8.00.0730.001   | Exchange Server 2007 RTM             | EX2007_UR3            |
Update Rollup 2 for Exchange Server 2007                   | 5/8/2007        | 8.0.711.2        | 8.00.0711.002   | Exchange Server 2007 RTM             | EX2007_UR2            |
Update Rollup 1 for Exchange Server 2007                   | 4/17/2007       | 8.0.708.3        | 8.00.0708.003   | Exchange Server 2007 RTM             | EX2007_UR1            |
Exchange Server 2007 RTM                                   | 3/8/2007        | 8.0.685.25       | 8.00.0685.025   | Exchange Server 2007 RTM             | EX2007_RTM            | TRUE
Exchange Server 2003 post-SP2                              | 8/1/2008        | 6.5.7654.4       |                 | Exchange Server 2003 post-SP2        | EX2003_post-SP2       | TRUE
Exchange Server 2003 post-SP2                              | 3/1/2008        | 6.5.7653.33      |                 | Exchange Server 2003 post-SP2        | EX2003_post-SP2       | TRUE
Exchange Server 2003 SP2                                   | 10/19/2005      | 6.5.7683         |                 | Exchange Server 2003 SP2             | EX2003_SP2            | TRUE
Exchange Server 2003 SP1                                   | 5/25/2004       | 6.5.7226         |                 | Exchange Server 2003 SP1             | EX2003_SP1            | TRUE
Exchange Server 2003                                       | 9/28/2003       | 6.5.6944         |                 | Exchange Server 2003                 | EX2003                |
Exchange 2000 Server post-SP3                              | 8/1/2008        | 6.0.6620.7       |                 | Exchange 2000 Server post-SP3        | EX2000_post-SP3       | TRUE
Exchange 2000 Server post-SP3                              | 3/1/2008        | 6.0.6620.5       |                 | Exchange 2000 Server post-SP3        | EX2000_post-SP3       | TRUE
Exchange 2000 Server post-SP3                              | 8/1/2004        | 6.0.6603         |                 | Exchange 2000 Server post-SP3        | EX2000_post-SP3       | TRUE
Exchange 2000 Server post-SP3                              | 4/1/2004        | 6.0.6556         |                 | Exchange 2000 Server post-SP3        | EX2000_post-SP3       | TRUE
Exchange 2000 Server post-SP3                              | 9/1/2003        | 6.0.6487         |                 | Exchange 2000 Server post-SP3        | EX2000_post-SP3       | TRUE
Exchange 2000 Server SP3                                   | 7/18/2002       | 6.0.6249         |                 | Exchange 2000 Server SP3             | EX2000_SP3            | TRUE
Exchange 2000 Server SP2                                   | 11/29/2001      | 6.0.5762         |                 | Exchange 2000 Server SP2             | EX2000_SP2            | TRUE
Exchange 2000 Server SP1                                   | 6/21/2001       | 6.0.4712         |                 | Exchange 2000 Server SP1             | EX2000_SP1            | TRUE
Exchange 2000 Server                                       | 11/29/2000      | 6.0.4417         |                 | Exchange 2000 Server                 | EX2000                |
Exchange Server version 5.5 SP4                            | 11/1/2000       | 5.5.2653         |                 | Exchange Server version 5.5 SP4      | EX55_SP4              | TRUE
Exchange Server version 5.5 SP3                            | 9/9/1999        | 5.5.2650         |                 | Exchange Server version 5.5 SP3      | EX55_SP3              | TRUE
Exchange Server version 5.5 SP2                            | 12/23/1998      | 5.5.2448         |                 | Exchange Server version 5.5 SP2      | EX55_SP2              | TRUE
Exchange Server version 5.5 SP1                            | 8/5/1998        | 5.5.2232         |                 | Exchange Server version 5.5 SP1      | EX55_SP1              | TRUE
Exchange Server version 5.5                                | 2/3/1998        | 5.5.1960         |                 | Exchange Server version 5.5          | EX55                  |
Exchange Server 5.0 SP2                                    | 2/19/1998       | 5.0.1460         |                 | Exchange Server 5.0 SP2              | EX50_SP2              | TRUE
Exchange Server 5.0 SP1                                    | 6/18/1997       | 5.0.1458         |                 | Exchange Server 5.0 SP1              | EX50_SP1              | TRUE
Exchange Server 5.0                                        | 5/23/1997       | 5.0.1457         |                 | Exchange Server 5.0                  | EX50                  |
Exchange Server 4.0 SP5                                    | 5/5/1998        | 4.0.996          |                 | Exchange Server 4.0 SP5              | EX40_SP5              |
Exchange Server 4.0 SP4                                    | 3/28/1997       | 4.0.995          |                 | Exchange Server 4.0 SP4              | EX40_SP4              | TRUE
Exchange Server 4.0 SP3                                    | 10/29/1996      | 4.0.994          |                 | Exchange Server 4.0 SP3              | EX40_SP3              | TRUE
Exchange Server 4.0 SP2                                    | 7/19/1996       | 4.0.993          |                 | Exchange Server 4.0 SP2              | EX40_SP2              | TRUE
Exchange Server 4.0 SP1                                    | 5/1/1996        | 4.0.838          |                 | Exchange Server 4.0 SP1              | EX40_SP1              | TRUE
Exchange Server 4.0 Standard Edition                       | 6/11/1996       | 4.0.837          |                 | Exchange Server 4.0 Standard Edition | EX40_SE               | TRUE
"@ | convertFrom-MarkdownTable ;    
            if($Version){ 
                $xopBuildsndx = @{}
                $Key = 'BuildNumberShort' ;             
                write-verbose "building indexed hash on $($Key)" ; 
                Foreach ($Item in $xopBuilds){
                    $Procd++ ; 
                    if($Key -eq 'BuildNumberLong' -AND ($null -eq $Item.BuildNumberLong)){
                        $Item.BuildNumberLong = $Item.BuildNumberShort
                    } ; 
                    $xopBuildsndx[$Item.$Key.ToString()] = $Item ;             
                } ; 
                $xopBuilds = $null ; 
            } ; 
        } ;  # BEG-E
        PROCESS {
            if($Version){
                foreach($Vers in $Version){
                    if($BuildData = $xopBuildsndx[$Vers.tostring()]){
                        #Return $BuildData ; 
                        Return [pscustomobject]$BuildData ;
                    }else{
                        Return $false ; 
                    } ;
                } ; 
            } ;
            if($AllVersions){
                write-host "-AllVersions: Returning full builds table to pipeline (for post-filtering)" ; 
                Return [pscustomobject]$xopBuilds ; 
            } ; 
        };  # PROC-E
        END {}
    } ; 
    #endregion RESOLVE_XOPBUILDSEMVERSTOTEXTNAMETDO ; #*------^ END Resolve-xopBuildSemVersToTextNameTDO ^------