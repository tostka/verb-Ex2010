# get-xopServerAdminDisplayVersion.ps1
#*----------v Function get-xopServerAdminDisplayVersion() v----------
function get-xopServerAdminDisplayVersion {
    <#
    .SYNOPSIS
    get-xopServerAdminDisplayVersion.ps1 - Retrieves specified ComputerName's (get-exchangeserver).AdminDisplayVersion ~ Cumulative Update version (can optionally retrieve actual Service Update version w -getSErviceUpdate param)
    .NOTES
    Version     : 1.0.2
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2025-
    FileName    : get-xopServerAdminDisplayVersion.ps1
    License     : (non asserted)
    Copyright   : (non asserted)
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell,ExchangeServer,Version
    AddedCredit : theSysadminChannel
    AddedWebsite: https://thesysadminchannel.com/get-exchange-cumulative-update-version-and-build-numbers-using-powershell/
    AddedTwitter: URL
    REVISION
    * 2:58 PM 7/12/2025 updated BuildToProductName indexed hash to specs posted as of 04/25/2025 at https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates
            Maxes reflected in this script, as of that time:
             - Exchange Server SE RTM 	July 1, 2025 	15.2.2562.17 	15.02.2562.017
             - Exchange Server 2019 CU15 May25HU 	May 29, 2025 	15.2.1748.26 	15.02.1748.026
             - Exchange Server 2016 CU23 May25HU 	May 29, 2025 	15.1.2507.57 	15.01.2507.057
             - Exchange Server 2013 CU23 Mar23SU 	March 14, 2023 	15.0.1497.48 	15.00.1497.048
             - Update Rollup 32 for Exchange Server 2010 SP3 	March 2, 2021 	14.3.513.0 	14.03.0513.000
    * 11:06 AM 4/4/2025:
        - udpated CBH (updates more extensive demos);
        - add: -GetServiceUpdate; echo's last BuildTable date, and url to screen; 
            pre-resolve specified Computername to DNS A record fqdn (( via Resolve-DNS -name Computername -Type A).name)
            Object returned to pipeline includes ServiceUpdateVersion & ServiceUpdateProduct , when -getServiceUpdate param is used.
        - removed: fqdn retry code (no longer needed if pre-fqdn'ing)
    * 1:28 PM 3/26/2025 updated CBH, with description/specifics on admindisplayversion returned by get-exchangeserver
        strange: initial atttempts on the lowest # local hub consisntly failed lookup on nbname, so added code to retry with resolved A record/fqdn; 
        ren Get-ExchangeVersion -> get-xopServerAdminDisplayVersion ; 
        updated CBH; 
        updated BuildToProductName indexed hash to specs posted as of 3/26/2025 at https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates
            Maxes reflected in this script, as of that time:
             - Exchange Server 2019 CU15 (2025H1) 	February 10, 2025 	15.2.1748.10 	15.02.1748.010
             - Exchange Server 2016 CU23 Nov24SUv2 	November 27, 2024 	15.1.2507.44 	15.01.2507.044
             - Exchange Server 2013 CU23 Mar23SU 	March 14, 2023 	15.0.1497.48 	15.00.1497.048
             - Update Rollup 32 for Exchange Server 2010 SP3 	March 2, 2021 	14.3.513.0 	14.03.0513.000
    * 2021-Nov-9 tSC's posted vers at https://thesysadminchannel.com
    .DESCRIPTION
    get-xopServerAdminDisplayVersion.ps1 - Retrieves specified ComputerName's (get-exchangeserver).AdminDisplayVersion ~ Cumulative Update version (can optionally retrieve actual Service Update version w -getSErviceUpdate param)

    Expanded variant of tSC's posted script at https://thesysadminchannel.com. 
    - Expands coverage back through 'Exchange Server 2010 RTM', and as of latest BuildToProductName update, through 'Exchange Server 2019 CU15 (2025H1)'
    - Adds SU retrieval (via -getServiceUpdate param, reads remote server ExSetup.exe file version)
    - More fault tolerance (pre-expands specified -ComputerNames into DNS A record fqdn's - seems to avoid sporadic issues retrieving get-exchangeserver & remote invoke-expression; 
        adds code to retry failing queries)
    - BuildToProductName is updated through 3/25/25 current info, and reflects the MS Build table product name strings (unmodified; simpler to maintain over time). 

    Per [Exchange Server build numbers and release dates | Microsoft Learn](https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates?view=exchserver-2019)

    Method details: 
    get-exchangeserver [server] returns Admindisplayversion like : 
    
        Version nn.n (Build nnn.n)
    
    The Buildnumber Shortstring can be converted to a value suitable for BuildToProductName indexed hash lookup, 
    by combining the 'Version nn.n' digits with the '(Build nnn.n)' digits, into: nn.n.nnn.n
    Handled via this regex: 
    
        [regex]::Matches($AdminDisplayVersion,"(\d*\.\d*)").value -join '.'

    If using version-specific code, do any matching on the nn.n for each Major rev, if you want to be sure of your supported commandset targets
        - 2019, 'Version 15.2'
        - 2016, 'Version 15.1'
        - 2013, 'Version 15.0'
        - 2010sp3, 'Version 14.3'
        - 2010sp2, 'Version 14.2'
        - 2010sp1, 'Version 14.1'
        - 2010, 'Version 14.0'

    .PARAMETER ComputerName
    Array of Exchange server names to be queried
    .PARAMETER GetServiceUpdate
    Switch to remote-query the ServiceUpdate revision (polling Version on Exsetup.exe)
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    None. Returns no objects or output (.NET types)
    System.Boolean
    [| get-member the output to see what .NET obj TypeName is returned, to use here]
    .EXAMPLE
    PS> Get-ExchangeServer | get-xopServerAdminDisplayVersion ; 

        ComputerName Edition    BuildNumber   ProductName
        ------------ -------    ------------  -----------
        nnnnnnn1     Standard   15.2.1748.10  Exchange Server 2019 CU15 (2025H1)
        nnnnnnn2     Enterprise 15.2.1748.10  Exchange Server 2019 CU15 (2025H1)

    Pipeline demo
    .EXAMPLE
    PS> get-xopServerAdminDisplayVersion -ComputerName @(ExchSrv01, ExchSrv02) ; 
    Typical pass on an array of servers
    .EXAMPLE
    PS> get-xopServerAdminDisplayVersion -ComputerName @(ExchSrv01, ExchSrv02) -getServiceUpdate ; 

        ComputerName         : xxxxxxxx
        Edition              : Standard
        BuildNumber          : 14.3.123.4
        ProductName          : Exchange Server 2010 SP3
        ServiceUpdateVersion : 14.3.513.0
        ServiceUpdateProduct : Update Rollup 32 for Exchange Server 2010 SP3
        ...

    Typical pass on an array of servers, and return SU version (in addition to CU reported by AdminDisplayVersion)
    .EXAMPLE
    PS> $results = get-exchangeserver | get-xopServerAdminDisplayVersion -getSU ; 
    PS> $results | %{
    PS>     $smsg = "`n$(($_ | ft -a ($_.psobject.Properties.name[0..3])|out-string).trim())" ; 
    PS>     $smsg += "`n$(($_ | ft -a ($_.psobject.Properties.name[-2..-1])|out-string).trim())`n" ;
    PS>     write-host -foregroundcolor green $smsg ; 
    PS> } ;     

        ComputerName Edition    BuildNumber ProductName             
        ------------ -------    ----------- -----------             
        xxxxxxxx     Enterprise 14.3.123.4  Exchange Server 2010 SP3
        ServiceUpdateVersion ServiceUpdateProduct                         
        -------------------- --------------------                         
        14.3.513.0           Update Rollup 32 for Exchange Server 2010 SP3

    Fancier formatted output demo, using -getSU alias for -getServiceUpdate
    .LINK
    https://github.com/tostka/verb-ex2010
    .LINK
    https://thesysadminchannel.com/get-exchange-cumulative-update-version-and-build-numbers-using-powershell/
    .LINK
    https://docs.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates?
    #>
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory = $true,ValueFromPipeline=$true,HelpMessage="Array of Exchange server names to be queried")]
            [string[]] $ComputerName,
        [Parameter(Mandatory = $true,ValueFromPipeline=$true,HelpMessage="Switch to remote-query the ServiceUpdate revision (polling Version on Exsetup.exe")]
            [Alias('getSU')]
            [switch]$GetServiceUpdate
    ) ; 
    BEGIN {
        # when updating $BuildToProductName table (below), also record date of last update here (echos to console, for awareness on results)
        [datetime]$lastBuildTableUpedate = '3/26/2025' ; 
        $BuildTableUpedateUrl = 'https://docs.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-date' ; 
        #Creating the hash table with build numbers and cumulative updates
        # updated as of 9:56 AM 3/26/2025 to curr https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates?view=exchserver-2019
        # also using unmodified MS Build names, from the chart (changing just burns time)
        write-verbose "`$ComputerName:$($ComputerName)" ; 
        $smsg = "NOTE:`$BuildToProductName table was last updated on $($lastBuildTableUpedate.ToShortDateString())" ; 
        $smsg += "`n(update from:$($BuildTableUpedateUrl))" ;
        write-host -foregroundcolor yellow $smsg ; 
        $BuildToProductName = @{
            '15.2.2562.17' = 'Exchange Server SE RTM'
            '15.2.1748.26' = 'Exchange Server 2019 CU15 May25HU'
            '15.2.1748.24' =  'Exchange Server 2019 CU15 Apr25HU'
            '15.2.1748.10' = 	'Exchange Server 2019 CU15 (2025H1)'
            '15.2.1544.14' = 	'Exchange Server 2019 CU14 Nov24SUv2'
            '15.2.1544.13' = 	'Exchange Server 2019 CU14 Nov24SU'
            '15.2.1544.11' = 	'Exchange Server 2019 CU14 Apr24HU'
            '15.2.1544.9' = 	'Exchange Server 2019 CU14 Mar24SU'
            '15.2.1544.4' = 	'Exchange Server 2019 CU14 (2024H1)'
            '15.2.1258.39' = 	'Exchange Server 2019 CU13 Nov24SUv2'
            '15.2.1258.38' = 	'Exchange Server 2019 CU13 Nov24SU'
            '15.2.1258.34' = 	'Exchange Server 2019 CU13 Apr24HU'
            '15.2.1258.32' = 	'Exchange Server 2019 CU13 Mar24SU'
            '15.2.1258.28' = 	'Exchange Server 2019 CU13 Nov23SU'
            '15.2.1258.27' = 	'Exchange Server 2019 CU13 Oct23SU'
            '15.2.1258.25' = 	'Exchange Server 2019 CU13 Aug23SUv2'
            '15.2.1258.23' = 	'Exchange Server 2019 CU13 Aug23SU'
            '15.2.1258.16' = 	'Exchange Server 2019 CU13 Jun23SU'
            '15.2.1258.12' = 	'Exchange Server 2019 CU13 (2023H1)'
            '15.2.1118.40' = 	'Exchange Server 2019 CU12 Nov23SU'
            '15.2.1118.39' = 	'Exchange Server 2019 CU12 Oct23SU'
            '15.2.1118.37' = 	'Exchange Server 2019 CU12 Aug23SUv2'
            '15.2.1118.36' = 	'Exchange Server 2019 CU12 Aug23SU'
            '15.2.1118.30' = 	'Exchange Server 2019 CU12 Jun23SU'
            '15.2.1118.26' = 	'Exchange Server 2019 CU12 Mar23SU'
            '15.2.1118.25' = 	'Exchange Server 2019 CU12 Feb23SU'
            '15.2.1118.21' = 	'Exchange Server 2019 CU12 Jan23SU'
            '15.2.1118.20' = 	'Exchange Server 2019 CU12 Nov22SU'
            '15.2.1118.15' = 	'Exchange Server 2019 CU12 Oct22SU'
            '15.2.1118.12' = 	'Exchange Server 2019 CU12 Aug22SU'
            '15.2.1118.9' = 	'Exchange Server 2019 CU12 May22SU'
            '15.2.1118.7' = 	'Exchange Server 2019 CU12 (2022H1)'
            '15.2.986.42' = 	'Exchange Server 2019 CU11 Mar23SU'
            '15.2.986.41' = 	'Exchange Server 2019 CU11 Feb23SU'
            '15.2.986.37' = 	'Exchange Server 2019 CU11 Jan23SU'
            '15.2.986.36' = 	'Exchange Server 2019 CU11 Nov22SU'
            '15.2.986.30' = 	'Exchange Server 2019 CU11 Oct22SU'
            '15.2.986.29' = 	'Exchange Server 2019 CU11 Aug22SU'
            '15.2.986.26' = 	'Exchange Server 2019 CU11 May22SU'
            '15.2.986.22' = 	'Exchange Server 2019 CU11 Mar22SU'
            '15.2.986.15' = 	'Exchange Server 2019 CU11 Jan22SU'
            '15.2.986.14' = 	'Exchange Server 2019 CU11 Nov21SU'
            '15.2.986.9' = 	'Exchange Server 2019 CU11 Oct21SU'
            '15.2.986.5' = 	'Exchange Server 2019 CU11'
            '15.2.922.27' = 	'Exchange Server 2019 CU10 Mar22SU'
            '15.2.922.20' = 	'Exchange Server 2019 CU10 Jan22SU'
            '15.2.922.19' = 	'Exchange Server 2019 CU10 Nov21SU'
            '15.2.922.14' = 	'Exchange Server 2019 CU10 Oct21SU'
            '15.2.922.13' = 	'Exchange Server 2019 CU10 Jul21SU'
            '15.2.922.7' = 	'Exchange Server 2019 CU10'
            '15.2.858.15' = 	'Exchange Server 2019 CU9 Jul21SU'
            '15.2.858.12' = 	'Exchange Server 2019 CU9 May21SU'
            '15.2.858.10' = 	'Exchange Server 2019 CU9 Apr21SU'
            '15.2.858.5' = 	'Exchange Server 2019 CU9'
            '15.2.792.15' = 	'Exchange Server 2019 CU8 May21SU'
            '15.2.792.13' = 	'Exchange Server 2019 CU8 Apr21SU'
            '15.2.792.10' = 	'Exchange Server 2019 CU8 Mar21SU'
            '15.2.792.3' = 	'Exchange Server 2019 CU8'
            '15.2.721.13' = 	'Exchange Server 2019 CU7 Mar21SU'
            '15.2.721.2' = 	'Exchange Server 2019 CU7'
            '15.2.659.12' = 	'Exchange Server 2019 CU6 Mar21SU'
            '15.2.659.4' = 	'Exchange Server 2019 CU6'
            '15.2.595.8' = 	'Exchange Server 2019 CU5 Mar21SU'
            '15.2.595.3' = 	'Exchange Server 2019 CU5'
            '15.2.529.13' = 	'Exchange Server 2019 CU4 Mar21SU'
            '15.2.529.5' = 	'Exchange Server 2019 CU4'
            '15.2.464.15' = 	'Exchange Server 2019 CU3 Mar21SU'
            '15.2.464.5' = 	'Exchange Server 2019 CU3'
            '15.2.397.11' = 	'Exchange Server 2019 CU2 Mar21SU'
            '15.2.397.3' = 	'Exchange Server 2019 CU2'
            '15.2.330.11' = 	'Exchange Server 2019 CU1 Mar21SU'
            '15.2.330.5' = 	'Exchange Server 2019 CU1'
            '15.2.221.18' = 	'Exchange Server 2019 RTM Mar21SU'
            '15.2.221.12' = 	'Exchange Server 2019 RTM'
            '15.2.196.0' = 	'Exchange Server 2019 Preview'
            '15.1.2507.57' = 	'Exchange Server 2016 CU23 May25HU'
            '15.1.2507.55' = 	'Exchange Server 2016 CU23 Apr25HU'
            '15.1.2507.44' = 	'Exchange Server 2016 CU23 Nov24SUv2'
            '15.1.2507.43' = 	'Exchange Server 2016 CU23 Nov24SU'
            '15.1.2507.39' = 	'Exchange Server 2016 CU23 Apr24HU'
            '15.1.2507.37' = 	'Exchange Server 2016 CU23 Mar24SU'
            '15.1.2507.35' = 	'Exchange Server 2016 CU23 Nov23SU'
            '15.1.2507.34' = 	'Exchange Server 2016 CU23 Oct23SU'
            '15.1.2507.32' = 	'Exchange Server 2016 CU23 Aug23SUv2'
            '15.1.2507.31' = 	'Exchange Server 2016 CU23 Aug23SU'
            '15.1.2507.27' = 	'Exchange Server 2016 CU23 Jun23SU'
            '15.1.2507.23' = 	'Exchange Server 2016 CU23 Mar23SU'
            '15.1.2507.21' = 	'Exchange Server 2016 CU23 Feb23SU'
            '15.1.2507.17' = 	'Exchange Server 2016 CU23 Jan23SU'
            '15.1.2507.16' = 	'Exchange Server 2016 CU23 Nov22SU'
            '15.1.2507.13' = 	'Exchange Server 2016 CU23 Oct22SU'
            '15.1.2507.12' = 	'Exchange Server 2016 CU23 Aug22SU'
            '15.1.2507.9' = 	'Exchange Server 2016 CU23 May22SU'
            '15.1.2507.6' = 	'Exchange Server 2016 CU23 (2022H1)'
            '15.1.2375.37' = 	'Exchange Server 2016 CU22 Nov22SU'
            '15.1.2375.32' = 	'Exchange Server 2016 CU22 Oct22SU'
            '15.1.2375.31' = 	'Exchange Server 2016 CU22 Aug22SU'
            '15.1.2375.28' = 	'Exchange Server 2016 CU22 May22SU'
            '15.1.2375.24' = 	'Exchange Server 2016 CU22 Mar22SU'
            '15.1.2375.18' = 	'Exchange Server 2016 CU22 Jan22SU'
            '15.1.2375.17' = 	'Exchange Server 2016 CU22 Nov21SU'
            '15.1.2375.12' = 	'Exchange Server 2016 CU22 Oct21SU'
            '15.1.2375.7' = 	'Exchange Server 2016 CU22'
            '15.1.2308.27' = 	'Exchange Server 2016 CU21 Mar22SU'
            '15.1.2308.21' = 	'Exchange Server 2016 CU21 Jan22SU'
            '15.1.2308.20' = 	'Exchange Server 2016 CU21 Nov21SU'
            '15.1.2308.15' = 	'Exchange Server 2016 CU21 Oct21SU'
            '15.1.2308.14' = 	'Exchange Server 2016 CU21 Jul21SU'
            '15.1.2308.8' = 	'Exchange Server 2016 CU21'
            '15.1.2242.12' = 	'Exchange Server 2016 CU20 Jul21SU'
            '15.1.2242.10' = 	'Exchange Server 2016 CU20 May21SU'
            '15.1.2242.8' = 	'Exchange Server 2016 CU20 Apr21SU'
            '15.1.2242.4' = 	'Exchange Server 2016 CU20'
            '15.1.2176.14' = 	'Exchange Server 2016 CU19 May21SU'
            '15.1.2176.12' = 	'Exchange Server 2016 CU19 Apr21SU'
            '15.1.2176.9' = 	'Exchange Server 2016 CU19 Mar21SU'
            '15.1.2176.2' = 	'Exchange Server 2016 CU19'
            '15.1.2106.13' = 	'Exchange Server 2016 CU18 Mar21SU'
            '15.1.2106.2' = 	'Exchange Server 2016 CU18'
            '15.1.2044.13' = 	'Exchange Server 2016 CU17 Mar21SU'
            '15.1.2044.4' = 	'Exchange Server 2016 CU17'
            '15.1.1979.8' = 	'Exchange Server 2016 CU16 Mar21SU'
            '15.1.1979.3' = 	'Exchange Server 2016 CU16'
            '15.1.1913.12' = 	'Exchange Server 2016 CU15 Mar21SU'
            '15.1.1913.5' = 	'Exchange Server 2016 CU15'
            '15.1.1847.12' = 	'Exchange Server 2016 CU14 Mar21SU'
            '15.1.1847.3' = 	'Exchange Server 2016 CU14'
            '15.1.1779.8' = 	'Exchange Server 2016 CU13 Mar21SU'
            '15.1.1779.2' = 	'Exchange Server 2016 CU13'
            '15.1.1713.10' = 	'Exchange Server 2016 CU12 Mar21SU'
            '15.1.1713.5' = 	'Exchange Server 2016 CU12'
            '15.1.1591.18' = 	'Exchange Server 2016 CU11 Mar21SU'
            '15.1.1591.10' = 	'Exchange Server 2016 CU11'
            '15.1.1531.12' = 	'Exchange Server 2016 CU10 Mar21SU'
            '15.1.1531.3' = 	'Exchange Server 2016 CU10'
            '15.1.1466.16' = 	'Exchange Server 2016 CU9 Mar21SU'
            '15.1.1466.3' = 	'Exchange Server 2016 CU9'
            '15.1.1415.10' = 	'Exchange Server 2016 CU8 Mar21SU'
            '15.1.1415.2' = 	'Exchange Server 2016 CU8'
            '15.1.1261.35' = 	'Exchange Server 2016 CU7'
            '15.1.1034.26' = 	'Exchange Server 2016 CU6'
            '15.1.845.34' = 	'Exchange Server 2016 CU5'
            '15.1.669.32' = 	'Exchange Server 2016 CU4'
            '15.1.544.27' = 	'Exchange Server 2016 CU3'
            '15.1.466.34' = 	'Exchange Server 2016 CU2'
            '15.1.396.30' = 	'Exchange Server 2016 CU1'
            '15.1.225.42' = 	'Exchange Server 2016 RTM'
            '15.1.225.16' = 	'Exchange Server 2016 Preview'
            '15.0.1497.48' = 	'Exchange Server 2013 CU23 Mar23SU'
            '15.0.1497.47' = 	'Exchange Server 2013 CU23 Feb23SU'
            '15.0.1497.45' = 	'Exchange Server 2013 CU23 Jan23SU'
            '15.0.1497.44' = 	'Exchange Server 2013 CU23 Nov22SU'
            '15.0.1497.42' = 	'Exchange Server 2013 CU23 Oct22SU'
            '15.0.1497.40' = 	'Exchange Server 2013 CU23 Aug22SU'
            '15.0.1497.36' = 	'Exchange Server 2013 CU23 May22SU'
            '15.0.1497.33' = 	'Exchange Server 2013 CU23 Mar22SU'
            '15.0.1497.28' = 	'Exchange Server 2013 CU23 Jan22SU'
            '15.0.1497.26' = 	'Exchange Server 2013 CU23 Nov21SU'
            '15.0.1497.24' = 	'Exchange Server 2013 CU23 Oct21SU'
            '15.0.1497.23' = 	'Exchange Server 2013 CU23 Jul21SU'
            '15.0.1497.18' = 	'Exchange Server 2013 CU23 May21SU'
            '15.0.1497.15' = 	'Exchange Server 2013 CU23 Apr21SU'
            '15.0.1497.12' = 	'Exchange Server 2013 CU23 Mar21SU'
            '15.0.1497.2' = 	'Exchange Server 2013 CU23'
            '15.0.1473.6' = 	'Exchange Server 2013 CU22 Mar21SU'
            '15.0.1473.3' = 	'Exchange Server 2013 CU22'
            '15.0.1395.12' = 	'Exchange Server 2013 CU21 Mar21SU'
            '15.0.1395.4' = 	'Exchange Server 2013 CU21'
            '15.0.1367.3' = 	'Exchange Server 2013 CU20'
            '15.0.1365.1' = 	'Exchange Server 2013 CU19'
            '15.0.1347.2' = 	'Exchange Server 2013 CU18'
            '15.0.1320.4' = 	'Exchange Server 2013 CU17'
            '15.0.1293.2' = 	'Exchange Server 2013 CU16'
            '15.0.1263.5' = 	'Exchange Server 2013 CU15'
            '15.0.1236.3' = 	'Exchange Server 2013 CU14'
            '15.0.1210.3' = 	'Exchange Server 2013 CU13'
            '15.0.1178.4' = 	'Exchange Server 2013 CU12'
            '15.0.1156.6' = 	'Exchange Server 2013 CU11'
            '15.0.1130.7' = 	'Exchange Server 2013 CU10'
            '15.0.1104.5' = 	'Exchange Server 2013 CU9'
            '15.0.1076.9' = 	'Exchange Server 2013 CU8'
            '15.0.1044.25' = 	'Exchange Server 2013 CU7'
            '15.0.995.29' = 	'Exchange Server 2013 CU6'
            '15.0.913.22' = 	'Exchange Server 2013 CU5'
            '15.0.847.64' = 	'Exchange Server 2013 SP1 Mar21SU'
            '15.0.847.32' = 	'Exchange Server 2013 SP1'
            '15.0.775.38' = 	'Exchange Server 2013 CU3'
            '15.0.712.24' = 	'Exchange Server 2013 CU2'
            '15.0.620.29' = 	'Exchange Server 2013 CU1'
            '15.0.516.32' = 	'Exchange Server 2013 RTM'
            '14.3.513.0' = 	'Update Rollup 32 for Exchange Server 2010 SP3'
            '14.3.509.0' = 	'Update Rollup 31 for Exchange Server 2010 SP3'
            '14.3.496.0' = 	'Update Rollup 30 for Exchange Server 2010 SP3'
            '14.3.468.0' = 	'Update Rollup 29 for Exchange Server 2010 SP3'
            '14.3.461.1' = 	'Update Rollup 28 for Exchange Server 2010 SP3'
            '14.3.452.0' = 	'Update Rollup 27 for Exchange Server 2010 SP3'
            '14.3.442.0' = 	'Update Rollup 26 for Exchange Server 2010 SP3'
            '14.3.435.0' = 	'Update Rollup 25 for Exchange Server 2010 SP3'
            '14.3.419.0' = 	'Update Rollup 24 for Exchange Server 2010 SP3'
            '14.3.417.1' = 	'Update Rollup 23 for Exchange Server 2010 SP3'
            '14.3.411.0' = 	'Update Rollup 22 for Exchange Server 2010 SP3'
            '14.3.399.2' = 	'Update Rollup 21 for Exchange Server 2010 SP3'
            '14.3.389.1' = 	'Update Rollup 20 for Exchange Server 2010 SP3'
            '14.3.382.0' = 	'Update Rollup 19 for Exchange Server 2010 SP3'
            '14.3.361.1' = 	'Update Rollup 18 for Exchange Server 2010 SP3'
            '14.3.352.0' = 	'Update Rollup 17 for Exchange Server 2010 SP3'
            '14.3.336.0' = 	'Update Rollup 16 for Exchange Server 2010 SP3'
            '14.3.319.2' = 	'Update Rollup 15 for Exchange Server 2010 SP3'
            '14.3.301.0' = 	'Update Rollup 14 for Exchange Server 2010 SP3'
            '14.3.294.0' = 	'Update Rollup 13 for Exchange Server 2010 SP3'
            '14.3.279.2' = 	'Update Rollup 12 for Exchange Server 2010 SP3'
            '14.3.266.2' = 	'Update Rollup 11 for Exchange Server 2010 SP3'
            '14.3.248.2' = 	'Update Rollup 10 for Exchange Server 2010 SP3'
            '14.3.235.1' = 	'Update Rollup 9 for Exchange Server 2010 SP3'
            '14.3.224.2' = 	'Update Rollup 8 v2 for Exchange Server 2010 SP3'
            '14.3.224.1' = 	'Update Rollup 8 v1 for Exchange Server 2010 SP3 (recalled)'
            '14.3.210.2' = 	'Update Rollup 7 for Exchange Server 2010 SP3'
            '14.3.195.1' = 	'Update Rollup 6 for Exchange Server 2010 SP3'
            '14.3.181.6' = 	'Update Rollup 5 for Exchange Server 2010 SP3'
            '14.3.174.1' = 	'Update Rollup 4 for Exchange Server 2010 SP3'
            '14.3.169.1' = 	'Update Rollup 3 for Exchange Server 2010 SP3'
            '14.3.158.1' = 	'Update Rollup 2 for Exchange Server 2010 SP3'
            '14.3.146.0' = 	'Update Rollup 1 for Exchange Server 2010 SP3'
            '14.3.123.4' = 	'Exchange Server 2010 SP3'
            '14.2.390.3' = 	'Update Rollup 8 for Exchange Server 2010 SP2'
            '14.2.375.0' = 	'Update Rollup 7 for Exchange Server 2010 SP2'
            '14.2.342.3' = 	'Update Rollup 6 Exchange Server 2010 SP2'
            '14.2.328.10' = 	'Update Rollup 5 v2 for Exchange Server 2010 SP2'
            '14.3.328.5' = 	'Update Rollup 5 for Exchange Server 2010 SP2'
            '14.2.318.4' = 	'Update Rollup 4 v2 for Exchange Server 2010 SP2'
            '14.2.318.2' = 	'Update Rollup 4 for Exchange Server 2010 SP2'
            '14.2.309.2' = 	'Update Rollup 3 for Exchange Server 2010 SP2'
            '14.2.298.4' = 	'Update Rollup 2 for Exchange Server 2010 SP2'
            '14.2.283.3' = 	'Update Rollup 1 for Exchange Server 2010 SP2'
            '14.2.247.5' = 	'Exchange Server 2010 SP2'
            '14.1.438.0' = 	'Update Rollup 8 for Exchange Server 2010 SP1'
            '14.1.421.3' = 	'Update Rollup 7 v3 for Exchange Server 2010 SP1'
            '14.1.421.2' = 	'Update Rollup 7 v2 for Exchange Server 2010 SP1'
            '14.1.421.0' = 	'Update Rollup 7 for Exchange Server 2010 SP1'
            '14.1.355.2' = 	'Update Rollup 6 for Exchange Server 2010 SP1'
            '14.1.339.1' = 	'Update Rollup 5 for Exchange Server 2010 SP1'
            '14.1.323.6' = 	'Update Rollup 4 for Exchange Server 2010 SP1'
            '14.1.289.7' = 	'Update Rollup 3 for Exchange Server 2010 SP1'
            '14.1.270.1' = 	'Update Rollup 2 for Exchange Server 2010 SP1'
            '14.1.255.2' = 	'Update Rollup 1 for Exchange Server 2010 SP1'
            '14.1.218.15' = 	'Exchange Server 2010 SP1'
            '14.0.726.0' = 	'Update Rollup 5 for Exchange Server 2010'
            '14.0.702.1' = 	'Update Rollup 4 for Exchange Server 2010'
            '14.0.694.0' = 	'Update Rollup 3 for Exchange Server 2010'
            '14.0.689.0' = 	'Update Rollup 2 for Exchange Server 2010'
            '14.0.682.1' = 	'Update Rollup 1 for Exchange Server 2010'
            '14.0.639.21' = 	'Exchange Server 2010 RTM'            
        }; 
        $Retries = 4 ;
        $RetrySleep = 2 ;
    } ;  # BEG-E
    PROCESS {
        foreach ($Computer in $ComputerName) {
            $Computer = $Computer.ToUpper()
            write-verbose "==Computer:$($Computer)" ; 
            $Exit = 0 ;
            Do {
                TRY {
                    # getting errors on the invoke-expression: always preconvert DNS nbname to FQDN: 
                    if($cFQDN = (resolve-dnsname -type A $computer | sort { $_.IPAddress -replace '\d+', { $_.Value.PadLeft(3, '0') } } )[-1].name){
                        write-verbose "`$cFQDN:$($cFQDN)" ; 
                        if($Server = get-exchangeserver -Identity $cFQDN -ErrorAction Stop -Verbose:($PSBoundParameters['Verbose'] -eq $true)){
                            $Exit = $Retries ;
                        }; 
                    }elseif($Server = get-exchangeserver -Identity $Computer -ErrorAction Stop -Verbose:($PSBoundParameters['Verbose'] -eq $true)){
                        $Exit = $Retries ;
                    } else { 
                        write-warning "Unable to either:resolve-dnsname -type A $($computer) to FQDN, and/or get-exchangeserver -Identity $($Computer)!`nSKIPPING!" ; 
                        CONTINUE
                    }  ; 
                    write-verbose "`$Server:`n$(($Server|out-string).trim())" ; 
                } CATCH {
                    $ErrorTrapped=$Error[0] ;
                    Write-warning "Failed to exec cmd because: $($ErrorTrapped.Exception.Message )" ;
                    Start-Sleep -Seconds $RetrySleep ;
                    $Exit ++ ;
                    Write-Verbose "Try #: $Exit" ;
                    If ($Exit -eq $Retries) {Write-Warning "Unable to exec cmd!"; BREAK ; } ;
                }  ;
            } Until ($Exit -eq $Retries) ; 

            TRY {
                $Version = $Server.AdminDisplayVersion
                $Version = [regex]::Matches($Version, "(\d*\.\d*)").value -join '.'
                $Product = $BuildToProductName[$Version]
                if($GetServiceUpdate){
                    if($cFQDN){
                        $targetName = $cFQDN ; 
                    }else{
                        $targetName = $Server.Name ; 
                    }
                    if($FileversionInfo = Invoke-Command -ComputerName $targetName -ScriptBlock { Get-Command Exsetup.exe | ForEach-Object { $_.FileversionInfo } } ){
                        write-verbose "`$FileversionInfo:`n$(($FileversionInfo | ft -a |out-string).trim())" ; 
                        [version]$ExsetupRev = (@($FileversionInfo.FileMajorPart,$FileversionInfo.FileMinorPart,$FileversionInfo.FileBuildPart,$FileversionInfo.FilePrivatePart) -join '.')
                        $ExsetupProduct = $BuildToProductName[$ExsetupRev.tostring()]
                        write-verbose "`$ExsetupProduct:$($ExsetupProduct)" ; 
                    } else { 
                        throw "$($Server.name):Unable to remote retrieve: Get-Command Exsetup.exe | ForEach-Object { $_.FileversionInfo}"
                    } ; 
                    $Object = [pscustomobject]@{
                        ComputerName = $Computer
                        Edition      = $Server.Edition
                        BuildNumber  = $Version
                        ProductName  = $Product
                        ServiceUpdateVersion = $ExsetupRev.tostring() ; 
                        ServiceUpdateProduct = $ExsetupProduct ; 
                    }
                }else{
                    $Object = [pscustomobject]@{
                        ComputerName = $Computer ; 
                        Edition      = $Server.Edition ; 
                        BuildNumber  = $Version ; 
                        ProductName  = $Product ; 
                    } ; 
                } ; 
                Write-Output $Object
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
            } FINALLY {
                $Server  = $null ; 
                $Version = $null ; 
                $Product = $null ; 
                $ExsetupRev = $null  ; 
                $ExsetupProduct = $null  ; 
                $FileversionInfo = $null  ; 
                $FileversionInfo = $null  ; 
                $ExsetupProduct = $null  ; 
            }
        } ;  # loop-E
    };  # PROC-E
    END {}
} ; 
#*------^ END Function get-xopServerAdminDisplayVersion() ^------