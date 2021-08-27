#*----------v Function get-ExchangeOPConnectionVersion() v----------
function get-ExchangeOPConnectionVersion {
    <#
    .SYNOPSIS
    get-ExchangeOPConnectionVersion.ps1 - Simple current connection Exchange revision check. Pulls the PSSession matching Name 'Exchange*', and runs a (get-exchangeserver).AdminDisplayVersion to pull back the broad revision of the server on the other side of the connection. 
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-08-25
    FileName    : 
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 3:36 PM 8/25/2021 init
    .DESCRIPTION
    get-ExchangeOPConnectionVersion.ps1 - Code to approximate EmailAddressTemplate-generated email addresses
    Note: This is a quick & dirty *approximation* of the server revision, accorind to a get-ExchangeServer against the computername on the other end of the PSSession named 'Exchange*'
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    System.String
    .EXAMPLE
    PS> $exv = get-ExchangeOPConnectionVersion ;
    PS> if($exv -like '2010*'){'2010'} elseif($exv -like '2007*'){'2007'}else{$exv} ;
    Example demonstrating if/then'ing into a specific revision of code for Exchange onPrem.
    .LINK
    https://github.com/tostka/verb-ex2010
    .LINK
    https://docs.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates
    #>
    ###Requires -Version 5
    ###Requires -Modules verb-Ex2010 - disabled, moving into the module
    #Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.toro\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding(DefaultParameterSetName='EAP')]
    PARAM(
    [Parameter(HelpMessage="Switch to suppress checks for multiple AdminDisplayVersions [-ignoreMulti]")]
    [switch] $ignoreMulti,
    [Parameter(HelpMessage="Switch to report solely the root revision (ignores servicepacks, true by default)[-WholeRevisionsOnly]")]
    [switch] $WholeRevisionsOnly=$true
    ) ;
    BEGIN { 
        # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;            
        if($ignoreMulti){write-verbose "-IgnoreMulti specified: Multiple AdminDisplayVersion checks are being suppressed."} ; 
        
        $pssprops = 'Id','Name','ComputerName','ComputerType','State','ConfigurationName','Availability' ;
        #reconnect-ex2010 -verbose:$false ;
        
        # below sourced in git md (with links edited out):
        # https://docs.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates
        $hVersTable =@"
|ProductName|Releasedate|BuildNumbershort|BuildNumberlong|
|---|---|:---:|:---:|
|Exchange Server 2019 CU10 Jul21SU|July 13, 2021|15.2.922.13|15.02.0922.013|
|Exchange Server 2019 CU10|June 29, 2021|15.2.922.7|15.02.0922.007|
|Exchange Server 2019 CU9 Jul21SU|July 13, 2021|15.2.858.15|15.02.0858.015|
|Exchange Server 2019 CU9 May21SU|May 11, 2021|15.2.858.12|15.02.0858.012|
|Exchange Server 2019 CU9 Apr21SU|April 13, 2021|15.2.858.10|15.02.0858.010|
|Exchange Server 2019 CU9|March 16, 2021|15.2.858.5|15.02.0858.005|
|Exchange Server 2019 CU8 May21SU|May 11, 2021|15.2.792.15|15.02.0792.015|
|Exchange Server 2019 CU8 Apr21SU|April 13, 2021|15.2.792.13|15.02.0792.013|
|Exchange Server 2019 CU8 Mar21SU|March 2, 2021|15.2.792.10|15.02.0792.010|
|Exchange Server 2019 CU8|December 15, 2020|15.2.792.3|15.02.0792.003|
|Exchange Server 2019 CU7 Mar21SU|March 2, 2021|15.2.721.13|15.02.0721.013|
|Exchange Server 2019 CU7|September 15, 2020|15.2.721.2|15.02.0721.002|
|Exchange Server 2019 CU6 Mar21SU|March 2, 2021|15.2.659.12|15.02.0659.012|
|Exchange Server 2019 CU6|June 16, 2020|15.2.659.4|15.02.0659.004|
|Exchange Server 2019 CU5 Mar21SU|March 2, 2021|15.2.595.8|15.02.0595.008|
|Exchange Server 2019 CU5|March 17, 2020|15.2.595.3|15.02.0595.003|
|Exchange Server 2019 CU4 Mar21SU|March 2, 2021|15.2.529.13|15.02.0529.013|
|Exchange Server 2019 CU4|December 17, 2019|15.2.529.5|15.02.0529.005|
|Exchange Server 2019 CU3 Mar21SU|March 2, 2021|15.2.464.15|15.02.0464.015|
|Exchange Server 2019 CU3|September 17, 2019|15.2.464.5|15.02.0464.005|
|Exchange Server 2019 CU2 Mar21SU|March 2, 2021|15.2.397.11|15.02.0397.011|
|Exchange Server 2019 CU2|June 18, 2019|15.2.397.3|15.02.0397.003|
|Exchange Server 2019 CU1 Mar21SU|March 2, 2021|15.2.330.11|15.02.0330.011|
|Exchange Server 2019 CU1|February 12, 2019|15.2.330.5|15.02.0330.005|
|Exchange Server 2019 RTM Mar21SU|March 2, 2021|15.2.221.18|15.02.0221.018|
|Exchange Server 2019 RTM|October 22, 2018|15.2.221.12|15.02.0221.012|
|Exchange Server 2019 Preview|July 24, 2018|15.2.196.0|15.02.0196.000|
|Exchange Server 2016 CU21 Jul21SU|July 13, 2021|15.1.2308.14|15.01.2308.014|
|Exchange Server 2016 CU21|June 29, 2021|15.1.2308.8|15.01.2308.008|
|Exchange Server 2016 CU20 Jul21SU|July 13, 2021|15.1.2242.12|15.01.2242.012|
|Exchange Server 2016 CU20 May21SU|May 11, 2021|15.1.2242.10|15.01.2242.010|
|Exchange Server 2016 CU20 Apr21SU|April 13, 2021|15.1.2242.8|15.01.2242.008|
|Exchange Server 2016 CU20|March 16, 2021|15.1.2242.4|15.01.2242.004|
|Exchange Server 2016 CU19 May21SU|May 11, 2021|15.1.2176.14|15.01.2176.014|
|Exchange Server 2016 CU19 Apr21SU|April 13, 2021|15.1.2176.12|15.01.2176.012|
|Exchange Server 2016 CU19 Mar21SU|March 2, 2021|15.1.2176.9|15.01.2176.009|
|Exchange Server 2016 CU19|December 15, 2020|15.1.2176.2|15.01.2176.002|
|Exchange Server 2016 CU18 Mar21SU|March 2, 2021|15.1.2106.13|15.01.2106.013|
|Exchange Server 2016 CU18|September 15, 2020|15.1.2106.2|15.01.2106.002|
|Exchange Server 2016 CU17 Mar21SU|March 2, 2021|15.1.2044.13|15.01.2044.013|
|Exchange Server 2016 CU17|June 16, 2020|15.1.2044.4|15.01.2044.004|
|Exchange Server 2016 CU16 Mar21SU|March 2, 2021|15.1.1979.8|15.01.1979.008|
|Exchange Server 2016 CU16|March 17, 2020|15.1.1979.3|15.01.1979.003|
|Exchange Server 2016 CU15 Mar21SU|March 2, 2021|15.1.1913.12|15.01.1913.012|
|Exchange Server 2016 CU15|December 17, 2019|15.1.1913.5|15.01.1913.005|
|Exchange Server 2016 CU14 Mar21SU|March 2, 2021|15.1.1847.12|15.01.1847.012|
|Exchange Server 2016 CU14|September 17, 2019|15.1.1847.3|15.01.1847.003|
|Exchange Server 2016 CU13 Mar21SU|March 2, 2021|15.1.1779.8|15.01.1779.008|
|Exchange Server 2016 CU13|June 18, 2019|15.1.1779.2|15.01.1779.002|
|Exchange Server 2016 CU12 Mar21SU|March 2, 2021|15.1.1713.10|15.01.1713.010|
|Exchange Server 2016 CU12|February 12, 2019|15.1.1713.5|15.01.1713.005|
|Exchange Server 2016 CU11 Mar21SU|March 2, 2021|15.1.1591.18|15.01.1591.018|
|Exchange Server 2016 CU11|October 16, 2018|15.1.1591.10|15.01.1591.010|
|Exchange Server 2016 CU10 Mar21SU|March 2, 2021|15.1.1531.12|15.01.1531.012|
|Exchange Server 2016 CU10|June 19, 2018|15.1.1531.3|15.01.1531.003|
|Exchange Server 2016 CU9 Mar21SU|March 2, 2021|15.1.1466.16|15.01.1466.016|
|Exchange Server 2016 CU9|March 20, 2018|15.1.1466.3|15.01.1466.003|
|Exchange Server 2016 CU8 Mar21SU|March 2, 2021|15.1.1415.10|15.01.1415.010|
|Exchange Server 2016 CU8|December 19, 2017|15.1.1415.2|15.01.1415.002|
|Exchange Server 2016 CU7|September 19, 2017|15.1.1261.35|15.01.1261.035|
|Exchange Server 2016 CU6|June 27, 2017|15.1.1034.26|15.01.1034.026|
|Exchange Server 2016 CU5|March 21, 2017|15.1.845.34|15.01.0845.034|
|Exchange Server 2016 CU4|December 13, 2016|15.1.669.32|15.01.0669.032|
|Exchange Server 2016 CU3|September 20, 2016|15.1.544.27|15.01.0544.027|
|Exchange Server 2016 CU2|June 21, 2016|15.1.466.34|15.01.0466.034|
|Exchange Server 2016 CU1|March 15, 2016|15.1.396.30|15.01.0396.030|
|Exchange Server 2016 RTM|October 1, 2015|15.1.225.42|15.01.0225.042|
|Exchange Server 2016 Preview|July 22, 2015|15.1.225.16|15.01.0225.016|
|Exchange Server 2013 CU23 Jul21SU|July 13, 2021|15.0.1497.23|15.00.1497.023|
|Exchange Server 2013 CU23 May21SU|May 11, 2021|15.0.1497.18|15.00.1497.018|
|Exchange Server 2013 CU23 Apr21SU|April 13, 2021|15.0.1497.15|15.00.1497.015|
|Exchange Server 2013 CU23 Mar21SU|March 2, 2021|15.0.1497.12|15.00.1497.012|
|Exchange Server 2013 CU23)|June 18, 2019|15.0.1497.2|15.00.1497.002|
|Exchange Server 2013 CU22 Mar21SU|March 2, 2021|15.0.1473.6|15.00.1473.006|
|Exchange Server 2013 CU22|February 12, 2019|15.0.1473.3|15.00.1473.003|
|Exchange Server 2013 CU21 Mar21SU|March 2, 2021|15.0.1395.12|15.00.1395.012|
|Exchange Server 2013 CU21|June 19, 2018|15.0.1395.4|15.00.1395.004|
|Exchange Server 2013 CU20|March 20, 2018|15.0.1367.3|15.00.1367.003|
|Exchange Server 2013 CU19|December 19, 2017|15.0.1365.1|15.00.1365.001|
|Exchange Server 2013 CU18|September 19, 2017|15.0.1347.2|15.00.1347.002|
|Exchange Server 2013 CU17|June 27, 2017|15.0.1320.4|15.00.1320.004|
|Exchange Server 2013 CU16|March 21, 2017|15.0.1293.2|15.00.1293.002|
|Exchange Server 2013 CU15|December 13, 2016|15.0.1263.5|15.00.1263.005|
|Exchange Server 2013 CU14|September 20, 2016|15.0.1236.3|15.00.1236.003|
|Exchange Server 2013 CU13|June 21, 2016|15.0.1210.3|15.00.1210.003|
|Exchange Server 2013 CU12|March 15, 2016|15.0.1178.4|15.00.1178.004|
|Exchange Server 2013 CU11|December 15, 2015|15.0.1156.6|15.00.1156.006|
|Exchange Server 2013 CU10|September 15, 2015|15.0.1130.7|15.00.1130.007|
|Exchange Server 2013 CU9|June 17, 2015|15.0.1104.5|15.00.1104.005|
|Exchange Server 2013 CU8|March 17, 2015|15.0.1076.9|15.00.1076.009|
|Exchange Server 2013 CU7|December 9, 2014|15.0.1044.25|15.00.1044.025|
|Exchange Server 2013 CU6|August 26, 2014|15.0.995.29|15.00.0995.029|
|Exchange Server 2013 CU5|May 27, 2014|15.0.913.22|15.00.0913.022|
|Exchange Server 2013 SP1 Mar21SU|March 2, 2021|15.0.847.64|15.00.0847.064|
|Exchange Server 2013 SP1|February 25, 2014|15.0.847.32|15.00.0847.032|
|Exchange Server 2013 CU3|November 25, 2013|15.0.775.38|15.00.0775.038|
|Exchange Server 2013 CU2|July 9, 2013|15.0.712.24|15.00.0712.024|
|Exchange Server 2013 CU1|April 2, 2013|15.0.620.29|15.00.0620.029|
|Exchange Server 2013 RTM|December 3, 2012|15.0.516.32|15.00.0516.032|
|Update Rollup 32 for Exchange Server 2010 SP3|March 2, 2021|14.3.513.0|14.03.0513.000|
|Update Rollup 31 for Exchange Server 2010 SP3|December 1, 2020|14.3.509.0|14.03.0509.000|
|Update Rollup 30 for Exchange Server 2010 SP3|February 11, 2020|14.3.496.0|14.03.0496.000|
|Update Rollup 29 for Exchange Server 2010 SP3|July 9, 2019|14.3.468.0|14.03.0468.000|
|Update Rollup 28 for Exchange Server 2010 SP3|June 7, 2019|14.3.461.1|14.03.0461.001|
|Update Rollup 27 for Exchange Server 2010 SP3|April 9, 2019|14.3.452.0|14.03.0452.000|
|Update Rollup 26 for Exchange Server 2010 SP3|February 12, 2019|14.3.442.0|14.03.0442.000|
|Update Rollup 25 for Exchange Server 2010 SP3|January 8, 2019|14.3.435.0|14.03.0435.000|
|Update Rollup 24 for Exchange Server 2010 SP3|September 5, 2018|14.3.419.0|14.03.0419.000|
|Update Rollup 23 for Exchange Server 2010 SP3|August 13, 2018|14.3.417.1|14.03.0417.001|
|Update Rollup 22 for Exchange Server 2010 SP3|June 19, 2018|14.3.411.0|14.03.0411.000|
|Update Rollup 21 for Exchange Server 2010 SP3|May 7, 2018|14.3.399.2|14.03.0399.002|
|Update Rollup 20 for Exchange Server 2010 SP3|March 5, 2018|14.3.389.1|14.03.0389.001|
|Update Rollup 19 for Exchange Server 2010 SP3|December 19, 2017|14.3.382.0|14.03.0382.000|
|Update Rollup 18 for Exchange Server 2010 SP3|July 11, 2017|14.3.361.1|14.03.0361.001|
|Update Rollup 17 for Exchange Server 2010 SP3|March 21, 2017|14.3.352.0|14.03.0352.000|
|Update Rollup 16 for Exchange Server 2010 SP3|December 13, 2016|14.3.336.0|14.03.0336.000|
|Update Rollup 15 for Exchange Server 2010 SP3|September 20, 2016|14.3.319.2|14.03.0319.002|
|Update Rollup 14 for Exchange Server 2010 SP3|June 21, 2016|14.3.301.0|14.03.0301.000|
|Update Rollup 13 for Exchange Server 2010 SP3|March 15, 2016|14.3.294.0|14.03.0294.000|
|Update Rollup 12 for Exchange Server 2010 SP3|December 15, 2015|14.3.279.2|14.03.0279.002|
|Update Rollup 11 for Exchange Server 2010 SP3|September 15, 2015|14.3.266.2|14.03.0266.002|
|Update Rollup 10 for Exchange Server 2010 SP3|June 17, 2015|14.3.248.2|14.03.0248.002|
|Update Rollup 9 for Exchange Server 2010 SP3|March 17, 2015|14.3.235.1|14.03.0235.001|
|Update Rollup 8 v2 for Exchange Server 2010 SP3|December 12, 2014|14.3.224.2|14.03.0224.002|
|Update Rollup 8 v1 for Exchange Server 2010 SP3|December 9, 2014|14.3.224.1|14.03.0224.001|
|Update Rollup 7 for Exchange Server 2010 SP3|August 26, 2014|14.3.210.2|14.03.0210.002|
|Update Rollup 6 for Exchange Server 2010 SP3|May 27, 2014|14.3.195.1|14.03.0195.001|
|Update Rollup 5 for Exchange Server 2010 SP3|February 24, 2014|14.3.181.6|14.03.0181.006|
|Update Rollup 4 for Exchange Server 2010 SP3|December 9, 2013|14.3.174.1|14.03.0174.001|
|Update Rollup 3 for Exchange Server 2010 SP3|November 25, 2013|14.3.169.1|14.03.0169.001|
|Update Rollup 2 for Exchange Server 2010 SP3|August 8, 2013|14.3.158.1|14.03.0158.001|
|Update Rollup 1 for Exchange Server 2010 SP3|May 29, 2013|14.3.146.0|14.03.0146.000|
|Exchange Server 2010 SP3|February 12, 2013|14.3.123.4|14.03.0123.004|
|Update Rollup 8 for Exchange Server 2010 SP2|December 9, 2013|14.2.390.3|14.02.0390.003|
|Update Rollup 7 for Exchange Server 2010 SP2|August 3, 2013|14.2.375.0|14.02.0375.000|
|Update Rollup 6 Exchange Server 2010 SP2|February 12, 2013|14.2.342.3|14.02.0342.003|
|Update Rollup 5 v2 for Exchange Server 2010 SP2|December 10, 2012|14.2.328.10|14.02.0328.010|
|Update Rollup 5 for Exchange Server 2010 SP2|November 13, 2012|14.3.328.5|14.03.0328.005|
|Update Rollup 4 v2 for Exchange Server 2010 SP2|October 9, 2012|14.2.318.4|14.02.0318.004|
|Update Rollup 4 for Exchange Server 2010 SP2|August 13, 2012|14.2.318.2|14.02.0318.002|
|Update Rollup 3 for Exchange Server 2010 SP2|May 29, 2012|14.2.309.2|14.02.0309.002|
|Update Rollup 2 for Exchange Server 2010 SP2|April 16, 2012|14.2.298.4|14.02.0298.004|
|Update Rollup 1 for Exchange Server 2010 SP2|February 13, 2012|14.2.283.3|14.02.0283.003|
|Exchange Server 2010 SP2|December 4, 2011|14.2.247.5|14.02.0247.005|
|Update Rollup 8 for Exchange Server 2010 SP1|December 10, 2012|14.1.438.0|14.01.0438.000|
|Update Rollup 7 v3 for Exchange Server 2010 SP1|November 13, 2012|14.1.421.3|14.01.0421.003|
|Update Rollup 7 v2 for Exchange Server 2010 SP1|October 10, 2012|14.1.421.2|14.01.0421.002|
|Update Rollup 7 for Exchange Server 2010 SP1|August 8, 2012|14.1.421.0|14.01.0421.000|
|Update Rollup 6 for Exchange Server 2010 SP1|October 27, 2011|14.1.355.2|14.01.0355.002|
|Update Rollup 5 for Exchange Server 2010 SP1|August 23, 2011|14.1.339.1|14.01.0339.001|
|Update Rollup 4 for Exchange Server 2010 SP1|July 27, 2011|14.1.323.6|14.01.0323.006|
|Update Rollup 3 for Exchange Server 2010 SP1|April 6, 2011|14.1.289.7|14.01.0289.007|
|Update Rollup 2 for Exchange Server 2010 SP1|December 9, 2010|14.1.270.1|14.01.0270.001|
|Update Rollup 1 for Exchange Server 2010 SP1|October 4, 2010|14.1.255.2|14.01.0255.002|
|Exchange Server 2010 SP1|August 23, 2010|14.1.218.15|14.01.0218.015|
|Update Rollup 5 for Exchange Server 2010|December 13, 2010|14.0.726.0|14.00.0726.000|
|Update Rollup 4 for Exchange Server 2010|June 10, 2010|14.0.702.1|14.00.0702.001|
|Update Rollup 3 for Exchange Server 2010|April 13, 2010|14.0.694.0|14.00.0694.000|
|Update Rollup 2 for Exchange Server 2010|March 4, 2010|14.0.689.0|14.00.0689.000|
|Update Rollup 1 for Exchange Server 2010|December 9, 2009|14.0.682.1|14.00.0682.001|
|Exchange Server 2010 RTM|November 9, 2009|14.0.639.21|14.00.0639.021|
|Update Rollup 23 for Exchange Server 2007 SP3|March 21, 2017|8.3.517.0|8.03.0517.000|
|Update Rollup 22 for Exchange Server 2007 SP3|December 13, 2016|8.3.502.0|8.03.0502.000|
|Update Rollup 21 for Exchange Server 2007 SP3|September 20, 2016|8.3.485.1|8.03.0485.001|
|Update Rollup 20 for Exchange Server 2007 SP3|June 21, 2016|8.3.468.0|8.03.0468.000|
|Update Rollup 19 forExchange Server 2007 SP3|March 15, 2016|8.3.459.0|8.03.0459.000|
|Update Rollup 18 forExchange Server 2007 SP3|December, 2015|8.3.445.0|8.03.0445.000|
|Update Rollup 17 forExchange Server 2007 SP3|June 17, 2015|8.3.417.1|8.03.0417.001|
|Update Rollup 16 for Exchange Server 2007 SP3|March 17, 2015|8.3.406.0|8.03.0406.000|
|Update Rollup 15 for Exchange Server 2007 SP3|December 9, 2014|8.3.389.2|8.03.0389.002|
|Update Rollup 14 for Exchange Server 2007 SP3|August 26, 2014|8.3.379.2|8.03.0379.002|
|Update Rollup 13 for Exchange Server 2007 SP3|February 24, 2014|8.3.348.2|8.03.0348.002|
|Update Rollup 12 for Exchange Server 2007 SP3|December 9, 2013|8.3.342.4|8.03.0342.004|
|Update Rollup 11 for Exchange Server 2007 SP3|August 13, 2013|8.3.327.1|8.03.0327.001|
|Update Rollup 10 for Exchange Server 2007 SP3|February 11, 2013|8.3.298.3|8.03.0298.003|
|Update Rollup 9 for Exchange Server 2007 SP3|December 10, 2012|8.3.297.2|8.03.0297.002|
|Update Rollup 8-v3 for Exchange Server 2007 SP3|November 13, 2012|8.3.279.6|8.03.0279.006|
|Update Rollup 8-v2 for Exchange Server 2007 SP3|October 9, 2012|8.3.279.5|8.03.0279.005|
|Update Rollup 8 for Exchange Server 2007 SP3|August 13, 2012|8.3.279.3|8.03.0279.003|
|Update Rollup 7 for Exchange Server 2007 SP3|April 16, 2012|8.3.264.0|8.03.0264.000|
|Update Rollup 6 for Exchange Server 2007 SP3|January 26, 2012|8.3.245.2|8.03.0245.002|
|Update Rollup 5 for Exchange Server 2007 SP3|September 21, 2011|8.3.213.1|8.03.0213.001|
|Update Rollup 4 for Exchange Server 2007 SP3|May 28, 2011|8.3.192.1|8.03.0192.001|
|Update Rollup 3-v2 for Exchange Server 2007 SP3|March 30, 2011|8.3.159.2|8.03.0159.002|
|Update Rollup 2 for Exchange Server 2007 SP3|December 10, 2010|8.3.137.3|8.03.0137.003|
|Update Rollup 1 for Exchange Server 2007 SP3|September 9, 2010|8.3.106.2|8.03.0106.002|
|Exchange Server 2007 SP3|June 7, 2010|8.3.83.6|8.03.0083.006|
"@ ; 
         $VersTable = convertFrom-MarkdownTable -markdowntext $hVersTable ;
         
         if($Pss = Get-PSSession |?{$_.name -like 'Exchange*'}){
            $PSSConnServer = $PSS.computername ;
            $smsg = "Existing Ex Connection:`n$((Get-PSSession |?{$_.name -like 'Exchange*'}| ft -auto $pssprops |out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } else {
            $smsg = "UNABLE TO LOCATE AN EXISTING EXCHANGE REMOTE POWERSHELL CONNECTION!`nPlease initiate a connection before running this comand" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        
    } ;  # BEGIN-E
    PROCESS {
        $VersString = $null ; 
        
        if($Pss){
            $AdmVers = get-exchangeserver | group AdminDisplayVersion
            if(($AdmVers| measure).count -gt 1 -AND -not($ignoreMulti)){
                $smsg = "Multiple Exchange AdminDisplayVersion values returned!:"
                $smsg += "n$(($AdmVers| select -expand Name| out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } elSe { 
                # pull AdminDisplayVersion on current connection
                
                $ExConnVers = (get-exchangeserver $PSSConnServer)
                <# simple integer revision chk: - seldom accurate, reporting boxes patched with 
                ex10sp3-ru32 as Version 14.3 (Build 123.4) == Exchange Server 2010 SP3|February 12, 2013|14.3.123.4|14.03.0123.004| waaaay off, esp if you need a CU for a supported PS cmdlet
                #>
                $ExConnNN = $ExConnVers.AdminDisplayVersion.split(' ')[1] ; 
                switch ($ExConnVers){
                  '15.2' { $VersString = '2019' }
                  '15.1' { $VersString = '2016' }
                  '15.0' { $VersString = '2013' }
                  '14.3' { 
                      if($WholeRevisionsOnly){'2010'}
                      else { $VersString = '2010sp3'}
                   }
                  '14.2' { 
                      if($WholeRevisionsOnly){'2010'}
                      else { $VersString = '2010sp2' }
                  }
                  '14.1' { 
                      if($WholeRevisionsOnly){'2010'}
                      else { $VersString = '2010sp1' }
                  }
                  '14.0' { 
                      if($WholeRevisionsOnly){'2010'}
                      else { $VersString = '2010rtm' }
                  } 
                  '8.3' { 
                      if($WholeRevisionsOnly){'2007'}
                      else { $VersString = '2007SP3' }
                  }
                  '8.2' { 
                      if($WholeRevisionsOnly){'2007'}
                      else { $VersString = '2007sp2' }
                  } ;
                  '8.1' { 
                      if($WholeRevisionsOnly){'2010'}
                      else { $VersString = '2007sp1' }
                  } 
                  '8.0' { 
                      if($WholeRevisionsOnly){'2010'}
                      else { $VersString = '2007rtm' }
                  } 
                  default { 
                    $smsg = "Unrecognized AdminDisplayVersion retrieved!:`n$($ExConnVers)"
                    throw $smsg ; 
                  }
                } ; 
                
                # going deeper:
                <# remotely Find Exchange version with PowerShell including Security Update
                # below measure-command{}'s as 17secs (2nd attempt, first was probably gt 1min)
$ExchangeServers = Get-ExchangeServer | Sort-Object Name ; 
ForEach ($Server in $ExchangeServers) {
    Invoke-Command -ComputerName $Server.Name -ScriptBlock { Get-Command Exsetup.exe | ForEach-Object { $_.FileversionInfo } }
}
                #>
                # single box measure-commands as  1.01sec
                $ExSetupVers = Invoke-Command -ComputerName $PSSConnServer -ScriptBlock { Get-Command Exsetup.exe | ForEach-Object { $_.FileversionInfo } } ;
                <# obj returned:
                PSComputerName     : FQDN.DOMAIN.com
                RunspaceId         : GUID
                PSShowComputerName : True
                Comments           : Service Pack 3
                CompanyName        : Microsoft Corporation
                FileBuildPart      : 513
                FileDescription    :
                FileMajorPart      : 14
                FileMinorPart      : 3
                FileName           : D:\Program Files\Microsoft\Exchange Server\V14\bin\ExSetup.exe
                FilePrivatePart    : 0
                FileVersion        : 14.03.0513.000
                InternalName       : ExSetup.exe
                IsDebug            : False
                IsPatched          : False
                IsPrivateBuild     : False
                IsPreRelease       : False
                IsSpecialBuild     : False
                Language           : Language Neutral
                LegalCopyright     : © 2011 Microsoft Corporation. All rights reserved.
                LegalTrademarks    : Microsoft® is a registered trademark of Microsoft Corporation.
                OriginalFilename   : ExSetup.exe
                PrivateBuild       :
                ProductBuildPart   : 513
                ProductMajorPart   : 14
                ProductMinorPart   : 3
                ProductName        : Microsoft® Exchange
                ProductPrivatePart : 0
                ProductVersion     : 14.03.0513.000
                SpecialBuild       :
                #>                
               
            } ;
        } else { 
            throw "No existing PSSession could be matched to:`n$_.name -like 'Exchange*'`nOpen a Remote PS Session into an onprem Exchange server and then run this command" ; 
        } ; 
    } ;  # PROC-E
    END {
         $VersString| write-output ;
    } ;  
} ; #*------^ END Function get-ExchangeOPConnectionVersion ^------