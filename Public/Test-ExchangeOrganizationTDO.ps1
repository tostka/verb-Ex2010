# Test-ExchangeOrganizationTDO.ps1

    #region TEST_EXCHANGEORGANIZATIONTDO ; #*------v Test-ExchangeOrganizationTDO v------
    function Test-ExchangeOrganizationTDO{
        <#
        .SYNOPSIS
        Test-ExchangeOrganizationTDO - Tests specified Exchange Organization within the local Forest 
        .NOTES
        Version     : 0.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 20250917-0114PM
        FileName    : Test-ExchangeOrganizationTDO.ps1
        License     : (none asserted)
        Copyright   : (none asserted)
        Github      : https://github.com/tostka/verb-io
        Tags        : Powershell,ActiveDirectory,Forest,Domain
        AddedCredit : Michel de Rooij / michel@eightwone.com
        AddedWebsite: http://eightwone.com
        AddedTwitter: URL        
        REVISIONS
        * 1:14 PM 9/17/2025 port to vx10 from xopBuildLibrary; add CBH, and Adv Function specs
        Only used for Ex version upgrades, schema & domain updates; not needed routinely, parking a copy in uwes as a _func.ps1 for loading when needed.
        .DESCRIPTION
        Test-ExchangeOrganizationTDO - Tests specified Exchange Organization within the local Forest
        .PARAMETER Organization
        Exchange Organization name                
        .INPUTS
        None, no piped input.
        .OUTPUTS
        System.String local ForestRoot 
        .EXAMPLE 
        PS> Write-MyOutput 'Checking Exchange organization existence'
        PS> $EX2016_MINFORESTLEVEL          = 15317
        PS> $EX2016_MINDOMAINLEVEL          = 13236
        PS> $EX2019_MINFORESTLEVEL          = 17000
        PS> $EX2019_MINDOMAINLEVEL          = 13236
        PS> If( $null -ne ( Test-ExchangeOrganizationTDO $Organization)) {
        PS>     Write-MyOutput "No existing Org in Forest" ; 
        PS> } Else {
        PS>     Write-MyOutput 'Organization exist; checking Exchange Forest Schema and Domain versions'
        PS>     $forestlvl= Get-ExchangeForestLevelTDO
        PS>     $domainlvl= Get-ExchangeDomainLevelTDO
        PS>     Write-MyOutput "Exchange Forest Schema version: $forestlvl, Domain: $domainlvl)"
        PS>     $MinFFL= $EX2016_MINFORESTLEVEL
        PS>     $MinDFL= $EX2016_MINDOMAINLEVEL
        PS>     If(( $forestlvl -lt $MinFFL) -or ( $domainlvl -lt $MinDFL)) {
        PS>         Write-MyOutput "Exchange Forest Schema or Domain needs updating (Required: $MinFFL/$MinDFL)"
        PS>     } Else {
        PS>         Write-MyOutput 'Active Directory looks already updated'
        PS>     }
        PS> }
        .LINK
        https://github.org/tostka/verb-Network/
        #>
        [CmdletBinding()]
        [alias('Test-ExchangeOrganization821')]
        PARAM(
            [Parameter(HelpMessage = "Exchange Organization Name to be tested")]
                [string]$Organization
        ) ;
        $CNC= Get-ForestConfigurationNCTDO
        return( [ADSI]"LDAP://CN=$Organization,CN=Microsoft Exchange,CN=Services,$CNC")
    } ; 
    #endregion TEST_EXCHANGEORGANIZATIONTDO ; #*------^ END Test-ExchangeOrganizationTDO ^------