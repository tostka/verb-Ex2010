# Get-ExchangeOrganizationTDO.ps1


#region GET_EXCHANGEORGANIZATIONTDO ; #*------v Get-ExchangeOrganizationTDO v------
function Get-ExchangeOrganizationTDO{
        <#
        .SYNOPSIS
        Get-ExchangeOrganizationTDO - Returns the Exchange Organization Name
        .NOTES
        Version     : 0.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 20250917-0114PM
        FileName    : Get-ExchangeOrganizationTDO.ps1
        License     : (none asserted)
        Copyright   : (none asserted)
        Github      : https://github.com/tostka/verb-ex2010
        Tags        : Powershell,ActiveDirectory,Forest,Domain
        AddedCredit : Michel de Rooij / michel@eightwone.com
        AddedWebsite: http://eightwone.com
        AddedTwitter: URL        
        REVISIONS
        * 1:14 PM 9/17/2025 port to vx10 from xopBuildLibrary; add CBH, and Adv Function specs
        .DESCRIPTION
        Get-ExchangeOrganizationTDO - Returns the Exchange Organization Name
                
        .INPUTS
        None, no piped input.
        .OUTPUTS
        System.String local ForestRoot 
        .EXAMPLE ; 
        PS> $ExOrg= Get-ExchangeOrganization
        .LINK
        https://github.com/tostka/verb-ex2010
        #>
        [CmdletBinding()]
        [alias('Get-ExchangeOrganization')]
        PARAM() ;
        $CNC= Get-ForestConfigurationNCTDO
        Try {
            $ExOrgContainer= [ADSI]"LDAP://CN=Microsoft Exchange,CN=Services,$CNC"
            $rval= ($ExOrgContainer.PSBase.Children | Where-Object { $_.objectClass -eq 'msExchOrganizationContainer' }).Name
        }
        Catch {
            Write-MyVerbose "Can't find Exchange Organization object"
            $rval= $null
        }
        return $rval
    }
#endregion GET_EXCHANGEORGANIZATIONTDO ; #*------^ END Get-ExchangeOrganizationTDO ^------

