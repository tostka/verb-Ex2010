#*------v get-ExRootSiteOUs.ps1 v------
#region GET_EXROOTSITEOUS ; #*------v get-ExRootSiteOUs v------
function get-ExRootSiteOUs {
    <#
    .SYNOPSIS
    get-ExRootSiteOUs.ps1 - Gather & return array of objects for root OU's matching a regex filter on the DN (if target OUs have a consistent name structure)
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-07-30
    FileName    : get-ExRootSiteOUs.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell,Exchange,ExchangeOnline,ActiveDirectory
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    # 1:10 PM 1/30/2026 add support for discovery of _Migration sub roots 
    # 12:34 PM 8/4/2021 ren'd getADSiteOus -> get-ExRootSiteOUs (avoid overlap with verb-adms\get-ADRootSiteOus())
    # 12:49 PM 7/25/2019 get-ExRootSiteOUs:updated $RegexBanned to cover TAC (no users or DL resource 
        OUs - appears to be variant of LYN w a single disabled users (obsolete disabled 
        TimH acct) 
    # 12:08 PM 6/20/2019 init vers
    .DESCRIPTION
    get-ExRootSiteOUs.ps1 - Gather & return array of objects for root OU's matching a regex filter on the DN (if target OUs have a consistent name structure)
    .DESCRIPTION
    Convert the passed-in ADUser object RecipientType from RemoteUserMailbox to RemoteSharedMailbox.
    .PARAMETER  Regex
    OU DistinguishedName regex, to identify 'Site' OUs [-Regex [regularexpression]]
    .PARAMETER RegexBanned
    OU DistinguishedName regex, to EXCLUDE non-legitimate 'Site' OUs [-RegexBanned [regularexpression]]
    .PARAMETER ShowDebug
    Parameter to display Debugging messages [-ShowDebug switch]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns an array of Site OU distinguishedname strings
    .EXAMPLE
    $SiteOUs=get-ExRootSiteOUs ;
    Retrieve the DNS for the default SiteOU
    .LINK
    https://github.com/tostka/verb-exo
    #>
    [CmdletBinding()]
    ##[Alias('ulu')]
    Param(
        [Parameter(Position = 0, HelpMessage = "OU DistinguishedName regex, to identify 'Site' OUs [-ADUser [regularexpression]]")]
            [ValidateNotNullOrEmpty()]
            [string]$rgxTTCRootOUs = '^OU=(\w{3}|PACRIM),DC=global,DC=ad,DC=toro((lab)*),DC=com$',
        [Parameter(Position = 0, HelpMessage = "OU DistinguishedName regex, to EXCLUDE non-legitimate 'Site' OUs [-RegexBanned [regularexpression]]")]
            [ValidateNotNullOrEmpty()]
            [string]$rgxTTCBannedRootOUs = '^OU=(BCC|EDC|NC1|NDS|TAC),DC=global,DC=ad,DC=toro((lab)*),DC=com$',
        [Parameter(Position = 0, HelpMessage = "OU DistinguishedName regex, to identify 'Site' OUs [-ADUser [regularexpression]]")]
            [ValidateNotNullOrEmpty()]
            [string]$rgxMIGRRootOUs = '^OU=(\w{3}|PACRIM),OU=_MIGRATIONS,DC=global,DC=ad,DC=toro,DC=com$',
        [Parameter(Position = 0, HelpMessage = "OU DistinguishedName regex, to EXCLUDE non-legitimate 'Site' OUs [-RegexBanned [regularexpression]]")]
            [ValidateNotNullOrEmpty()]
            [string]$rgxMigrBannedRootOUs = '^OU=(BCC|EDC|NC1|NDS|TAC),DC=global,DC=ad,DC=toro((lab)*),DC=com$',
        [Parameter(HelpMessage = "Domain Controller [-domaincontroller server.fqdn.com]")]
            [string] $domaincontroller,
        [Parameter(HelpMessage = "Debugging Flag [-showDebug]")]
            [switch] $showDebug
    ) # PARAM BLOCK END
    $verbose = ($VerbosePreference -eq "Continue") ; 
    $error.clear() ;
    TRY {
        if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
        $SiteOUs = @() ; 
        write-verbose "Retrieve net TTC SiteOUs..." ; 
        $SiteOUs += Get-OrganizationalUnit |?{($_.distinguishedname -match $rgxTTCRootOUs) -AND ($_.distinguishedname -notmatch $rgxTTCBannedRootOUs) }|sort distinguishedname ; 
        write-verbose "Retrieve net _MIGRATION SiteOUs..." ; 
        $SiteOUs += Get-OrganizationalUnit |?{($_.distinguishedname -match $rgxMIGRRootOUs) -AND ($_.distinguishedname -notmatch $rgxMigrBannedRootOUs) }|sort distinguishedname ; 
    } CATCH {
        $ErrTrapd=$Error[0] ;
        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #-=-record a STATUSWARN=-=-=-=-=-=-=
        $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
        if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
        if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
        #-=-=-=-=-=-=-=-=
        $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
    } ; 
    if ($SiteOUs) {
        $SiteOUs | write-output ;
    } else {
        $smsg= "Unable to retrieve OUs matching specified rgx:`n$($rgxTTCRootOUs)|$()";
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $false | write-output ;
    }
} ; 
#endregion GET_EXROOTSITEOUS ; #*------^ END get-ExRootSiteOUs ^------
