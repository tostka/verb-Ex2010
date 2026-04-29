# new-MailboxShared.ps1

#region NEW_MAILBOXSHARED ; #*------v new-MailboxShared v------
function new-MailboxShared {
    <#
    .SYNOPSIS
    new-MailboxShared.ps1 - Create New Generic Mbx
    .NOTES
    Version     : 1.0.2
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-
    FileName    : new-MailboxShared.ps1
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite:	URL
    AddedTwitter:	URL
    REVISIONS
    * 4:38 PM 4/29/2026 added -BrandUPN (defaults $TRUE) & -NoOutput (defaults TRUE) params; 
        brought in ENVIRO_DISCOVER , RV_ENVIRO, RV_TRANSCRIPT & SVCCONN_LITE, CALL_CONNECT_OPSERVICES 
        Added code to do accurate conversion of dname into alias, Dir/localPart eml addr, UPN, CName. Revised UPN email clean/build on the returned MailboxObject.upn (vs redund pull aduser.upn)
        Implimented external verb-ex2010\Repair-xopEmailAddresses to do emailaddress cleanup; 
        implemented -BrandUPN to clean the UPN dirname, and use the existing SMTP: psmtp domain as upn.
        implemented -NoOutput to suppress the new default emit of the updated mailbox 
        object to pipeline (to permit return of the full object to calling 
        new-MailboxGeneric, which in turn hands it off to new-xopMailboxSREtdo.ps1 
    * 9:38 AM 4/23/2026 transhol requires full rvEnv, spliced back ;  spliced over latest begin/svc conn, no-trans block
    * 2:02 PM 4/21/2026 #2650, added `n to bump givenname down off of the INFO: header line
    * 3:00 PM 4/15/2026 UPDATED SAMACCOUNTNAME & NAME (LDAP CN) TO STRICT STRIP - ALPHANUMS LEAVE UNRECOGNIZABLE OBJECTS!
        fixed Cu5 rgx error: doesn't accomodate any migr domains, just TTC, rem'd out forces 
        cu5 to take precedence (it's tested via EAP ch3eck up top anyway) 
        Updated ADDesc to use $env:username, instead of hardcode initials.  
    * 3:34 PM 4/14/2026 added dbg ipmo -fo -verb non-mod function testing;       
    * 6:42 PM 4/9/2026 extended revising f'ery, around locating - manually polling, 
        often creating - missing or non-standard OU storage for shared, room/equip & 
        secgrps for email access - and then updated this, new-mailboxgenerictor; 
        get-sitembxou; and new-mailboxshared to accomodate the migration blank-ups 
    * 5:15 PM 4/3/2026 substantial recode to support non-standard OU names (LF) in migrations OUs. 
    * 4:18 PM 3/30/2026 fixed borked $FallBackBaseUserOU typo; updated latest START_LOG_HOLISTIC
    * 4:17 PM 1/30/2026 fixed $odn owner.mbx.dn bug
    * 9:17 AM 1/29/2026 Cleanup ; EXIT ; -> Cleanup ; BREAK ;
    * 12:41 PM 1/27/2026 latest conn_svcs block updated
    * 2:48 PM 1/19/2026 -whatif's find ; 
    * 10:48 AM 1/19/2026 bugfix: $pltCcOPSvcs.UserRole (postfilter, not match test)
    # 12:20 PM 1/16/2026 check out against AAD->MG migr mandate, bring in latest logging & SERVICE_CONNECTIONS blocks
    # 2:46 PM 1/24/2025 add support for $OfficeOverride = 'Pune, IN' ; support for Office that doesn't match SITE OU code: $OfficeOverride = 'Pune, IN' ; 
    # 10:56 AM 4/12/2024 fix: echo typo 889:FIRSTNAME -> LASTNAME
    # 2:36 PM 8/2/2023 have to bump up password complexity - revised policy., it does support fname.lname naming & email addreses, just have to pass in dname with period. but the dname will also come out with the same period (which if they specified the eml, implies they don't mind if the name has it)
    # 10:30 AM 10/13/2021 pulled [int] from $ticket , to permit non-numeric & multi-tix
    # 10:01 AM 9/14/2021 had a random creation bug - but debugged fine in ISE (bad PSS?), beefed up Catch block outputs, captured new-mailbox output & recycled; added 7pswhsplat outputs prior to cmds.
    # 4:37 PM 5/18/2021 fixed broken start-log call (wasn't recycling logspec into logfile & transcrpt)
    # 11:10 AM 5/11/2021 swapped parentpath code for dyn module-support code (moving the new-mailboxgenerictor & add-mbxaccessgrant preproc .ps1's to ex2010 mod functions)
    # 1:52 PM 5/5/2021 added dot-divided displayname support (split fname & lname for generics, to auto-gen specific requested eml addresses) ; diverted parentpath log pref to d: before c: w test; untested code
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods
    # 1:29 PM 4/23/2020 updated dynpath & logging, unwrapped loadmod,
    # 4:28 PM 4/22/2020 updated logging code, to accomodate dynamic locations and $ParentPath
    # 4:36 PM 4/8/2020 works fully on jumpbox, but ignores whatif, renamed $bwhatif -> $whatif (as the b variant was prev set in the same-script, now separate scopes); swapped out CU5 switch, moved settings into infra file, genericized
    # 2:15 PM 4/7/2020 updated to reflect debugging on jumpbox
    # 2:35 PM 4/3/2020 new-MailboxShared: genericized for pub, moved material into infra, updated hybrid mod loads, cleaned up comments/remmed material ; updated to use start-log, debugged to funciton on jumpbox, w divided modules ; added -ParentPath to pass through a usable path for start-log, within new-mailboxshared()
    # 8:48 AM 11/26/2019 new-MailboxShared():moved the Office spec from $MbxSplat => $MbxSetSplat. New-Mailbox syntax set that supports -Shared, doesn't support -Office
    # 12:14 PM 10/4/2019 splice in Room/Equip code from new-mailboxConfRm.ps1's variant (not functionalized yet), added new -Room & -Equipement flags to trigger ConfRm code
    # 2:22 PM 10/1/2019 2076 & 1549 added: $FallBackBaseUserOU, OU that's used when can't find any baseuser for the owner's OU, default to a random shared from SITECODE (avoid crapping out):
    # 9:48 AM 9/27/2019 new-MailboxShared:added `a 'beep' to YYY prompt
    # 1:47 PM 9/20/2019 switched to auto mailbox distribution (ALL TO SITECODE dbs)
    # 12:06 PM 6/12/2019 functionalized version new-MailboxGenericTOR.ps1 => new-MailboxShared
    # 8:04 AM 5/31/2019 2441 exempted EAS info reporting from shared/generic mailboxes (EAS isn't core supported by EXO, no sense pointing out it is onprem)
    # 11:17 AM 5/8/2019 spliced in updated hybrid baseuser resolve code from new-mailboxConfRm.ps1
    # 8:41 AM 4/30/2019 #2327: add lab support for UPN build fr email addr:
    # 2:19 PM 4/29/2019 add $($DomTOLfqdn) to the domain param validateset on get-gcfast copy (sync'd in from verb-ex2010.ps1 vers), lab psv2 LACKS the foreign-lang cleanup funtions below, exempt use on psv2
    # 10:39 AM 4/1/2019 added: Remove-StringDiacritic, Remove-StringLatinCharacters(), and purging non-AAD-syncable chars from upn, mail alias etc. Ran a test pass. Also pre-cleaning the fname/lname/dname before using it to build samaccountname. Debugged & ran fine, creating a mbx.
    # 11:15 AM 2/15/2019 copied in bug-fixed write-log() with fixed debug support
    # 10:41 AM 2/15/2019 updated write-log to latest deferring version
    # 2:26 PM 12/13/2018 SITECODE has all mbxs moved out, need to poll remotembxs for baseuser!
    # 10:00 AM 11/20/2018 major update, switched 99% of write-xxx to write-log support, so it now produces a realtime 'log' of the build of the mailbox. Better than transcript because it still logs the changes right up to crashes. And it's only the _relevant) changes.
    # downside: lacks color coding unless I want to code in WARNs, which would be logged as warns.
    # 12:04 PM 7/18/2018 made display of password conditional (!shared)
    # 10:28 AM 6/27/2018 add $domaincontroller param option - skips dc discovery process and uses the spec, also updated $findOU code to work with TOL dom
    # 10:59 AM 5/21/2018 Fixed broken -nongeneric $true functionality: (corrected samaccountname gen code for blank fname field). Also added Mailuser/RemoteMailbox support for -Owner value. Validated functional for creation of the LYNCmonacct1 mbx.
    # 11:25 AM 12/22/2017 missing casmailbox splat construct for psv2 section, update CU5 test regx for perrotde & perrotpl, output an error when it fails to find a BaseUser (new empty site with empty target OU to draw from, prompts to hand spec -BaseUser param with mbx in another OU or loc
    # 11:58 AM 11/15/2017 1321: accommodate EXO-hosted Owners by testing with get-remotemailbox -AND get-mailbox on the owner spec. Created a mailbox, seemed to work. Not sure of access grant script yet.
    # 11:43 AM 10/6/2017 made $Cu5 Mandatry:$false
    # 10:30 AM 10/6/2017 fix typo in $cu5 switch
    # 8:41 AM 10/6/2017 major re-splict to read lost set-mailbox (alt domain assignment) & set-casmailbox material. Need to splice this code
            into the other new-mailbox based scripts. This explains why nothing has been setting owner-based domains in recent (months)?.
    # 3:22 PM 10/4/2017 midway through adding CU5 support
    # 2:38 PM 6/22/2017 LastName too, strip all names back
    # 2:30 PM 6/22/2017 Mbx.Name attrib appears to be a 64char limit!
    # 12:49 PM 5/31/2017 add $NoPrompt
    # 8:18 AM 5/9/2017 fixed minotr #region/#endregion typos
    # 1:29 PM 4/3/2017 #1764: this should have the -server spec
    # 2:39 PM 3/21/2017 spliced over sitemailboxOU() from new-mailboxconfrm.ps1
    # 8:03 AM 3/16/2017 suppress error make sure the $($script:ExPSS).ID  is EVEN POPULATED!
    # 7:56 AM 3/16/2017 gadu need a dawdle loop, also add -server $(InputSplat.domaincontroller)
    # 9:48 AM 3/2/2017 merged in updated Add-EMSRemote Set
    # 1:30 PM 2/27/2017 neither vscan cu9 nor owner's cu5 values got properly populated.
    # 12:36 PM 2/27/2017 get-SiteMbxOU(): fixed to cover breaks frm AD reorg OU name changes, Generics are all now in a single OU per site
    # 1:04 PM 2/24/2017 tweak below
    #12:56 PM 2/24/2017 doesn't run worth a damn SITECODE-> $($ADSiteCodeAU)/$($ADSiteCodeUK), force it to abort (avoid half-built remote objects that take too long to replicate back to SITECODE)
    # 11:37 AM 2/24/2017 Reporting loop: add RecipientType,RecipientTypeDetails
    # 11:35 AM 2/24/2017 ran initial debug pass, may work.
    # # 11:09 AM 2/24/2017 DMG gone: switch generics to real shared mbxs "Shared" = $True / $Inputsplat.shared
    #* 9:11 AM 9/30/2016 added pretest if(get-command -name set-AdServerSettings -ea 0)
    # 10:38 AM 6/7/2016 tested debugged 378194, generic creation. Now has new UPN set based on Primary SMTP dirname@DOMAIN.com.
    # 8:17 AM 6/7/2016 fixed to missing )'s in the splat dummy refactor bloc
    # 7:51 AM 6/7/2016 roughed in retries, and if/then cleanupexit blocks to make more fault tolerant. Needs debugging
    # 12:51 PM 5/10/2016 updated debug BP blocks
    # 2:07 PM 4/8/2016 minor tweaking
    # 12:32 PM 4/8/2016 submain: added validation that an existing $script:ExPSS is actually functional, or forces an Add-Emsremote
    # 11:28 AM 4/8/2016 passed EMSRemote dynamic
    # 11:25 AM 4/8/2016 passed initial test on Ex EMS local
    # 11:22 AM 4/8/2016 I think I've finally got it properly managing EMSRemote, purging redudant, and outputing functional report. needs to be tested in Ex EMS local
    # 12:31 PM 4/6/2016 it's crapping out in local EMS on SITECODE-3V6KSY1 Add-EMSRemote isn't picking up on the existing verbs, and noclobber etc.
    # 12:29 PM 4/6/2016 validated Generic in rEMS
    # 12:23 PM 4/6/2016: seems functional for testing.
        Added Validate-Password, and looping pw gen, to generate consistently compliant complexity.
        Debugged through a lot of inconsistencies. I think it can now serve as the base template for other scripts.
        needs testing to confirm that the new Add-EMSRemote will work in EMS, rEMS, and v2ps EMS on servers.
    # 12:28 PM 4/1/2016 synced all $whatif tests against the std $bWhatif, not a mix
    # 11:39 AM 3/31/2016 ren Manager/ManagedBy -> Owner in splats, dropping ManagedBy use on AD Objects
    # 2:32 PM 3/22/2016 rem out ManagedBy support (need to implement Owner)
    # 1:12 PM 2/11/2016 fixed new bug in get-GCFast, wasn't detecting blank $site
    # 12:20 PM 2/11/2016 updated to standard EMS/AD Call block & Add-EMSRemote()
    # 11:31 AM 2/11/2016 updated [ordered] to exempt psv2
    #10:49 AM 2/11/2016: updated get-GCFast to current spec, updated any calls for "-site 'lyndale'" to just default to local machine lookup
    # 7:40 AM Add-EMSRemote: 2/5/2016 another damn cls REM IT! I want to see all the connectivity info, switched wh->wv, added explicit echo's of what it's doing.
    # 11:08 AM 1/15/2016 re-updated Add-EMSRemote, using a -eq v -like with a wildcard string. Have to repush copies all over now. Also removed 2 Clear-Host's
    # 10:02 AM 1/13/2016: fixed cls bug due to spurious ";cls" included in the try/catch boilerplate: Write-Error "$((get-date).ToString('HH:mm:ss')): Command: $($_.InvocationInfo.MyCommand)" ;cls => Write-Error "$((get-date).ToString('HH:mm:ss')): Command: $($_.InvocationInfo.MyCommand)" ;
    # 1:48 PM 10/29/2015 fixed blank surname/givenname - was generic-only setting it. Also sub'd insput fname/lname for std firstname lastname field names (ported from new-mailboxCN.ps1)
    # 1:30 PM 10/29/2015: 780 get-aduser needs -server, fails:, also added -server to the get-AD* cmds missing it, and added XIA sitecode steering to the $OU find process
    # 2:07 PM 10/26/2015 had to split out the new-mailbox | set-mailbox, with a do/while wait for the mbx to be visible, in between, because the set wasn't finding a mbx, when it executed
    # 1:53 PM 10/26/2015 fixed failure to assign $InputSplat.SiteCode, for Generic mbxs
    # 12:24 PM 10/21/2015 added/debugged -Vscan YES|NO|null param. Created OEV\Generic\Test XXXOffboard mbx with it
    # 11:41 AM 10/21/2015 update clean-up param & help info
    # 11:32 AM 10/21/2015 613: added NonGeneric detect and trigger ResetPasswordOnNextLogon flag
    # 11:26 AM 10/21/2015 fixed rampant issues created by OEV's non-standard OU naming choices: had a Generic*Win7*Computers ou, that had to be one-off excluded, and also had GPOTest users that likewise had to be excluded to ensure only a single OU came through.
    # 9:40 AM 10/21/2015 #531:fix missing trailing )
    # 9:08 AM 10/14/2015 added debugpref maint code to get write-debug to work
    # 7:31 AM 10/14/2015 added -dc specs to all *-user & *-mailbox cmds, to ensure we're pulling back data from same dc that was updated in the set-* commands
    # 7:27 AM 10/14/2015 rplcd all get-timestamps -> $((get-date).ToString('HH:mm:ss'))
    # 1:12 PM 10/6/2015 updated to spec, looks functional
    # 10:49 AM 10/6/2015: updated vers of Get-AdminInitials
    * 2:48 PM 10/2/2015 updated Catch blocks to be specific on crash
    * 10:23 AM 10/2/2015 initial port from add-mbxaccessgrant & bp code
.DESCRIPTION
new-MailboxShared.ps1 - Create New Generic Mbx
.PARAMETER DisplayName
    Display Name for mailbox ["fname lname","genericname"]
    .PARAMETER MInitial
    Middle Initial for mailbox (for non-Generic)["a"]
    .PARAMETER Owner
    Specify the userid to be responsible for access-grant-approvals[name,emailaddr,alias]
    .PARAMETER SiteOverride
    Optionally specify a 3-letter Site Code. Used to force DL name/placement to vary from Owner's site)[3-letter Site code]
    .PARAMETER OfficeOverride
    Optionally specify an override Office value (assigned to mailbox Office, instead of SiteCode)['City, CN']
    .PARAMETER BaseUser
    Optionally specify an existing mailbox upon which to base the new mailbox & OU settings[name,emailaddr,alias]
    .PARAMETER Room
    Optional parameter indicating new mailbox Is Room-type[-Room `$true]
    .PARAMETER Equip
    Optional parameter indicating new mailbox Is Equipment-type[-Equip `$true]
    .PARAMETER NonGeneric
    Optionally specify new mailbox Is NON-Generic-type (defaults $false)[-NonGeneric $true]
    .PARAMETER IsContractor
    Parameter indicating new mailbox belongs to a Contractor[-IsContractor switch]
    .PARAMETER Vscan
    Parameter indicating new mailbox will have Vscan access (prompts if not specified)[-Vscan Yes|No|Null]
    .PARAMETER Cu5
    Optionally force CU5 (variant domain assign) [-Cu5 Exmark]
    .PARAMETER ticket
    Incident number for the change request[[int]nnnnnn]
    .PARAMETER domaincontroller
    Option to hardcode a specific DC [-domaincontroller xxxx]
    .PARAMETER ParentPath
    Calling script path (used for log construction)[-ParentPath c:\pathto\script.ps1]
    .PARAMETER NoPrompt
    Suppress YYY confirmation prompts
    .PARAMETER BrandUPN
    Switch to not force UPN to central suffix (uses primarysmtpaddr sdomain, if avail forest upn suffix; otherwise uses central forestdom)
    .PARAMETER NoOutput
    NoOutput Flag [-NoOutput]
    .PARAMETER Whatif
    Parameter to run a Test no-change pass, and log results [-Whatif switch]
    .PARAMETER ShowDebug
    Parameter to display Debugging messages [-ShowDebug switch]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    new-MailboxShared.ps1 -ticket "355925" -DisplayName "XXX Confirms"  -MInitial ""  -Owner "LOGON" -NonGeneric $true -showDebug -whatIf ;
    Testing syntax with explicit BaseUSer specified, Whatif test & Debug messages displayed:
    .EXAMPLE
    new-MailboxShared.ps1 -ticket "355925" -DisplayName "XXX Confirms"  -MInitial ""  -Owner "LOGON" -BaseUser "AccountsReceivable" -NonGeneric -showDebug -whatIf ;
    Testing syntax with no explict BaseUSer specified (draws random from Generic OU of Owner's Site), Whatif test & Debug messages displayed:
    .EXAMPLE
    #-=-=-=-=-=-=-=-=
    $logging = $True ; # need to set in scope outside of functions
    $pltInput=[ordered]@{} ;
    if($DisplayName){$pltInput.add("DisplayName",$DisplayName) } ;
    if($MInitial){$pltInput.add("MInitial",$MInitial) } ;
    if($Owner){$pltInput.add("Owner",$Owner) } ;
    if($BaseUser){$pltInput.add("BaseUser",$BaseUser) } ;
    if($IsContractor){$pltInput.add("IsContractor",$IsContractor) } ;
    if($Room){$pltInput.add("Room",$Room) } ;
    if($Equip){$pltInput.add("Equip",$Equip) } ;
    if($Ticket){$pltInput.add("Ticket",$Ticket) } ;
    if($domaincontroller){$pltInput.add("domaincontroller",$domaincontroller) } ;
    if($NoPrompt){$pltInput.add("NoPrompt",$NoPrompt) } ;
    if($showDebug){$pltInput.add("showDebug",$showDebug) } ;
    if($whatIf){$pltInput.add("whatIf",$whatIf) } ;
    # only reset from defaults on explicit -NonGeneric $true param
    if($NonGeneric -eq $true){
        # switching over generics to real 'shared' mbxs: "Shared" = $True
    } else {
        # force it if not true
        $NonGeneric=$false;
    } ;
    if($NonGeneric){$pltInput.add("NonGeneric",$NonGeneric) } ;
    if ($Vscan){
        if ($Vscan -match "(?i:^(YES|NO)$)" ) {
            $Vscan = $Vscan.ToString().ToUpper() ;
            if($Vscan){$pltInput.add("Vscan",$Vscan) } ;
        } else {
            $Vscan = $null ;
            #$pltInput.Vscan=$Vscan ;
            # 1:32 PM 2/27/2017 force em  on all, no reason not to have external email!
            if($Vscan){$pltInput.add("Vscan","YES") } ;
        }  ;
    }; # If not explicit yes/no, prompt for input
    # Cu5 override support (normally inherits from assigned owner/manager)
    if ($Cu5){
        if($Cu5){$pltInput.add("Cu5",$Cu5) } ;
    } else {
        $pltInput.add("Cu5",$null) ;
    }  ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):new-MailboxShared w`n$(($pltInput|out-string).trim())" ;
    new-MailboxShared @pltInput
    CleanUp ;
    #-=-=-=-=-=-=-=-=
    Full prod call code (from new-MailboxGenericTOR.ps1)
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,HelpMessage="Display Name for mailbox [fname lname,genericname]")]
          [string]$DisplayName,
        [Parameter(HelpMessage="Middle Initial for mailbox (for non-Generic)[a]")]
            [string]$MInitial,
        [Parameter(Mandatory=$true,HelpMessage="Specify the userid to be responsible for access-grant-approvals[name,emailaddr,alias]")]
            [string]$Owner,
        [Parameter(HelpMessage="Optionally a specific existing mailbox upon which to base the new mailbox settings (default is to draw a random mbx from the target OU)[name,emailaddr,alias]")]
            [string]$BaseUser,
        [Parameter(HelpMessage="Optional parameter indicating new mailbox Is Room-type[-Room `$true]")]
            [bool]$Room,
        [Parameter(HelpMessage="Optional parameter indicating new mailbox Is Equipment-type[-Equip `$true]")]
            [bool]$Equip,
        [Parameter(HelpMessage="Optional parameter indicating new mailbox Is NonGeneric-type[-NonGeneric `$true]")]
            [bool]$NonGeneric,
        [Parameter(HelpMessage="Optional parameter indicating new mailbox belongs to a Contractor[-IsContractor switch]")]
            [switch]$IsContractor,
        [Parameter(HelpMessage="Optional parameter controlling Vscan (CU9) access (prompts if not specified)[-Vscan YES|NO|NULL]")]
            [string]$Vscan="YES",
        [Parameter(Mandatory=$false,HelpMessage="Optionally force CU5 (variant domain assign) [-Cu5 Exmark]")]
            [string]$Cu5,
        [Parameter(HelpMessage="Optionally specify a 3-letter Site Code o force OU placement to vary from Owner's current site[3-letter Site code]")]
            [string]$SiteOverride,
        # 2:49 PM 1/24/2025 add support for Office that doesn't match SITE OU code: $OfficeOverride = 'Pune, IN' ; 
        [Parameter(HelpMessage="Optionally specify an override Office value (assigned to mailbox Office, instead of SiteCode)['City, CN']")]
            [string]$OfficeOverride,
        [Parameter(Mandatory=$true,HelpMessage="Incident number for the change request[[int]nnnnnn]")]
            # [int] # 10:30 AM 10/13/2021 pulled, to permit non-numeric & multi-tix
            $Ticket,
        [Parameter(HelpMessage="Option to hardcode a specific DC [-domaincontroller xxxx]")]
            [string]$domaincontroller,
        [Parameter(HelpMessage="Calling script path (used for log construction)[-ParentPath c:\pathto\script.ps1]")]
            [string]$ParentPath,
        [Parameter(HelpMessage="Suppress YYY confirmation prompts [-NoPrompt]")]
            [switch] $NoPrompt,
        [Parameter(HelpMessage="Switch to not force UPN to central suffix (uses primarysmtpaddr sdomain, if avail forest upn suffix; otherwise uses central forestdom)[-BrandUPN]")]
            [switch]$BrandUPN = $true,
        [Parameter(HelpMessage='NoOutput Flag [-NoOutput]')]
            [switch]$NoOutput=$true,
        [Parameter(HelpMessage='Debugging Flag [-showDebug]')]
          [switch] $showDebug,
        [Parameter(HelpMessage='Whatif Flag [$switch]')]
          [switch] $whatIf
    ) ;
    BEGIN{
        #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
        #region RV_ENVIRO ; #*------v RV_ENVIRO v------
        $pltRvEnv=[ordered]@{
            PSCmdletproxy = $rPSCmdlet ; 
            PSScriptRootproxy = $rPSScriptRoot ; 
            PSCommandPathproxy = $rPSCommandPath ; 
            MyInvocationproxy = $rMyInvocation ;
            PSBoundParametersproxy = $rPSBoundParameters
            verbose = [boolean]($PSBoundParameters['Verbose'] -eq $true) ; 
        } ;
        write-verbose "(Purge no value keys from splat)" ; 
        $mts = $pltRVEnv.GetEnumerator() |?{ ($_.value -eq $null) -OR ($_.value -eq '') -OR ($_.value.length -eq 0)} ; $mts |foreach-object{$pltRVEnv.remove($_.Name)} ; remove-variable mts -ea 0 ; 
        $smsg = "resolve-EnvironmentTDO w`n$(($pltRVEnv|out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        if(get-command resolve-EnvironmentTDO -ea STOP){}ELSE{
            $smsg = "UNABLE TO gcm resolve-EnvironmentTDO!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            BREAK ; 
        } ; 
        $rvEnv = resolve-EnvironmentTDO @pltRVEnv ; 
        $smsg = "`$rvEnv returned:`n$(($rvEnv |out-string).trim())" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        <#
        #region RV_TRANSCRIPT ; #*------v RV_TRANSCRIPT v------
        # SIMPLE AUTORESOLVE $TRANSCRIPT DOESN'T DO BUDRV TESTS OR MODDIR REDIR; INITIAL TESTING WORKS
        if ($rvEnv.isScript) {
            if ($rvEnv.PSCommandPathproxy) { 
                $ps1Path = $rvEnv.PSCommandPathproxy
                $CmdName= split-path $rvEnv.PSCommandPathproxy -ea 0 -leaf ;
                $CmdParentDir = split-path $rvEnv.PSCommandPathproxy -ea 0 ;
                $CmdNameNoExt = [system.io.path]::GetFilenameWithoutExtension($rvEnv.PSCommandPathproxy) ;
            }
            elseif ($script:PSCommandPath) { 
                $ps1Path = $script:PSCommandPath  ; 
                $CmdName= split-path $script:PSCommandPath -ea 0 -leaf ;
                $CmdParentDir = split-path $script:PSCommandPath -ea 0 ;
                $CmdNameNoExt = [system.io.path]::GetFilenameWithoutExtension($script:PSCommandPath) ;
            }
            elseif ($rPSCommandPath) { 
                $ps1Path = $rPSCommandPath ; 
                $CmdName= split-path $rPSCommandPath -ea 0 -leaf ;
                $CmdParentDir = split-path $rPSCommandPath -ea 0 ;
                $CmdNameNoExt = [system.io.path]::GetFilenameWithoutExtension($rPSCommandPath) ;
            } ;
        } elseif ($rvEnv.isFunc) {
            if ($rvEnv.FuncDir -AND $rvEnv.FuncName) { 
                $fctName = $rvEnv.FuncName  ; 
                $CmdName= $rvEnv.FuncName ;
                $CmdParentDir = $rvEnv.FuncDir ;
                $CmdNameNoExt = $rvEnv.FuncName ;            
            } ;
        } elseif ($CmdletName) { $CmdName = $rvEnv.CmdletName };
        if($CmdParentDir -AND $CmdNameNoExt){
            $transcript = (join-path -path $CmdParentDir -ChildPath 'logs') ;
            if(-not (test-path $transcript  -PathType Container -ea 0)){ mkdir $transcript -verbose } ;
            $transcript = join-path -path $transcript -childpath $CmdNameNoExt ;      
            #if($ticket){$transcript += "-$($ticket)" }
            #if($whatif -OR $WhatIf.IsPresent){$transcript += "-WHATIF"}ELSE{$transcript += "-EXEC"} ;
            #if($thisXoMbx.userprincipalname){$transcript += "-$($thisXoMbx.userprincipalname)"} ;
            #$transcript += "-LASTPASS-trans-log.txt" ;        
            $transcript += "-$(get-date -format 'yyyyMMdd-HHmmtt')-trans-log.txt" ;
        }else{ write-warning "INSUFFICIENT ENVIRONMENT INPUTS TO AUTOBUILD A STRANSCRIPT!" } ; 
        # tag demo (doesn't do budrv redir etc, use start-log etc for full cover)
        #if($transcript -AND $CmdName -AND $DisplayName){
        #    $transcript = $transcript -replace $cmdName,"$($CmdName)-$($DisplayName)" ;
        #}
        #endregion RV_TRANSCRIPT ; #*------^ END RV_TRANSCRIPT ^------
        #>
        #endregion RV_ENVIRO ; #*------^ END RV_ENVIRO ^------
    
        #endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------
        
        #region MODULES_FORCE_LOAD ; #*------v MODULES_FORCE_LOAD v------
        # core modes in dep order
        $tmods = @('verb-IO', 'verb-Text', 'verb-logging', 'verb-Desktop', 'verb-dev', 'verb-Mods', 'verb-Network', 'verb-Auth', 'verb-ADMS', 'VERB-ex2010', 'verb-EXO', 'VERB-mg') ;
        # task mods in dep order
        $tmods += @('ExchangeOnlineManagement', 'ActiveDirectory', 'Microsoft.Graph.Users')
        $oWPref = $WarningPreference ; $WarningPreference = 'SilentlyContinue' ;
        $tmods | % { $thismod = $_ ; TRY { $thismod | ipmo -fo  -ea STOP }CATCH { write-host -foregroundcolor yellow "Missing module:$($thismod)`ntrying find-module lookup..." ; find-module $thismod } } ; $WarningPreference = $oWPref ;
        #endregion MODULES_FORCE_LOAD ; #*------^ END MODULES_FORCE_LOAD ^------

        #region CONSTANTS_AND_ENVIRO ; #*======v CONSTANTS_AND_ENVIRO v======
        #region LOCAL_CONSTANTS ; #*------v LOCAL_CONSTANTS v------

        # SamAccountname banned characters + space, 
        $rgxBanSamA = '[\"\/\\\[\]\:\;\|\=\,\+\*\?\<\>\s]' ; 
        # alias has a separate ban list from SamAccountName! use of the samaactname as alias won't resolve to the user.
        #$rgxBanMailAlias = "[\ @!\#\$%\^&\*\(\)\+=\{}\[\]\|\\:;"'<>,\?/]"
        $rgxPermitMailAlias = "[a-zA-Z0-9._-]" ; # chars from cleaned dname that are joined to create the alias/aduser.mailnickname - the script *sets* the alias, overrides normal Exchange sanitizing of the auto-gen'd alias!
        # Email/UPN localpart/directoryname ban chars
        $rgxBanEmlDirName = "[\s@!\#\$%\^&\*\(\)\+=\{}\[\]\|\\:;','<>,\?/]" ; 
        # Name/LDAP cn= name, banned characters
        $rgxBanCNName = '[\,\+\"\\\<\>\;\=\/]' 
        $dbgDate = '4/27/2026'; # debugging ipmo force loads variants not in modules
        $Retries = 4 ; # number of re-attempts
        $RetrySleep = 5 ; # seconds to wait between retries
        # $rgxCU5 = [infra file]
        # OU that's used when can't find any baseuser for the owner's OU, default to a random shared from ($ADSiteCodeUS) (avoid crapping out):
        $FallBackBaseUserOU = "$($DomTORfqdn)/$($ADSiteCodeUS)/Generic Email Accounts" ;
        # 3:46 PM 4/3/2026 add Migrations OU variant support
        $rgxOUMigrations = ',OU=_MIGRATIONS,DC=global,DC=ad,DC=toro,DC=com$' ;
        $rgxMigationsSite = ',OU=(\w+),OU=_MIGRATIONS,DC=global,DC=ad,DC=toro,DC=com$' ;

        #endregion LOCAL_CONSTANTS ; #*------^ END LOCAL_CONSTANTS ^------
        #region COMMON_CONSTANTS ; #*------v COMMON_CONSTANTS v------
        if (-not $DoRetries) { $DoRetries = 4 } ;    # # times to repeat retry attempts
        if (-not $RetrySleep) { $RetrySleep = 10 } ; # wait time between retries
        if (-not $RetrySleep) { $DawdleWait = 30 } ; # wait time (secs) between dawdle checks
        if (-not $DirSyncInterval) { $DirSyncInterval = 30 } ; # AADConnect dirsync interval
        if (-not $ThrottleMs) { $ThrottleMs = 50 ; }
        if (-not $rgxDriveBanChars) { $rgxDriveBanChars = '[;~/\\\.:]' ; } ; # ;~/\.:,
        if (-not $rgxCertThumbprint) { $rgxCertThumbprint = '[0-9a-fA-F]{40}' } ; # if it's a 40char hex string -> cert thumbprint
        if (-not $rgxSmtpAddr) { $rgxSmtpAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ; } ; # email addr/UPN
        if (-not $rgxDomainLogon) { $rgxDomainLogon = '^[a-zA-Z][a-zA-Z0-9\-\.]{0,61}[a-zA-Z]\\\w[\w\.\- ]+$' } ; # DOMAIN\samaccountname
        if (-not $exoMbxGraceDays) { $exoMbxGraceDays = 30 } ;
        if (-not $XOConnectionUri ) { $XOConnectionUri = 'https://outlook.office365.com' } ;
        if (-not $SCConnectionUri) { $SCConnectionUri = 'https://ps.compliance.protection.outlook.com' } ;
        if (-not $XODefaultPrefix) { $XODefaultPrefix = 'xo' };
        if (-not $SCDefaultPrefix) { $SCDefaultPrefix = 'sc' };
        #$rgxADDistNameGAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 1 ) -join ',')"
        #$rgxADDistNameAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 2 ) -join ',')"
        #region TEST_METAS ; #*------v TEST_METAS v------
        # critical dependancy Meta variables
        $MetaNames = 'TOR', 'CMW', 'TOL' # ,'NOSUCH' ;
        # critical dependancy Meta variable properties
        $MetaProps = 'legacyDomain', 'o365_TenantDomain' #,'DOESNTEXIST' ;
        $isBased = $true ; $gvMiss = @() ; $ppMiss = @() ;
        foreach ($met in $metanames) {
            write-verbose "chk:`$$($met)Meta" ;
            if (-not (gv -name "$($met)Meta" -ea 0)) {
                $isBased = $false; $gvMiss += "$($met)Meta" ;
            } ;
            foreach ($mp in $MetaProps) {
                write-verbose "chk:`$$($met)Meta.$($mp)" ;
                #if(-not (gv -name "$($met)Meta" -ea 0).value[$mp]){ # testing has a value, not is present as a spec!
                if (-not (gv -name "$($met)Meta" -ea 0).value.keys -contains $mp) { $isBased = $false; $ppMiss += "$($met)Meta.$($mp)" ; } ;
            } ;
        } ;
        if ($gvmiss) { write-warning "Missing Dependant Meta variables:`n$(($gvMiss |%{"`$$($_)" }) -join ',')" } ;
        if ($ppMiss) { write-warning "Missing Dependant Meta vari properties:`n$(($ppMiss |%{"`$$($_)" }) -join ',')" } ;
        if (-not $isBased) { write-warning  "missing critical dependancy profile config!" } ;
        #endregion TEST_METAS ; #*------^ END TEST_METAS ^------
        # email trigger vari, it will be semi-delimd list of mail-triggering events
        $script:PassStatus = $null ;
        # TenOrg or other looped-specific PassStatus (auto supported by 7pswlt)
        #New-Variable -Name PassStatus_$($tenorg) -scope Script -Value $null ;
        [array]$SmtpAttachment = $null ;
        #write-verbose "start-Timer:Master" ;
        $swM = [Diagnostics.Stopwatch]::StartNew() ;
        # $ByPassLocalExchangeServerTest = $true # rough in, code exists below for exempting service/regkey testing on this variable status. Not yet implemented beyond the exemption code, ported in from orig source.
        #endregion COMMON_CONSTANTS ; #*------^ END COMMON_CONSTANTS ^------
        #region ENCODED_CONTANTS ; #*------v ENCODED_CONTANTS v------
        # ENCODED CONsTANTS & SUPPORT FUNCTIONS:
        #region 2B4 ; #*------v 2B4 v------
        if (-not (get-command 2b4 -ea 0)) { function 2b4 { [CmdletBinding()][Alias('convertTo-Base64String')] PARAM([Parameter(ValueFromPipeline = $true)][string[]]$str) ; PROCESS { $str | % { [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_)) } }; } ; } ;
        #endregion 2B4 ; #*------^ END 2B4 ^------
        #region 2B4C ; #*------v 2B4C v------
        # comma-quoted return
        if (-not (get-command 2b4c -ea 0)) { function 2b4c { [CmdletBinding()][Alias('convertto-Base64StringCommaQuoted')] PARAM([Parameter(ValueFromPipeline = $true)][string[]]$str) ; BEGIN { $outs = @() } PROCESS { [array]$outs += $str | % { [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_)) } ; } END { '"' + $(($outs) -join '","') + '"' | out-string | set-clipboard } ; } ; } ;
        #endregion 2B4C ; #*------^ END 2B4C ^------
        #region FB4 ; #*------v FB4 v------
        # DEMO: $SitesNameList = 'THluZGFsZQ==','U3BlbGxicm9vaw==','QWRlbGFpZGU=' | fb4 ;
        if (-not (get-command fb4 -ea 0)) { function fb4 { [CmdletBinding()][Alias('convertFrom-Base64String')] PARAM([Parameter(ValueFromPipeline = $true)][string[]]$str) ; PROCESS { $str | % { [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($_)) }; } ; } ; };
        #endregion FB4 ; #*------^ END FB4 ^------
        # FOLLOWING CONSTANTS ARE USED FOR DEPENDANCY-LESS CONNECTIONS
        if (-not $o365_Toroco_SIDUpn) { $o365_Toroco_SIDUpn = 'cy10b2RkLmthZHJpZUB0b3JvLmNvbQ==' | fb4 } ;
        $o365_SIDUpn = $o365_Toroco_SIDUpn ;
        switch ($env:Userdomain) {
            'CMW' {
                if (-not $CMW_logon_SID) { $CMW_logon_SID = 'Q01XXGQtdG9kZC5rYWRyaWU=' | fb4 } ;
                $logon_SID = $CMW_logon_SID ;
            }
            'TORO' {
                if (-not $TOR_logon_SID) { $TOR_logon_SID = 'VE9ST1xrYWRyaXRzcw==' | fb4 } ;
                $logon_SID = $TOR_logon_SID ;
            }
            $env:COMPUTERNAME {
                $smsg = "%USERDOMAIN% -EQ %COMPUTERNAME%: $($env:computername) => non-domain-connected, likely edge role Ex server!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent }
                else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                if ($WorkgroupName = (Get-WmiObject -Class Win32_ComputerSystem).Workgroup) {
                    $smsg = "WorkgroupName:$($WorkgroupName)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                }
                if (($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup')) -or (
                        $isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup')) -or
                    $ByPassLocalExchangeServerTest) {
                    $smsg = "We are on Exchange Server"
                    if ($verbose) {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else { write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $IsEdgeTransport = $false
                    if ((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole') -or (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole')) {
                        $smsg = "We are on Exchange Edge Transport Server"
                        if ($verbose) {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else { write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        $IsEdgeTransport = $true
                    } ;
                } else {
                    $isLocalExchangeServer = $false
                    $IsEdgeTransport = $false ;
                } ;
            } ;
            default {
                $smsg = "$($env:userdomain):UNRECOGIZED/UNCONFIGURED USER DOMAIN STRING!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent }
                else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                THROW $SMSG
                BREAK ;
            }
        } ;
        #endregion ENCODED_CONTANTS ; #*------^ END ENCODED_CONTANTS ^------

        #region FUNCTIONS_FULLYEXTERNAL ; #*======v FUNCTIONS_FULLYEXTERNAL v======
        # Optional block that relies on local module installs (vs the FUNCTIONS_LOCAL integrated block that follows below, and the FUNCTIONS_LOCAL_INTERNAL that is used for completely non-shared local functions.)

        #region RESOLVE_ENVIRONMENTTDO ; #*------v verb-io\resolve-EnvironmentTDO v------
        if(-not(gi function:resolve-EnvironmentTDO -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-io\resolve-EnvironmentTDO !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ;
        #endregion RESOLVE_ENVIRONMENTTDO ; #*------^ END verb-io\resolve-EnvironmentTDO ^------

        #region WRITE_LOG ; #*------v verb-logging\write-log v------
        if(-not(gi function:write-log -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-logging\write-log !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ; 
        #endregion WRITE_LOG ; #*------^ END verb-logging\write-log  ^------

        #region START_LOG ; #*------v verb-logging\Start-Log v------
        if(-not(gi function:start-log -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-logging\start-log !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ; 
        #endregion START_LOG ; #*------^ END verb-logging\start-log ^------

        #region RESOLVE_NETWORKLOCALTDO ; #*------v verb-Network\resolve-NetworkLocalTDO v------
        if(-not(gi function:resolve-NetworkLocalTDO -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-Network\resolve-NetworkLocalTDO!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        }
        #endregion RESOLVE_NETWORKLOCALTDO ; #*------^ END verb-Network\resolve-NetworkLocalTDO ^------

        #region PUSH_TLSLATEST ; #*------v verb-Network\push-TLSLatest v------
        if(-not(gi function:push-TLSLatest -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-Network\push-TLSLatest!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ; 
        #endregion PUSH_TLSLATEST ; #*------^ END verb-Network\push-TLSLatest ^------

        #region TEST_EXCHANGEINFO ; #*------v verb-Ex2010\test-LocalExchangeInfoTDO v------
        if(-not (get-item function:test-LocalExchangeInfoTDO -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-Ex2010\test-LocalExchangeInfoTDO!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ; 
        #endregion TEST_EXCHANGEINFO ; #*------^ END verb-Ex2010\test-LocalExchangeInfoTDO ^------

        #region CONNECT_O365SERVICES ; #*======v verb-exo\connect-O365Services v======
        if(-not (get-childitem function:connect-O365Services -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-exo\connect-O365Services!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ;
        #endregion CONNECT_O365SERVICES ; #*======^ END verb-exo\connect-o365services ^======

        #region OUT_CLIPBOARD ; #*------v verb-IO\out-Clipboard v------
        if(-not(gci function:out-Clipboard -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-IO\out-Clipboard!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ; 
        #endregion OUT_CLIPBOARD ; #*------^ END verb-IO\out-Clipboard ^------

        #region START_SLEEPCOUNTDOWN ; #*------v verb-IO\start-sleepcountdown v------
        if (-not (get-command start-sleepcountdown -ea 0)) {
            $smsg = "MISSING DEPENDANT: verb-IO\start-sleepcountdown!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ;
        #endregion START_SLEEPCOUNTDOWN ; #*------^ END verb-IO\start-sleepcountdown ^------

        #region CONVERTFROM_MARKDOWNTABLE ; #*------v verb-IO\convertFrom-MarkdownTable v------
        if(-not(gci function:convertFrom-MarkdownTable -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-IO\convertFrom-MarkdownTable!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ; 
        #endregion CONVERTFROM_MARKDOWNTABLE ; #*------^ END verb-IO\convertFrom-MarkdownTable ^------

        #region REMOVE_INVALIDVARIABLENAMECHARS ; #*------v verb-IO\Remove-InvalidVariableNameChars v------        
        if(-not (gcm Remove-InvalidVariableNameChars -ea 0)){
            Function Remove-InvalidVariableNameChars ([string]$Name) {
                ($Name.tochararray() -match '[A-Za-z0-9_]') -join '' | write-output ;
            };
        } ;
        #endregion REMOVE_INVALIDVARIABLENAMECHARS ; #*------^ END verb-IO\Remove-InvalidVariableNameChars ^------
    
        #endregion FUNCTIONS_FULLYEXTERNAL ; #*======^ END FUNCTIONS_FULLYEXTERNAL ^======
    
        #region FUNCTIONS_INTERNAL ; #*======v FUNCTIONS_INTERNAL v======


        #endregion FUNCTIONS_INTERNAL ; #*======^ END FUNCTIONS_INTERNAL ^======

        #region SUBMAIN ; #*======v SUB MAIN v======
        if(gcm push-TLSLatest -ea Continue ){push-TLSLatest}else{write-warning "MISSING PUSH-TLSLATEST! O365 MAY NOT ACCEPT LESSER TLS LEVELS!" } ; 
        #region BANNER ; #*------v BANNER v------
        $sBnr = "#*======v $($CmdName) : v======" ; 
        $smsg = $sBnr ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } #Error|Warn|Debug
        else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #endregion BANNER ; #*------^ END BANNER ^------
        #region SVCCONN_LITE ; #*------v SVCCONN_LITE v------
        $isXoConn = [boolean]( (gcm Get-ConnectionInformation -ea 0) -AND (Get-ConnectionInformation -ea 0 | ? { $_.State -eq 'Connected' -AND $_.TokenStatus -eq 'Active' })) ; if (-not $isXoConn) {
            Connect-EXO -silent:$($silent)
        } else { write-verbose "EXO connected" };
        #region MG_CONNECT ; #*------v MG_CONNECT v------
        #$RequiredScopes = "Directory.AccessAsUser.All",'Directory.ReadWrite.All','User.ReadWrite.All','AuditLog.Read.All','openid','profile','User.Read','User.Read.All','email' ;
        if(gcm test-mgconnection -ea 0){
            $MGConn = test-mgconnection -Verbose:($VerbosePreference -eq 'Continue') -silent:$($silent) -ea STOP;
            if ($RequiredScopes) { $addScopes = @() ; $RequiredScopes | foreach-object { $thisPerm = $_ ; if ($mgconn.scopes -contains $thisPerm) { write-verbose "has scope: $($thisPerm)" } else { $addScopes += @($thisPerm) ; write-verbose "ADD scope: $($thisPerm)" } } ; } ;
            $pltCcMG = [ordered]@{NoWelcome = $true; ErrorAction = 'STOP'; silent = $($silent) } ;
            if ($addScopes) { $pltCcMG.add('RequiredScopes', $addscopes); $pltCcMG.add('ContextScope', 'Process'); $pltCCMG.silent = $false; write-verbose "Adding non-default Scopes, setting non-persistant single-process ContextScope" } ;
            if ($MGConn.isConnected -AND $addScopes -AND $mgconn.CertificateThumbprint) {
                $smsg = "CBA cert lacking scopes :$($addscopes -join ',')!"  ; $smsg += "`nDisconnecting to use interactive connection: connect-mg -RequiredScopoes `"'$($addscopes -join "','")'`"" ; $smsg += "`n(alt: : connect-mggraph -Scopes `"'$($addscopes -join "','")'`" )" ; write-warning $smsg ;
                disconnect-mggraph ;
            } elseif ($MGConn.isConnected -AND $addScopes -and -not ($mgconn.CertificateThumbprint)) {write-verbose "Connect via Account with specified scopes" 
            } elseif (-NOT ($MGConn.isConnected) -AND $addScopes -and -not ($mgconn.CertificateThumbprint)) { $pltCCMG.add('Credential', $credO365TORSID) 
            } elseif (-NOT ($MGConn.isConnected)) { write-verbose "not isMGConnected" ; 
            }else { write-verbose "(currently connected with any specifically specified required Scopes)" ; 
                $pltCcMG = $null ; }
            if ($pltCcMG) {
                $smsg = "connect-mg w`n$(($pltCCMG.getenumerator() | ?{$_.name -notmatch 'requiredscopes'} | ft -a | out-string|out-string).trim())" ; 
                if($pltCCMG.requiredscopes){$smsg += "`n`n-requiredscopes:`n$(($pltCCMG.requiredscopes|out-string).trim())`n"} ; 
                if ($silent) {} else { if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                connect-mg @pltCCMG ;
            } ;
        }elseif(gcm get-mgcontext -ea STOP){
            write-host -foregroundcolor yellow "no verb-mg: using native mg cmdlets" ; 
            $MGCtxt = get-mgcontext ;
            if ($RequiredScopes) { $addScopes = @() ; $RequiredScopes | foreach-object { $thisPerm = $_ ; if (($MGCtxt.scopes|?{$_}) -contains $thisPerm) { write-verbose "has scope: $($thisPerm)" } else { $addScopes += @($thisPerm) ; write-verbose "ADD scope: $($thisPerm)" } } ; } ;
            $pltCcMG = [ordered]@{NoWelcome = $true; ErrorAction = 'STOP'} ;
            if ($addScopes) { $pltCcMG.add('Scopes', $addscopes); $pltCcMG.add('ContextScope', 'Process'); $pltCCMG.silent = $false; write-verbose "Adding non-default Scopes, setting non-persistant single-process ContextScope" } ;
            if ($MGCtxt -AND $addScopes -AND $MGCtxt.CertificateThumbprint) {
                $smsg = "CBA cert lacking scopes :$($addscopes -join ',')!"  ; $smsg += "`nDisconnecting to use interactive connection: connect-mg -RequiredScopoes `"'$($addscopes -join "','")'`"" ; $smsg += "`n(alt: : connect-mggraph -Scopes `"'$($addscopes -join "','")'`" )" ; write-warning $smsg ;
                disconnect-mggraph ;
            } elseif ($MGCtxt -AND $addScopes -and -not ($MGCtxt.CertificateThumbprint)) {write-verbose "(isMGConnected via Account)"
            } elseif (-NOT ($MGCtxt) -AND $addScopes -and -not ($MGCtxt.CertificateThumbprint)) {
            } elseif (-NOT ($MGCtxt)) { 
            }else { write-verbose "(currently connected with any specifically specified required Scopes)" ; 
                $pltCcMG = $null ; }
            if ($pltCcMG) {
                $smsg = "connect-mggraph w`n$(($pltCCMG.getenumerator() | ?{$_.name -notmatch 'scopes'} | ft -a | out-string|out-string).trim())" ; 
                if($pltCCMG.scopes){$smsg += "`n`n-scopes:`n$(($pltCCMG.requiredscopes|out-string).trim())`n"} ; 
                if ($silent) {} else { if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                connect-mggraph @pltCCMG ;
            } ;
        }
        if (-not (get-command Get-MgUser)) {
            $smsg = "Missing Get-MgUser!" ; $smsg += "`nPre-connect to Microsoft.Graph via:" ; $smsg += "`nConnect-MgGraph -Scopes `'$($requiredscopes -join "','")`'" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent }else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            BREAK ;
        } ;
        #endregion MG_CONNECT ; #*------^ END MG_CONNECT ^------
        $isXoPConn = [boolean]( (gcm get-pssession -ea 0) -AND (get-pssession -ea 0 | ? { $_.State -eq 'Opened' -AND $_.Availability -eq 'Available' })); if (-not $isXoPConn) {
            Reconnect-Ex2010 -silent:$($silent)
        } else { write-verbose "XOP connected" };
        $isADConn = [boolean](gcm get-aduser -ea 0) ; if (!$isADConn) { $env:ADPS_LoadDefaultDrive = 0 ; $sName = "ActiveDirectory"; if (!(Get-Module | where { $_.Name -eq $sName })) { Import-Module $sName -ea Stop } }else { write-verbose "ADMS connected" };
        #endregion SVCCONN_LITE ; #*------^ END SVCCONN_LITE ^------
        #region SVCCONN_TEST ; #*------v SVCCONN_TEST v------
        $isXoConn = [boolean]( (gcm Get-ConnectionInformation -ea 0) -AND (Get-ConnectionInformation -ea 0 | ? { $_.State -eq 'Connected' -AND $_.TokenStatus -eq 'Active' })) ; if (-not $isXoConn) { write-host "No EXO connection! RUN: connect-excchangeonline or rxo" ; BREAK }else { write-verbose "EXO connected" };
        $isMgConn = [boolean](get-mgcontext) ; if (-not $isMgConn) { write-host "No MG connection! RUN: Connect-MgGraph or Connect-Mg" ; BREAK }else { write-verbose "MG connected" };
        $isXoPConn = [boolean]( (gcm get-pssession -ea 0) -AND (get-pssession -ea 0 | ? { $_.State -eq 'Opened' -AND $_.Availability -eq 'Available' })); if (-not $isXoPConn) { WRITE-WARNING "NO XOP CONNECTION! run rx10" ; BREAK }else { write-verbose "XOP connected" };
        $isADConn = [boolean](gcm get-aduser -ea 0) ; if (!$isADConn) { TRY { $env:ADPS_LoadDefaultDrive = 0 ; $sName = "ActiveDirectory"; if (!(Get-Module | where { $_.Name -eq $sName })) { Import-Module $sName -ea Stop } } CATCH { $ErrTrapd = $Error[0] ; WRITE-WARNING "$(($ErrTrapd | fl * -Force|out-string).trim())" ; BREAK } }else { write-verbose "ADMS connected" };
        #endregion SVCCONN_TEST ; #*------^ END SVCCONN_TEST ^------

        if (!$domaincontroller) {
            if($env:userdomain -eq 'CMW'){
                $domaincontroller = get-addomaincontroller | select -expand hostname ;
            }else{
                $pltGDC = @{} ;
                if ($DCExclude) { $pltGDC.add('Exclude', $DCExclude) };
                if ($DCServerPrefix) { $pltGDC.add('ServerPrefix', $DCServerPrefix) };
                if ($pltgdc.GetEnumerator().name) { $domaincontroller = get-gcfast @pltgdc } else { $domaincontroller = get-gcfast } ;
            } ;
        } ;
        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):using common `$domaincontroller:$($domaincontroller)" ;
    
        #endregion SUBMAIN ; #*======^ END SUB MAIN ^======
        

    } #  # BEG-E
    PROCESS{

        #region START_LOG_OPTIONS #*======v START_LOG_OPTIONS v======
        $useSLogHOl = $true ; # one or 
        $useTransPath = $false ; # TRANSCRIPTPATH
        $useTransRotate = $false ; # TRANSCRIPTPATHROTATE
        $useStartTrans = $false ; # STARTTRANS
        $useTransNoDep = $false ; # TRANSCRIPT_NODEP
        $useTransNoDepLITE = $false ; # TRANSCRIPT_NODEPLITE
        $useTransBasicScript = $false ; # BASIC_SCRIPT_TRANSCRIPT
        #region START_LOG_HOLISTIC #*------v START_LOG_HOLISTIC v------
        if($useSLogHOl){
            # Single log for script/function example that accomodates detect/redirect from AllUsers scope'd installed code, and hunts a series of drive letters to find an alternate logging dir (defers to profile variables)
            #${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ;
            if(-not (get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
            foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
            if(-not (get-variable rgxPSAllUsersScope -ea 0)){$rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps(((d|m))*)1|dll)$" ;} ;
            if(-not (get-variable rgxPSCurrUserScope -ea 0)){$rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps((d|m)*)1|dll)$" ;} ;
            $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ;} ;
            if($whatif.ispresent){$pltSL.add('whatif',$($whatif))}
            elseif($WhatIfPreference.ispresent ){$pltSL.add('whatif',$WhatIfPreferenc)} ;         
            # if using [CmdletBinding(SupportsShouldProcess)] + -WhatIf:$($WhatIfPreference):
            #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
            #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag="$($ticket)-$($TenOrg)-LASTPASS-" ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
            #$pltSL.Tag = $((@($ticket,$usr) |?{$_}) -join '-')
            #if($ticket){$pltSL.Tag = $ticket} ;
            #$pltSL.Tag = $env:COMPUTERNAME ; 
            $pltSL.Tag = $DisplayName
            #$tagfields = 'ticket','DisplayName'
            #$tagfields = 'ticket','UserPrincipalName','folderscope' ; # DomainName TenOrg ModuleName 
            #$tagfields | foreach-object{$fld = $_ ; if(get-variable $fld -ea 0 |?{$_.value} ){$pltSL.Tag += @($((get-variable $fld).value))} } ; 
            if($pltSL.Tag -is [array]){$pltSL.Tag = $pltSL.Tag -join '-' } ; 
            #$transcript = ".\logs\$($Ticket)-$($DomainName)-$(split-path $rMyInvocation.InvocationName -leaf)-$(get-date -format 'yyyyMMdd-HHmmtt')-trans-log.txt" ; 
            #$pltSL.Tag += "-$($DomainName)"
            <# : optional block, variant, don't use both above $pltSL.Tag, and this (doubles trainling dash)
            if($rPSBoundParameters.keys){ # alt: leverage $rPSBoundParameters hash
                $sTag = @() ; 
                #$pltSL.TAG = $((@($rPSBoundParameters.keys) |?{$_}) -join ','); # join all params
                if($rPSBoundParameters['Summary']){ $sTag+= @('Summary') } ; # build elements conditionally, string
                if($rPSBoundParameters['Number']){ $sTag+= @("Number$($rPSBoundParameters['Number'])") } ; # and keyname,value
                $pltSL.Tag += "-$($sTag -join ',')" ; # 4:46 PM 7/16/2025 flipped to append, not assign
            } ; 
            #>
            if($rvEnv.isScript){
                write-host "`$script:PSCommandPath:$($script:PSCommandPath)" ;
                write-host "`$PSCommandPath:$($PSCommandPath)" ;
                if($rvEnv.PSCommandPathproxy){ $prxPath = $rvEnv.PSCommandPathproxy }
                elseif($script:PSCommandPath){$prxPath = $script:PSCommandPath}
                elseif($rPSCommandPath){$prxPath = $rPSCommandPath} ; 
            } ; 
            if($rvEnv.isFunc){
                if($rvEnv.FuncDir -AND $rvEnv.FuncName){
                       $prxPath = join-path -path $rvEnv.FuncDir -ChildPath $rvEnv.FuncName ; 
                } else {
                    write-warning "Missing either `$rvEnv.FuncDir -OR `$rvEnv.FuncName!" ; 
                } ; 
            } ; 
            if(-not $rvEnv.isFunc){
                # under funcs, this is the scriptblock of the func, not a path
                if($rvEnv.MyInvocationproxy.MyCommand.Definition){$prxPath2 = $rvEnv.MyInvocationproxy.MyCommand.Definition }
                elseif($rvEnv.MyInvocationproxy.MyCommand.Definition){$prxPath2 = $rvEnv.MyInvocationproxy.MyCommand.Definition } ; 
            } ; 
            if($prxPath){
                # 12/12/2025 new code to patch no-ext $prxPath
                if(-not [System.IO.Path]::GetExtension($prxPath)){
                    write-verbose "no-extension `$prxpath, asserting fake ext (.ps1|.psm1 as approp)" ;                         
                    switch($rvEnv.runSource){
                        'Function'{$prxPath = "$($prxPath).psm1" }
                        'ExternalScript'{$prxPath = "$($prxPath).ps1" }
                        default {
                            $smsg = "NO RECOGNIZED `$rvEnv.runSource: '$($rvEnv.runSource)'`nUNABLE TO SAFELY TEST FOR AllUsers or CU SCOPE!: ABORTING (Could log into module hosting dir!)" ; 
                            write-warning $smsg ; throw $smsg ; 
                            BREAK ; 
                        }
                    } ; 
                } ; 
                if(($prxPath -match $rgxPSAllUsersScope) -OR ($prxPath -match $rgxPSCurrUserScope)){
                    $bDivertLog = $true ; 
                    switch -regex ($prxPath){
                        $rgxPSAllUsersScope{$smsg = "AllUsers"} 
                        $rgxPSCurrUserScope{$smsg = "CurrentUser"}
                    } ;
                    $smsg += " context script/module, divert logging into [$budrv]:\scripts" 
                    write-verbose $smsg  ;
                    if($bDivertLog){
                        if((split-path $prxPath -leaf) -ne $rvEnv.CmdletName){
                            # function in a module/script installed to allusers|cu - defer name to Cmdlet/Function name
                            $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($rvEnv.CmdletName).ps1") ;
                        } else {
                            # installed allusers|CU script, use the hosting script name
                            $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $prxPath -leaf)) ;
                        }
                    } ;
                } else {
                    $pltSL.Path = $prxPath ;
                } ;
            }elseif($prxPath2){
                # 12/12/2025 new code to patch no-ext $prxPath2
                if(-not [System.IO.Path]::GetExtension($prxPath2)){
                    write-verbose "no-extension `$prxPath2, asserting fake ext (.ps1|.psm1 as approp)" ;                         
                    switch($rvEnv.runSource){
                        'Function'{$prxPath2 = "$($prxPath2).psm1" }
                        'ExternalScript'{$prxPath2 = "$($prxPath2).ps1" }
                        default {
                            $smsg = "NO RECOGNIZED `$rvEnv.runSource: '$($rvEnv.runSource)'`nUNABLE TO SAFELY TEST FOR AllUsers or CU SCOPE!: ABORTING (Could log into module hosting dir!)" ; 
                            write-warning $smsg ; throw $smsg ; 
                            BREAK ; 
                        }
                    } ; 
                } ; 
                if(($prxPath2 -match $rgxPSAllUsersScope) -OR ($prxPath2 -match $rgxPSCurrUserScope) ){
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $prxPath2 -leaf)) ;
                } elseif(test-path $prxPath2) {
                    $pltSL.Path = $prxPath2 ;
                } elseif($rvEnv.CmdletName){
                    $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($rvEnv.CmdletName).ps1") ;
                } else {
                    $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$rvEnv.CmdletName, FROM WHICH TO BUILD A START-LOG.PATH!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    BREAK ;
                } ; 
            } else{
                $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$rvEnv.CmdletName, FROM WHICH TO BUILD A START-LOG.PATH!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                BREAK ;
            }  ;
            write-verbose "start-Log w`n$(($pltSL|out-string).trim())" ; 
            $logspec = start-Log @pltSL ;
            $error.clear() ;
            TRY {
                if($logspec){
                    $logging=$logspec.logging ;
                    $logfile=$logspec.logfile ;
                    $transcript=$logspec.transcript ;
                    $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                    if($stopResults){
                        $smsg = "Stop-transcript:$($stopResults)" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    } ; 
                    $startResults = start-Transcript -path $transcript -whatif:$false -confirm:$false;
                    if($startResults){
                        $smsg = "start-transcript:$($startResults)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 
                } else {throw "Unable to configure logging!" } ;
            } CATCH [System.Management.Automation.PSNotSupportedException]{
                if($host.name -eq 'Windows PowerShell ISE Host'){
                    $smsg = "This version of $($host.name):$($host.version) does *not* support native (start-)transcription" ; 
                } else { 
                    $smsg = "This host does *not* support native (start-)transcription" ; 
                } ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                #region SendMailAlert ; #*------v SendMailAlert v------
                $SmtpBody += "`n===FAIL Summary:" ;
                $SmtpBody += "`n$('-'*50)" ;
                $SmtpBody += "`n$('-'*50)" ;
                $smsg += "`n$(($smsg |out-string).trim())" ; 
                $sdEmail = @{
                    smtpFrom = $SMTPFrom ;
                    SMTPTo = $SMTPTo ;
                    SMTPSubj = $SMTPSubj ;
                    #SMTPServer = $SMTPServer ;
                    SmtpBody = $SmtpBody ;
                    SmtpAttachment = $SmtpAttachment ;
                    BodyAsHtml = $false ; # let the htmltag rgx in Send-EmailNotif flip on as needed
                    verbose = $($VerbosePreference -eq "Continue") ;
                } ;
                $smsg = "Send-EmailNotif w`n$(($sdEmail|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Send-EmailNotif @sdEmail ;

                #endregion SendMailAlert ; #*------^ END SendMailAlert ^------
            } ;
        } ; 
        #endregion START_LOG_HOLISTIC #*------^ END START_LOG_HOLISTIC ^------
        #...
        #endregion START_LOG_OPTIONS #*======^ START_LOG_OPTIONS ^======

        #region SPLATDEFS ; # ------
        # dummy hashes
        if($host.version.major -ge 3){
            $InputSplat=[ordered]@{
                Dummy = $null ;
                Verbose = $($VerbosePreference -eq 'Continue') ; 
            } ;
            $MbxSplat=[ordered]@{
                Dummy = $null ;
                Verbose = $($VerbosePreference -eq 'Continue') ; 
            } ;
            $MbxSetSplat=[ordered]@{
                Dummy = $null ;
                Verbose = $($VerbosePreference -eq 'Continue') ; 
            } ;
            $MbxSetCASmbx=[ordered]@{
                Dummy = $null ;
                Verbose = $($VerbosePreference -eq 'Continue') ; 
            } ;
            $ADSplat=[ordered]@{
                Dummy = $null ;
                Verbose = $($VerbosePreference -eq 'Continue') ; 
            } ;
            $UsrSplat=[ordered]@{
                Dummy = $null ;
                Verbose = $($VerbosePreference -eq 'Continue') ; 
            } ;
        } else {
            $InputSplat=@{
                Dummy = $null ;
                Verbose = $($VerbosePreference -eq 'Continue') ; 
            } ;
            $MbxSplat=@{
                Dummy = $null ;
                Verbose = $($VerbosePreference -eq 'Continue') ; 
            } ;
            $MbxSetSplat=@{
                Dummy = $null ;
                Verbose = $($VerbosePreference -eq 'Continue') ; 
            } ;
            $MbxSetCASmbx=@{
                Dummy = $null ;
                Verbose = $($VerbosePreference -eq 'Continue') ; 
            } ;
            $ADSplat=@{
                Dummy = $null ;
                Verbose = $($VerbosePreference -eq 'Continue') ; 
            } ;
            $UsrSplat=@{
                Dummy = $null ;
                Verbose = $($VerbosePreference -eq 'Continue') ; 
            } ;
        } ;
        # then immediately remove the dummy value:
        $InputSplat.remove("Dummy") ;
        $MbxSplat.remove("Dummy") ;
        $MbxSetSplat.remove("Dummy") ;
        $ADSplat.remove("Dummy") ;
        $UsrSplat.remove("Dummy") ;
        # 8:08 AM 10/6/2017 add missing CASMailbox splat
        $MbxSetCASmbx.remove("Dummy") ;

        # also, less code post-decl to populate the $hash with fields, post creation:
        #$InputSplat.Add("NewField",$($NewValue)) ;
        $InputSplat.Add("Ticket",$($null)) ;
        $InputSplat.Add("DisplayName",$("")) ;
        $InputSplat.Add("MInitial",$("")) ;
        $InputSplat.Add("Owner",$("")) ;
        $InputSplat.Add("OwnerMbx",$($null)) ;
        $InputSplat.Add("BaseUser",$("")) ;
        $InputSplat.Add("IsContractor",$($false)) ;
        $InputSplat.Add("NonGeneric",$($false)) ;
        $InputSplat.Add("Vscan",$($null)) ;
        $InputSplat.Add("BUserAD",$($null)) ;
        $InputSplat.Add("ADDesc",$($null)) ;
        $InputSplat.Add("Domain",$($null)) ;
        $InputSplat.Add("DomainController",$($null)) ;
        $InputSplat.Add("SiteName",$($null)) ;
        $InputSplat.Add("OrganizationalUnit",$($null)) ;

        #$MbxSplat.Add("OrganizationalUnit",$($null)) ;
        $MbxSplat.Add("Shared",$($null)) ;
        $MbxSplat.Add("Name",$($null)) ;
        $MbxSplat.Add("DisplayName",$($null)) ;
        $MbxSplat.Add("userprincipalname",$($null)) ;
        $MbxSplat.Add("OrganizationalUnit", $($null)) ;
        #$MbxSplat.Add("Office", $($null)) ;
        # new-mailbox syntax set that includes -shared DOESN'T include -office!, move it to MbxSetSplat
        $MbxSplat.Add("database",$($null)) ;
        $MbxSplat.Add("password",$($null)) ;
        $MbxSplat.Add("FirstName",$($null)) ;
        $MbxSplat.Add("Initials",$($null));
        $MbxSplat.Add("LastName",$($null)) ;
        $MbxSplat.Add("samaccountname",$($null)) ;
        $MbxSplat.Add("alias",$($null)) ;
        $MbxSplat.Add("ResetPasswordOnNextLogon",$($false));
        $MbxSplat.Add("RetentionPolicy",$($TORMeta['RetentionPolicy'])) ;
        $MbxSplat.Add("ActiveSyncMailboxPolicy",$('Default')) ;
        $MbxSplat.Add("domaincontroller",$($null)) ;
        $MbxSplat.Add("whatif",$true) ;

        $MbxSetSplat.Add("identity",$($null)) ;
        $MbxSetSplat.Add("CustomAttribute9",$($null)) ;
        $MbxSetSplat.Add("CustomAttribute5",$($null)) ;
        $MbxSetSplat.Add("Office", $($null)) ; # 8:44 AM 11/26/2019 shifted unsupported syntax mix from new-mailbox to set-mailbox
        $MbxSetSplat.Add("domaincontroller",$($null)) ;
        $MbxSetSplat.Add("whatif",$true) ;

        # CASMailbox splat
        $MbxSetCASmbx.Add("identity",$($null)) ;
        $MbxSetCASmbx.Add("ActiveSyncMailboxPolicy",$($null)) ;
        $MbxSetCASmbx.Add("domaincontroller",$($null)) ;
        $MbxSetCASmbx.Add("whatif",$true) ;

        $ADSplat.Add("manager",$($null)) ;
        $ADSplat.Add("Description",$($null)) ;
        $ADSplat.Add("Server",$($null)) ;
        $ADSplat.Add("identity",$($null)) ;
        $ADSplat.Add("whatif",$($true)) ;

        $UsrSplat.Add("whatif",$($true)) ;
        $UsrSplat.Add("City",$($null)) ;
        $UsrSplat.Add("CountryOrRegion",$($null)) ;
        $UsrSplat.Add("Fax",$($null)) ;
        $UsrSplat.Add("Office",$($null)) ;
        $UsrSplat.Add("PostalCode",$($null)) ;
        $UsrSplat.Add("StateOrProvince",$($null)) ;
        $UsrSplat.Add("StreetAddress",$($null)) ;
        $UsrSplat.Add("Title",$($null)) ;
        $UsrSplat.Add("Phone",$($null)) ;
        $UsrSplat.Add("domaincontroller",$($null)) ;

        # start stocking in params into $InputSplat
        if($DisplayName){
            $InputSplat.DisplayName=$DisplayName
        };
        if($MInitial){$InputSplat.MInitial=$MInitial};
        if($Owner){$InputSplat.Owner=$Owner};

        if($BaseUser){$InputSplat.BaseUser=$BaseUser};
        # only reset from defaults on explicit -NonGeneric $true param
        if($NonGeneric -eq $true){

        } else {
            # force it if not true
            $NonGeneric=$false;
        } ;
        $InputSplat.NonGeneric=$NonGeneric
        if($IsContractor){$InputSplat.IsContractor=$IsContractor};

        if ($Vscan){
            if ($Vscan -match "(?i:^(YES|NO)$)" ) {
                $Vscan = $Vscan.ToString().ToUpper() ;
                $InputSplat.Vscan=$Vscan ;
            } else {
                $Vscan = $null ;
                $InputSplat.Vscan="YES";
            }  ;
        };

        # 3:07 PM 10/4/2017 Cu5 override support (normally inherits from assigned owner/manager)
        # 3:35 PM 4/15/2026 the below is blanking cu5 because dw isn't in the $rgxcu5 spec! dumpm it!
        if ($Cu5){
            #if ($Cu5 -match $rgxCU5 ) {
                # pulled switch out, it wasn't actually translating, just rgx of the final tags
                $InputSplat.Cu5=$Cu5;
            #} else {
            #    $InputSplat.Cu5=$null ;
            #}  ;
        }; #  # if-E Cu5


        if($SiteOverride){
            if($InputSplat.keys -contains 'SiteOverride'){
                $InputSplat.SiteOverride=$SiteOverride ; 
            }else{
                $InputSplat.add('SiteOverride',$SiteOverride) ;  
            }
        };
        # 2:59 PM 1/24/2025 new OfficeOverride
        if($OfficeOverride){
            if($InputSplat.keys -contains 'OfficeOverride'){
                $InputSplat.OfficeOverride=$OfficeOverride ; 
            }else{
                $InputSplat.add('OfficeOverride',$OfficeOverride) ;  
            }
        };
        if($Ticket){$InputSplat.Ticket=$Ticket};


        #endregion SPLATDEFS ; # ------
        
        # load ADMS
        #$reqMods+="load-ADMS;get-AdminInitials".split(";") ;
        #if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
        write-host -foregroundcolor darkgray "$((get-date).ToString('HH:mm:ss')):(loading ADMS...)" ;
        load-ADMS -cmdlet get-aduser,Set-ADUser,Get-ADGroupMember,Get-ADDomainController,Get-ADObject,get-adforest | out-null ;

        $AdminInits=get-AdminInitials ;

        #region LOADMODS ; # ------

        #region DATAPREP ; # ------
        if ( ($InputSplat.OwnerMbx=(get-mailbox -identity $($InputSplat.Owner) -ea 0)) -OR ($InputSplat.OwnerMbx=(get-remotemailbox -identity $($InputSplat.Owner) -ea 0)) ){
            $rgxOUMigrations = ',OU=_MIGRATIONS,DC=global,DC=ad,DC=toro,DC=com$' ;
            $rgxMigationsSite = ',OU=(\w+),OU=_MIGRATIONS,DC=global,DC=ad,DC=toro,DC=com$' ;            
            if($InputSplat.OwnerMbx.DistinguishedName -match $rgxOUMigrations){
                $OSiteCode = [regex]::Match($InputSplat.OwnerMbx.DistinguishedName ,$rgxMigationsSite).groups[1].value ;
                if($OSiteCode){
                    $smsg = "Resolved _MIGRATIONS tree OSiteCode:$($OSiteCode)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                }else{
                    throw "Unable to resolve $($pltNmbx.Owner.DistinguishedName) into a SiteCode" ;
                    Break ;
                }
            }else{
                $OSiteCode = $InputSplat.OwnerMbx.identity.tostring().split('/')[1]
            } ;
            $InputSplat.SiteCode=$OSiteCode;       
        } else {
          throw "Unable to resolve $($InputSplat.Owner) to any existing OP or EXO mailbox" ;
          Cleanup ; BREAK ;
        } ;

        $InputSplat.Domain=$($InputSplat.OwnerMbx.identity.tostring().split("/")[0]) ;
        #$InputSplat.SiteCode=($InputSplat.OwnerMbx.identity.tostring().split('/')[1]) ;

        $domain=$InputSplat.Domain ;
        if(!$domaincontroller){
            $domaincontroller =(get-gcfast -domain $domain) ;
        } else {
            $smsg= "Using hard-coded domaincontroller:$($domaincontroller)" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
        } ;
        $InputSplat.DomainController=$domaincontroller ;

        $MbxSplat.DomainController=$($InputSplat.DomainController) ;
        $MbxSetSplat.DomainController=$($InputSplat.DomainController) ;
        $ADSplat.Server=$($InputSplat.DomainController) ;
        $UsrSplat.DomainController=$($InputSplat.DomainController) ;

        if($InputSplat.SiteOverride){
            $SiteCode=$($InputSplat.SiteOverride);
            # force the $InputSplat.SiteCode to match the override too
            $InputSplat.SiteCode=$($InputSplat.SiteOverride);
        } else {
            $SiteCode=$InputSplat.SiteCode.tostring();
        } ;

        # if we need the OwnerDn 
        if ( ($OwnerMbx = (get-mailbox -identity $($pltNmbx.Owner) -ea 0)) -OR ($OwnerMbx = (get-remotemailbox -identity $($pltNmbx.Owner) -ea 0)) ) {
            if ($ownermbx.DistinguishedName -match $rgxOUMigrations) {
                $OSiteCode = [regex]::Match($ownermbx.DistinguishedName, $rgxMigationsSite).groups[1].value ;
                if ($OSiteCode) {
                    $smsg = "Resolved _MIGRATIONS tree OSiteCode:$($OSiteCode)" ;
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
                } else {
                    throw "Unable to resolve $($pltNmbx.Owner.DistinguishedName) into a SiteCode" ;
                    Break ;
                }
            } else {
                $OSiteCode = $Ownermbx.identity.tostring().split('/')[1]
            } ;
        } else {
            throw "Unable to resolve $($pltNmbx.Owner) to any existing OP or EXO mailbox" ;
            Break ;
        } ;

        $pltGSmbx = [ordered]@{
            Sitecode = $SiteCode ;
            #Generic = $null ;
            #Resource = $null; 
            #modelDistinguishedName = $null ;
        }
        if ($ownermbx.DistinguishedName -match $rgxOUMigrations) {
            $pltGSmbx.add('modelDistinguishedName',$ownermbx.DistinguishedName)
            #$pltGSmbx.modelDistinguishedName = $ownermbx.DistinguishedName ; 
        }else{
            $pltGSmbx.add('modelDistinguishedName',$ownermbx.DistinguishedName)
            #$pltGSmbx.modelDistinguishedName = $ownermbx.DistinguishedName ; 
        }
        
        If($InputSplat.NonGeneric) {
            if($pltGSmbx.keys -contains 'generic'){$pltGSmbx.remove('Generic')}
            #$pltGSmbx.Generic = $null ; 
            #$pltGSmbx.Resource = $null ; 
        } elseIf($Room -OR $Equipement) {
            #$pltGSmbx.Resource = $true ; 
            $pltGSmbx.add('Resource',$true) ; 
            #$pltGSmbx.Generic = $null ;
        } else {
            #$pltGSmbx.Generic = $true 
            $pltGSmbx.add('Generic',$true ) ; 
            #$pltGSmbx.Resource = $null ; 
        } ; 
        $tCmdlet = 'get-SiteMbxOU' ; $BMod = 'verb-ADMS' ; 
        if($psISE -AND ((get-date ).tostring() -match $dbgDate)){ 
            if((gcm $tCmdlet).source -eq $BMod){
                Do{
                    gci "D:\scripts\$($tCmdlet)_func.ps1" -ea STOP | ipmo -fo -verb  ;
                }until((gcm $tCmdlet).source -ne $BMod)
            } ;
        } ; 
        $smsg = "Get-SiteMbxOU w`n$(($pltGSmbx|out-string).trim())" ; 
        write-verbose $smsg ; 
        if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU @pltGSmbx)   ) {
        } else { Cleanup ; BREAK ;}

        <#xxx
        If($InputSplat.NonGeneric) {
            if ($ownermbx.DistinguishedName -match $rgxOUMigrations) {
                
                #if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU  -Sitecode $SiteCode -Generic $false -modelDistinguishedName $ownermbx.DistinguishedName)   ) {
                $pltGSmbx.Generic = $false ; 
                $pltGSmbx.modelDistinguishedName = $ownermbx.DistinguishedName ; 
                #if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU  -Sitecode $SiteCode -Generic $false -modelDistinguishedName $ownermbx.DistinguishedName)   ) {
                if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU @PltGSmbx)   ) {
                } else { Cleanup ; BREAK ;}
            } else {
                #$OSiteCode = $Ownermbx.identity.tostring().split('/')[1]
                #if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU  -Sitecode $SiteCode -Generic $false)   ) {
                #if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU  -Sitecode $SiteCode -Generic $false  -modelDistinguishedName $ownermbx.DistinguishedName)   ) {
                $pltGSmbx.Generic = $false ; 
                $pltGSmbx.modelDistinguishedName = $ownermbx.DistinguishedName ; 
                if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU @PltGSmbx)   ) {
                } else { Cleanup ; BREAK ;}
            } ;
            
        } elseIf($Room -OR $Equipement) {
            if ($ownermbx.DistinguishedName -match $rgxOUMigrations) {
                if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU  -Sitecode $SiteCode -Resource $true  -modelDistinguishedName $ownermbx.DistinguishedName) ) {
            }else{
                if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU  -Sitecode $SiteCode -Resource $true ) ) {
                } else { Cleanup ; BREAK ;}
            }
        } else {
            if ($ownermbx.DistinguishedName -match $rgxOUMigrations ) {
                if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU  -Sitecode $SiteCode -Generic $true -modelDistinguishedName $ownermbx.DistinguishedName) ) {
                } else { Cleanup ; BREAK ;}
            }else{
                if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU  -Sitecode $SiteCode -Generic $true ) ) {
                } else { Cleanup ; BREAK ;}
            }
        }
        #> # XXX

        # add forced office designation, to match $SiteCode/OU
        # New-Mailbox doesn't support both -Shared & -Office in the same syntax set, move it to $MbxSetSplat
        
        # 2:53 PM 1/24/2025 exempt OfficeOverride
        if($InputSplat.OfficeOverride){
            $smsg = "-OfficeOverride:$($InputSplat.OfficeOverride): overriding user object Office from SiteCode to specified Override!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            $MbxSetSplat.Office = $InputSplat.OfficeOverride ;
        } else { 
            $MbxSetSplat.Office = $SiteCode ;
        } ; 
        $smsg= "Site Located:`$SiteCode:$SiteCode`n`$OrganizationalUnit:$($MbxSplat.OrganizationalUnit)" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn


        if(Get-ADObject $($MbxSplat.OrganizationalUnit) -server $($InputSplat.Domain)){
            $InputSplat.SiteName=$($SiteCode) ;
            $InputSplat.OrganizationalUnit=$($MbxSplat.OrganizationalUnit) ;
            $smsg= "Target Dname: $($InputSplat.DisplayName)" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
            $InputSplat.Add("samaccountname",$null) ;
            If($InputSplat.NonGeneric) {
                $InputSplat.Add("Shared",$False ) ;
                # user-style acct, fname & lname
                if ($InputSplat.DisplayName.tostring().indexof(" ") -gt 0){
                    $InputSplat.Add("FirstName",$($InputSplat.DisplayName.tostring().split(" ")[0].trim()) ) ;
                    $InputSplat.Add("LastName",$($InputSplat.DisplayName.tostring().split(" ")[1].trim()) ) ;
                } else {
                    $InputSplat.Add("FirstName",$null ) ;
                    $InputSplat.Add("LastName",$InputSplat.DisplayName) ;
                } ;
                if (($InputSplat.IsContractor) -OR ($InputSplat.SiteName -eq "XIA") ){
                    $LnameClean=Remove-StringDiacritic -string $InputSplat.LastName ;
                    $LnameClean= Remove-StringLatinCharacters -string $LnameClean ;
                    $FnameClean=Remove-StringDiacritic -string $InputSplat.FirstName ;
                    $FnameClean= Remove-StringLatinCharacters -string $FnameClean ;
                    $InputSplat.samaccountname=$( ([System.Text.RegularExpressions.Regex]::Replace($LnameClean,"[^1-9a-zA-Z_]","").tostring().substring(0,[math]::min([System.Text.RegularExpressions.Regex]::Replace($LnameClean,"[^1-9a-zA-Z_]","").tostring().length,5)) + $FnameClean.tostring().substring(0,1) + "x").toLower() )  ;
                } else {
                    if($InputSplat.FirstName){
                        $LnameClean=Remove-StringDiacritic -string $InputSplat.LastName ;
                        $LnameClean= Remove-StringLatinCharacters -string $LnameClean ;
                        $FnameClean=Remove-StringDiacritic -string $InputSplat.FirstName ;
                        $FnameClean= Remove-StringLatinCharacters -string $FnameClean ;
                        $InputSplat.samaccountname=$( ([System.Text.RegularExpressions.Regex]::Replace($LnameClean,"[^1-9a-zA-Z_]","").tostring().substring(0,[math]::min([System.Text.RegularExpressions.Regex]::Replace($LnameClean,"[^1-9a-zA-Z_]","").tostring().length,5)) + $FnameClean.tostring().substring(0,1)).toLower() )  ;
                    } else {
                        $LnameClean=Remove-StringDiacritic -string $InputSplat.LastName ;
                        $LnameClean= Remove-StringLatinCharacters -string $LnameClean ;
                        $InputSplat.samaccountname=$( ([System.Text.RegularExpressions.Regex]::Replace($LnameClean,"[^1-9a-zA-Z_]","").tostring().substring(0,[math]::min([System.Text.RegularExpressions.Regex]::Replace($LnameClean,"[^1-9a-zA-Z_]","").tostring().length,20))).toLower() )  ;
                    }
                    # append non-blank MI
                    if($InputSplat.MInitial){
                        $InputSplat.samaccountname+=$($InputSplat.MInitial).tostring().toLower() ;
                    } # if-E
                } ;

                # need to accommodate EXO-hosted MailUser owners
                switch ((get-recipient -Identity $($InputSplat.Owner)).RecipientType ){
                    "UserMailbox" {
                        if ( ($Tmbx=(get-mailbox -identity $($InputSplat.Owner) -ea stop)) ){
                            if($showdebug){ $smsg= "Owner UserMailbox detected" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;  } ;
                            # base users off of Owner
                            if ( $InputSplat.BaseUser=(get-mailbox -identity $($InputSplat.Owner) -domaincontroller $($InputSplat.domaincontroller) -ea continue ) ) {
                            } else { Cleanup ; BREAK ;}
                        } else {
                            throw "Unable to resolve $($InputSplat.ManagedBy) to any existing OP or EXO mailbox" ;
                            Cleanup ; BREAK ;
                        } ;
                    }
                    "MailUser" {
                        if ( ($Tmbx=(get-remotemailbox -identity $($InputSplat.Owner) -ea stop)) ){
                            if($showdebug){ $smsg= "Owner MailUser detected" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;  } ;
                            # base users off of Owner
                            if ( $InputSplat.BaseUser=(get-Remotemailbox -identity $($InputSplat.Owner) -domaincontroller $($InputSplat.domaincontroller) -ea continue ) ) {
                            } else { Cleanup ; BREAK ;}
                        } else {
                            # without the -ea stop, we need an explicit error
                            throw "Unable to resolve $($InputSplat.ManagedBy) to any existing OP or EXO mailbox" ;
                            Cleanup ; BREAK ;
                        } ;
                    }
                    default {
                        throw "$($InputSplat.ManagedBy) Not found, or unrecognized RecipientType" ;
                        Cleanup ; BREAK ;
                    }
                } ;

            } else {
                # strict shared acct, no FirstName - revising, we'll support period in name and auto-split fname.lname, to create requested fname.lname@domain.com addresses wo post modification
                # support for shared/room/equip
                if(!$Equip -AND !$Room){
                    # only use Shared when not Equip or Room
                    $smsg= "SHARED Mailbox specified" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                    $InputSplat.Add("Shared",$true ) ;
                    # tear out the unused
                    $InputSplat.Remove("Room") ;
                    $InputSplat.Remove("Equip") ;
                    # add support divide on period
                    if ($InputSplat.DisplayName.tostring().indexof(".") -gt 0){
                        $InputSplat.Add("FirstName",$($InputSplat.DisplayName.tostring().split(".")[0].trim()) ) ;
                        $InputSplat.Add("LastName",$($InputSplat.DisplayName.tostring().split(".")[1].trim()) ) ;
                        # 2:07 PM 4/24/2026 pull the period out, it should never turn up in final product other than email addr
                        $InputSplat.DisplayName = $($InputSplat.DisplayName.tostring().replace("."," "))
                    } else { 
                        $InputSplat.Add("FirstName",$null ) ;
                        $InputSplat.Add("LastName",$InputSplat.DisplayName) ;
                    } ; 
                } else {
                    $InputSplat.Remove("Shared") ;
                    if($Room -AND !$Equip){
                        $smsg= "ROOM Mailbox specified" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                        $InputSplat.Add("Room",$true) ;
                        $InputSplat.Remove("Equip") ;
                    }elseif($Equip -AND !$Room){
                        $smsg= "EQUIPMENT Mailbox specified" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                        $InputSplat.Add("Equip",$true) ;
                        $InputSplat.Remove("Room") ;
                    } else { throw "INVALID OPTIONS: USE -Room OR -Equip BUT NOT BOTH" }
                    $InputSplat.Add("FirstName",$null ) ;
                    $InputSplat.Add("LastName",$InputSplat.DisplayName) ;
                } ;
                
                #region SAMACCTNAMECLEAN ; #*------v SAMACCTNAMECLEAN v------
                # psv2 LACKS the foreign-lang cleanup funtions below, skip on psv2
                if($host.version.major -lt 3){
                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):Psv2 detected, skipping foreign character normalization!" ;
                     $DnameClean=$InputSplat.DisplayName
                } else {
                    $DnameClean=Remove-StringDiacritic -string $InputSplat.DisplayName
                    $DnameClean= Remove-StringLatinCharacters -string $DnameClean ;
                } ;
                # strip all nonalphnums from samacct!
                #$InputSplat.samaccountname=$([System.Text.RegularExpressions.Regex]::Replace($DnameClean,"[^1-9a-zA-Z_]",""));
                #if($InputSplat.samaccountname.length -gt 20) { $InputSplat.samaccountname=$InputSplat.samaccountname.tostring().substring(0,20) };
                #$InputSplat.samaccountname =  -join (($DnameClean -replace '[^a-zA-Z0-9]', '').ToCharArray()[0..19]) ;
                # no do the strict strip on banned characters, plus spaces
                #$rgxBanSamA = '[\"\/\\\[\]\:\;\|\=\,\+\*\?\<\>\s]' ; 
                $InputSplat.samaccountname = -join (($DnameClean -replace $rgxBanSamA, '').ToCharArray()[0..19]) ;
                try {$conflict = $null ; $conflict = get-aduser $InputSplat.samaccountname -server $InputSplat.DomainController ;} CATCH {} ;
                $incr = 1 ;  
                Do{
                    if($conflict){
                        $incr++ ; 
                        $InputSplat.samaccountname = "$(-join (($DnameClean -replace '[^a-zA-Z0-9]', '').ToCharArray()[0..18]))$($incr)" ; 
                    }
                    try {$conflict = $null ; $conflict = get-aduser $InputSplat.samaccountname -server $InputSplat.DomainController ;} CATCH {} ;
                }while($conflict) ; 
                # THE ALIAS IS DONE SEPARATELY - MORE RESTRICTIVELY DOWN IN THE MBXSPLAT!
                #endregion SAMACCTNAMECLEAN ; #*------^ END SAMACCTNAMECLEAN ^------

                # base generics off of baseuser
                # deter BaseUser as a random user in the $($MbxSplat.OrganizationalUnit)
                # leave the -BaseUser param in to force an override, but if blank, draw random from the OU
                if(!($InputSplat.BaseUser)){
                    # draw a random from the $($MbxSplat.OrganizationalUnit)
                    if($InputSplat.SiteCode -eq "($ADSiteCodeUS)"){
                        if ( $InputSplat.BaseUser=get-mailbox -OrganizationalUnit $($MbxSplat.OrganizationalUnit) -resultsize 50 | Where-Object{($_.distinguishedname -notlike '*demo*') -AND (!$_.CustomAttribute5)} | get-random ) {

                        } elseif ( $InputSplat.BaseUser=get-remotemailbox -OnPremisesOrganizationalUnit  $($MbxSplat.OrganizationalUnit) -resultsize 50 | Where-Object{($_.distinguishedname -notlike '*demo*') -AND (!$_.CustomAttribute5)} | get-random ) {

                        } else {
                            $smsg= "UNABLE TO FIND A BASEUSER - USE -BASEUSER TO SPECIFY A SUITABLE ACCT *SOMEWHERE*" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                            Cleanup ; BREAK ;
                        } ;
                    } else {
                        if ( $InputSplat.BaseUser = get-remotemailbox -OnPremisesOrganizationalUnit $($MbxSplat.OrganizationalUnit) -resultsize 50 | Where-Object { $_.distinguishedname -notlike '*demo*' } | get-random   ) {

                        }elseif ( $InputSplat.BaseUser=get-mailbox -OrganizationalUnit $($MbxSplat.OrganizationalUnit) -resultsize 50 | Where-Object{$_.distinguishedname -notlike '*demo*'} | get-random
                          ) {
                        } elseif ( $InputSplat.BaseUser=get-remotemailbox -OnPremisesOrganizationalUnit  $FallBackBaseUserOU -resultsize 50 | Where-Object{($_.distinguishedname -notlike '*demo*') -AND (!$_.CustomAttribute5)} | get-random ) {
                            # if all else fails, pull a random remotemailbox from ($ADSiteCodeUS) - we're losing comps moving ahead
                        } else {
                            $smsg= "UNABLE TO FIND A BASEUSER - USE -BASEUSER TO SPECIFY A SUITABLE ACCT *SOMEWHERE*" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                            Cleanup ; BREAK ;
                        } ;
                    }
                    write-host -foregroundcolor darkgray "$((get-date).ToString("HH:mm:ss")):Drew Random BaseUser: $($InputSplat.BaseUser.DisplayName) ($($inputsplat.BaseUser.samaccountname))" ;
                } else {
                    switch ((get-recipient -Identity $($InputSplat.BaseUser)).RecipientType ){
                        "UserMailbox" {
                            if ( ($tmpBaseUser=(get-mailbox -identity $InputSplat.BaseUser -ea stop)) ){
                                    $InputSplat.BaseUser=$tmpBaseUser ;
                            } else {
                                throw "Unable to resolve $($InputSplat.BaseUser) to any existing OP or EXO mailbox" ;
                                Cleanup ; BREAK ;
                            } ;
                        } ;
                        "MailUser" {
                            if ( ($tmpBaseUser=(get-remotemailbox -identity $InputSplat.BaseUser -ea stop)) ){
                                $InputSplat.BaseUser=$tmpBaseUser ;
                            } else {
                                # without the -ea stop, we need an explicit error
                                throw "Unable to resolve $($InputSplat.BaseUser) to any existing OP or EXO mailbox" ;
                                Cleanup ; BREAK ;
                            } ;
                        } ;
                        default {
                            throw "$($InputSplat.ManagedBy) Not found, or unrecognized RecipientType" ;
                            Cleanup ; BREAK ;
                        }
                    } ;
                } ;

            };
            $InputSplat.Add("Phone","") ;
            # *** BREAKPOINT ;
            $InputSplat.Add("Title","") ;

             #region passwordgen #-----------
            # need to test complex, and if failed, pull another: (above doesn't consistently deliver Ad complexity req's)
            # 2:16 PM 8/2/2023 revised pol, don't need complex, but will pass with it, but leng is bumped; until rebuild, can push up default with explicit -minLen param
            # # method: GeneratePassword(int length, int numberOfNonAlphanumericCharacters)
            Do { $password = $([System.Web.Security.Membership]::GeneratePassword(14,2)) } Until (Validate-Password -pwd $password -minLen 14) ;
            $InputSplat.Add("pass",$($password));
            #region passwordgen #-----------

            # 9:43 AM 4/1/2016 secgrp is only necc for Generics
            if(!$InputSplat.NonGeneric){
                $InputSplat.Add("PermGrpName",$($InputSplat.SiteName + "-Data-Email-" + $InputSplat.DisplayName + "-G")) ;
            } else {
                $InputSplat.Add("PermGrpName",$null) ;
            } ;

            if ( $InputSplat.BUserAD=(get-user -identity $($InputSplat.BaseUser.samaccountname) -domaincontroller $($InputSplat.domaincontroller) -ea continue)  ) {
            } else { Cleanup ; BREAK ;} ;
            #$InputSplat.ADDesc="$(get-date -format 'MM/dd/yyyy') for $($InputSplat.OwnerMbx.samaccountname) $($InputSplat.ticket) -tsk" ;
            # flip it to $env:username
            $InputSplat.ADDesc="$(get-date -format 'MM/dd/yyyy') for $($InputSplat.OwnerMbx.samaccountname) $($InputSplat.ticket) -$($env:username)" ;

            # check for conflicting samaccountname, and increment
            $bConflicted=$false ;
            if (Get-User -identity $($InputSplat.samaccountname) -domaincontroller $($InputSplat.domaincontroller) -ea SilentlyContinue) {
                $InputSplat.samaccountname+="2"
                $smsg= "===Conflicting SamAccountName found ($($InputSplat.samaccountname)), incrementing SamAccountName $($InputSplat.samaccountname) to avoid conflict...===" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                $bConflicted=$true ;
            };

            If(!$InputSplat.NonGeneric -AND !($InputSplat.Equip -OR $InputSplat.Room)) {
                    $MbxSplat.shared=$true ;
            } ;
            # add Equipment, Room
            if($InputSplat.Equip){
                $MbxSplat.Equipment=$true ;
                $MbxSplat.Remove("shared");
            }elseif($InputSplat.Room){
                $MbxSplat.Room=$true ;
                $MbxSplat.Remove("shared");
            }

            #region NAME_CNAME_CLEANUP  ; #*------v NAME_CNAME_CLEANUP v------
            # build the name as an explicit LDAP cn Name strip
            # Name/LDAP cn= name, banned characters
            #$rgxBanCNName = '[\,\+\"\\\<\>\;\=\/]' 
            #$DnameClean is the diacritical/latin cleaned version fo the Displayname
            if($DnameClean){
                $MbxSplat.Name=-join (($DnameClean -replace $rgxBanCNName,'').ToCharArray()[0..63]) ;
            }else{
                $MbxSplat.Name=-join (($InputSplat.DisplayName -replace $rgxBanCNName,'').ToCharArray()[0..63]) ;
            }
            #endregion NAME_CNAME_CLEANUP ; #*------^ END NAME_CNAME_CLEANUP ^------
            #region DNAME_CLEANUP ; #*------v DNAME_CLEANUP v------
            # 64char limit on names
            if($InputSplat.DisplayName.length -gt 64){
                $smsg= "`n **** NOTE TRUNCATING DISPLAYNAME, -GT 64 CHARS!  ****`N" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;                
                $MbxSplat.DisplayName=$InputSplat.DisplayName.Substring(0,63) ;
            } else {
                $MbxSplat.DisplayName=$InputSplat.DisplayName;
            };
            #endregion DNAME_CLEANUP ; #*------^ END DNAME_CLEANUP ^------

            # using AMD (Automatic mailbox distribution), only subset are enabled for IsExcludedFromProvisioning $false
            # blank the db name and AMD will autopick db from avail block in site.
            $MbxSplat.database=$null ;
            $MbxSplat.password=ConvertTo-SecureString -a -f ($($InputSplat.pass));
            if($InputSplat.FirstName.length -gt 64){
                $smsg= "`n **** NOTE TRUNCATING FIRSTNAME, -GT 64 CHARS!  ****`N" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                $MbxSplat.FirstName=$InputSplat.FirstName.Substring(0,63) ;
            } else {
                $MbxSplat.FirstName=$InputSplat.FirstName;
            };
            $MbxSplat.Initials=$null;
            if($InputSplat.LastName.length -gt 64){
                $smsg= "`n **** NOTE TRUNCATING LASTNAME, -GT 64 CHARS!  ****`N" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                $MbxSplat.LastName=$InputSplat.LastName.Substring(0,63) ;
            } else {
                $MbxSplat.LastName=$InputSplat.LastName;
            };

            # 4:09 PM 4/23/2026 getting banned chars generating in email address, esp prim, which are BANNED, and supposed to be auto-fixing in Exch; but here we are. so check, remediate, and take out of EAP if needed.
            # then use the new prim cleaned as the UPN.
            #$rgxBanEmlDirName = "[\s@!\#\$%\^&\*\(\)\+=\{}\[\]\|\\:;','<>,\?/]" ; 
            # Local part≤ 64 characters ; Full address≤ 254 characters
            if($MbxSplat.FirstName -AND $MbxSplat.LastName){
                # dot-separated addr
                #$DirName="$([System.Text.RegularExpressions.Regex]::Replace($MbxSplat.LastName,'[^1-9a-zA-Z_]','')).$([System.Text.RegularExpressions.Regex]::Replace($MbxSplat.FirstName,'[^1-9a-zA-Z_]',''))" ;
                #$DirName = "$(-join (($MbxSplat.FirstName -replace $rgxUPNBan, '').ToCharArray()[0..63]))
                $FnameClean=Remove-StringDiacritic -string $MbxSplat.FirstName
                $FnameClean= Remove-StringLatinCharacters -string $FnameClean ;
                $LnameClean=Remove-StringDiacritic -string $MbxSplat.LastName ; 
                $LnameClean= Remove-StringLatinCharacters -string $LnameClean ;
                $DirName = "$($FnameClean).$($LnameClean)" ; 
                $DirName = -join (($DirName -replace $rgxBanEmlDirName, '').ToCharArray()[0..63]) ;
            } else {
                # no-dot, just the fname+lname concat'd, trimmed
                #$DirName=("$([System.Text.RegularExpressions.Regex]::Replace($MbxSplat.LastName,'[^1-9a-zA-Z_]',''))$([System.Text.RegularExpressions.Regex]::Replace($MbxSplat.FirstName,'[^1-9a-zA-Z_]',''))").trim() ;
                $LnameClean=Remove-StringDiacritic -string $MbxSplat.LastName ; 
                $LnameClean= Remove-StringLatinCharacters -string $MbxSplat.LastName ;
                $DirName = "$($LnameClean))" ; 
                $DirName = -join (($DirName -replace $rgxBanEmlDirName, '').ToCharArray()[0..63]) ;
            } ;
            switch ($Domain){
                # this probably damages the address!
              "$($DomTORfqdn)" {
                  $MbxSplat.userprincipalname="$($DirName)@$($toRmeta['o365_OPDomain'])" ;
              } ;
              "$($DomTOLfqdn)" {
                  $MbxSplat.userprincipalname="$($DirName)@$($toLmeta['o365_OPDomain'])" ;
              } ;
              default {
                  throw "unrecognized `Domain:$($Domain)!" ;
              } ;
            } ;

            $MbxSplat.samaccountname=$InputSplat.samaccountname;

            #region ALIAS_CLEANUP ; #*------v ALIAS_CLEANUP v------
            # Exchange.alias/AD.mailnickname - it's MORE RESTRICTIVE THAN SAMACCOUNTNAME! (HAD MAILBOXES THAT WOULDNT DISCOVER GET-XORECIPIENT  - MAY NOT HAVE BEEN REPLICATING HAD & CHAR IN ALIAS (PERMITTED IN SAMA)
            #$rgxPermitMailAlias = "[a-zA-Z0-9._-]" ; 
            $MbxSplat.Alias = -join ($DnameClean.ToCharArray() | ?{$_ -match '[a-zA-Z0-9._-]'} )[0..63] ; 
            try {$conflict = $null ; $conflict = get-recipient -id $MbxSplat.Alias -domaincontroller $MbxSplat.DomainController -EA 0} CATCH {} ;
            $incr = 1 ;
            Do{
                if($conflict){
                    $incr++ ;
                    $MbxSplat.Alias = "$(-join ($DnameClean.ToCharArray() | ?{$_ -match '[a-zA-Z0-9._-]'} )[0..62])$($incr)" ; 
                }
                try {$conflict = $null ; $conflict = get-recipient -id $MbxSplat.Alias -domaincontroller $MbxSplat.DomainController  -EA 0} CATCH {} ;
            }while($conflict) ;
            #endregion ALIAS_CLEANUP ; #*------^ END ALIAS_CLEANUP ^------

            $MbxSplat.RetentionPolicy=$TORMeta['RetentionPolicy'];
            $MbxSplat.domaincontroller=$($InputSplat.domaincontroller)

            #or nonshared, make them reset pw on logon
            if($NonGeneric -eq $true){
                write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):NonGeneric user, Forcing ResetPasswordOnNextLogon";
                $MbxSplat.ResetPasswordOnNextLogon=$true
                # move EAS pol into here - Shared have no pw, so how could they log onto EAS?
                $MbxSplat.ActiveSyncMailboxPolicy='Default';
            } else {
                $MbxSplat.ResetPasswordOnNextLogon=$false;
                # completely remove the policy spec (throws up on shared/equipment/room
                $MbxSplat.remove("ActiveSyncMailboxPolicy");
            };

            #for nonshared, make them reset pw on logon
            if($NonGeneric -eq $true){
                $smsg= "$((get-date).ToString("HH:mm:ss")):NonGeneric user, Forcing ResetPasswordOnNextLogon";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                $MbxSplat.ResetPasswordOnNextLogon=$true
            } ;

            # get the identity set to match
            $MbxSetSplat.identity=$MbxSplat.samaccountname ;
            $MbxSetSplat.CustomAttribute9=$null;
            $MbxSetSplat.CustomAttribute5=$null;

            if(!($InputSplat.Vscan)){
                # prompt if not explicitly set
                $smsg= "`a$($xxx)`nPROMPT:";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                $bRet= read-host "Enable Vscan?[Y/N]" ;
                # *** BREAKPOINT ;
                if($bRet.ToUpper() -eq "Y"){
                    $MbxSetSplat.CustomAttribute9 = $CU9Value ;
                } ;
            } else {
                $smsg= "$((get-date).ToString("HH:mm:ss")):-Vscan $($InputSplat.Vscan) parameter specified"; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                if($InputSplat.Vscan -eq "YES"){
                    $MbxSetSplat.CustomAttribute9 = $CU9Value  ;
                } elseif($InputSplat.Vscan -eq "NO"){
                    $MbxSetSplat.CustomAttribute9 = $null  ;
                } ;  # if-E vscan
            }  # if-E Vscan;

            if($InputSplat.Cu5){
                $smsg= "-CU5 override detected, forcing CustomAttribute5:$( $InputSplat.Cu5 )" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                $MbxSetSplat.CustomAttribute5 = $InputSplat.Cu5 ;
            } else {
                # switch CA5 to keying off of Mgr - too likely to draw inapprop odddomain.com users in ($ADSiteCodeUS) $InputSplat.Owner
                # Owner doesn't have a cu5 attr, shift to the new OwnerMbx prop
                if($InputSplat.OwnerMbx.CustomAttribute5){
                    $smsg= "OwnerMbx has Cu5 set: Applying '$($InputSplat.OwnerMbx.CustomAttribute5)' to CU5 on new mbx" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                    $MbxSetSplat.CustomAttribute5 =$InputSplat.BaseUser.CustomAttribute5 ;
                } ;
            } ;

            # casmbx matrl
            $MbxSetCASmbx.identity=$MbxSplat.samaccountname ;
            #$MbxSetCASmbx.ActiveSyncMailboxPolicy="Default" ;
            if($NonGeneric -eq $true){
                write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):NonGeneric user, Forcing ResetPasswordOnNextLogon & EAS";
                $MbxSetCASmbx.ResetPasswordOnNextLogon=$true
                # move EAS pol into here - Shared have no pw, so how could they log onto EAS?
                $MbxSetCASmbx.ActiveSyncMailboxPolicy='Default';
            } else {
                # completely remove the policy spec (throws up on shared/equipment/room
                $smsg= "$((get-date).ToString("HH:mm:ss")):Shared/Room/Equipment:Removing ActiveSyncMailboxPolicy"; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                $MbxSetCASmbx.remove("ActiveSyncMailboxPolicy");
            };
            $MbxSetCASmbx.domaincontroller=$($InputSplat.domaincontroller) ;

            #  blank setting manager
            $ADSplat.Identity=$($MbxSplat.samaccountname) ;
            $ADSplat.Description=$($InputSplat.ADDesc) ;

            $UsrSplat.City=$($InputSplat.BUserAD.City);
            $UsrSplat.CountryOrRegion=$($InputSplat.BUserAD.CountryOrRegion);
            $UsrSplat.Fax=$($InputSplat.BUserAD.Fax);
            $UsrSplat.Office=$($InputSplat.BUserAD.Office);
            $UsrSplat.PostalCode=$($InputSplat.BUserAD.PostalCode);
            $UsrSplat.StateOrProvince=$($InputSplat.BUserAD.StateOrProvince);
            $UsrSplat.StreetAddress=$($InputSplat.BUserAD.StreetAddress);
            $UsrSplat.Title=$Title;
            $UsrSplat.Phone=$Phone

            $ChangeLog="$((get-date -format "MM/dd/yyyy"))" ;

            if($($InputSplat.Ticket)){$ChangeLog+=" #$($InputSplat.Ticket)" } ;
            $ChangeLog+=" for $($InputSplat.Owner) -$($AdminInits)" ;

            $smsg= "=== v Input Specifications v ===";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
            $smsg= "`$InputSplat:`n$(($InputSplat|out-string).trim())" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
            if(!($MbxSplat.Shared -OR $MbxSplat.Room -OR $mbxSplat.Equipment)){
                $smsg= "`nInitial Password: $($InputSplat.pass)";

            } ;
            $smsg= "`$MbxSplat:`n$(($MbxSplat|out-string).trim())" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
            $smsg= "`$MbxSetSplat:`n$(($MbxSetSplat|out-string).trim())" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
            $smsg= "`$MbxSetCASmbx:`n$(($MbxSetCASmbx|out-string).trim())" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
            $smsg= "`$ADSplat:`n$(($ADSplat|out-string).trim())" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
            $smsg= "`$UsrSplat:`n$(($UsrSplat|out-string).trim())" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;

            $smsg= "=== ^ Input Specifications ^ ===";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;

            #endregion DATAPREP ; # ------

            #region MAKECHANGES ; # ------

            $smsg= "Checking for existing $($InputSplat.SamAccountname)..."  ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
            if($bDebug){$smsg= "`$Mbxsplat.DisplayName:$($Mbxsplat.DisplayName)"; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; }
            # 12:30 PM 4/1/2016 if samaccountname checks always bump to 2, this will never match existing!
            if($bConflicted) {
                $smsg= "Prior Conflict already found!" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                $oMbx = (get-mailbox -identity  $($Mbxsplat.samaccountname) -domaincontroller $($InputSplat.DomainController) -ea silentlycontinue) ;
            } else {

            } ;

            if($oMbx){
                if($bDebug){
                    $smsg= "Existing found: `$InputSplat.DisplayName:$($InputSplat.DisplayName)" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                    $smsg= "`$oMbx.DN:$($oMbx.DistinguishedName)" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                } ;
            } else {
                $smsg= "$($Mbxsplat.DisplayName) Not found. ..."  ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                $smsg= "Whatif $($Mbxsplat.Name) creation...";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                $MbxSplat.Whatif=$true ;
                TRY{
                    New-Mailbox @MbxSplat  ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                } ;                

                $smsg= "$((get-date).ToString("HH:mm:ss")):Continue with $($Mbxsplat.Name) creation?...`a";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                if($NoPrompt) {$bRet= "YYY"} else { $bRet=Read-Host "Enter YYY to continue. Anything else will exit" ;} ;
                if ($bRet.ToUpper() -eq "YYY") {
                    # *** BREAKPOINT ;
                    $smsg= "Executing...";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                    if($Whatif){
                        $smsg= "SKIPPING EXEC: Whatif-only pass";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                    } else {
                        $MbxSplat.Whatif=$false ;
                        $Exit = 0 ;
                        $smsg = "New-Mailbox  w`n$(($MbxSplat|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        # do loop until up to 4 retries...
                        Do {
                            TRY {
                                $oNMbx = New-Mailbox @MbxSplat -ea Stop ;
                                $Exit = $Retries ;
                            } Catch {
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
                                #Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                                Start-Sleep -Seconds $RetrySleep ;
                                $Exit ++ ;
                                #$smsg= "Failed to exec cmd because: $($Error[0])" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                $smsg= "Try #: $Exit" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                If ($Exit -eq $Retries) {$smsg= "Unable to exec cmd!"; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; } ;
                            } # try-E
                            
                        } Until ($Exit -eq $Retries) # loop-E

                        $MbxSplat.Whatif=$true ;
                    } ;

                    if($Whatif){
                        $smsg= "SKIPPING EXEC: Whatif-only pass";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                    } else {
                        if($oNMbx){
                            write-verbose "(using returned New-Mailbox output object...)" ; 
                            #do {Write-Host "." -NoNewLine;Start-Sleep -s 1} until ($oMbx = (get-mailbox -identity $Mbxsplat.samaccountname -domaincontroller $Mbxsplat.DomainController -ea silentlycontinue)) ;
                            $oMbx = $oNMbx ;
                        } else { 
                            # if $oNMbx is properly output from new-mailbox, use it
                            write-verbose "(New-Mailbox output did not return an object, reusing input SamAccountname for gmbx...)" ; 
                            do {Write-Host "." -NoNewLine;Start-Sleep -s 1} until ($oMbx = (get-mailbox -identity $oNMbx.samaccountname -domaincontroller $Mbxsplat.DomainController -ea silentlycontinue)) ;
                        } ; 
                        if($bDebug){
                            $smsg= "`$oMbx:$($Mbxsplat.DisplayName)" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                            $smsg= "`$oMbx.DN:$($oMbx.DistinguishedName)" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        } ;
                    }  # if-E
                } else {
                    $smsg= "INVALID KEY ABORTING NO CHANGE!" ;  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                    BREAK ;
                } ;

            } # if-E No oMbx

            if($oMbx){

                if($Whatif){
                        $smsg= "SKIPPING REMAINING AD CMDS - NO OBJECT YET: Whatif-only pass";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                } else {
                    $MbxSetSplat.Whatif=$false ;
                    $Exit = 0 ;
                    if($bDebug) {$smsg= "$((get-date).ToString("HH:mm:ss")):Updating Mbx" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; }
                    $smsg = "Set-Mailbox w`n$(($MbxSetSplat|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Do {
                        Try {
                            Set-Mailbox @MbxSetSplat -ea Stop  ;
                            $Exit = $Retries ;
                        } Catch {
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
                            #Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                            Start-Sleep -Seconds $RetrySleep
                            $Exit ++ ;
                            $smsg= "Failed to execute Set-Mailbox because: $($Error[0])" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                            $smsg= "Try #: $($Exit)" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                            If ($Exit -eq $Retries) {
                                $smsg= "Unable to update mailbox! (Set-Mailbox)" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                            } ;
                        } # try-E
                    } Until ($Exit -eq $Retries) # loop-E

                    if($bDebug) {$smsg= "$((get-date).ToString("HH:mm:ss")):Setting CASMbx settings" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; } ;
                    $MbxSetCASmbx.Whatif=$false ;
                    $smsg = "set-CASMailbox w`n$(($MbxSetCASmbx|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $Exit = 0 ;
                    Do {
                        Try {
                            set-CASMailbox @MbxSetCASmbx -ea Stop ;
                            $Exit = $Retries ;
                        } Catch {
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
                            #Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                            Start-Sleep -Seconds $RetrySleep
                            $Exit ++ ;
                            $smsg= "Failed to execute Set-CASMailbox because: $($Error[0])" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                            $smsg= "Try #: $Exit" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                            If ($Exit -eq $Retries) {$smsg= "Unable to update mailbox! (Set-CASMailbox)" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; } ;
                        } # try-E
                    } Until ($Exit -eq $Retries) # loop-E

                    if($MbxSetSplat.CustomAttribute5){
                        # force trigger EAP toggle
                        $smsg= "$((get-date).ToString("HH:mm:ss")):(toggling EAP to force variant email...)";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        set-mailbox $($InputSplat.samaccountname) -domaincontroller $($InputSplat.domaincontroller) -EmailAddressPolicyEnabled $false  ;Start-Sleep 1; set-mailbox $($InputSplat.samaccountname) -domaincontroller $($InputSplat.domaincontroller) -EmailAddressPolicyEnabled $true;
                    } ;

                    #=========== V NOTES PARSER
                    <# This block takes an existing Adobject Notes field and parses it and updates fields with a new ChangeLog reference
                    #>
                    # we should _append_ the $InfoStr into any existing Info for the object
                    # new obj properties for OtherInfo
                    $ADOtherInfoProps=@{
                        TargetMbx=$null ;
                        PermsExpire=$null ;
                        Incident=$null ;
                        Admin=$null ;
                        BusinessOwner=$null ;
                        ITOwner=$null ;
                        Owner=$null ;
                        ChangeLog=$null ;
                        ADNotes=$null ;
                    } ;
                    $oADOtherInfo = New-Object PSObject -Property $ADOtherInfoProps ;
                    $Exit = 0 ; # zero out $exit each new cmd try/retried

                    Do {
                        Try {
                            $oADUsr=get-aduser -identity $($MbxSplat.SamAccountname) -Properties * -server $($InputSplat.DomainController) -ErrorAction stop ;
                            # break-exit here, completes the Until block
                            $Exit = $Retries ;
                        } Catch {
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
                            #Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                            Start-Sleep -Seconds $RetrySleep ;
                            $Exit ++ ;
                            $smsg= "Failed to exec cmd because: $($Error[0])" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                            $smsg= "Try #: $($Exit)" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                            If ($Exit -eq $Retries) {$smsg= "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; } ;
                        } # try-E
                    } Until ($Exit -eq $Retries) # loop-E

                    # MailContact object has an explicit notes field
                    # but ADObjects just have info

                    if($oADUsr.info){
                        # if existing notes, grab all but the defined tags, and then we'll append a $ChangeLog to the head, and an Owner to the tail.
                        # mailcontact
                        $lns = ($oADUsr.info.tostring().split("`n")) ;
                        $UpdInfo=$null;
                        foreach ($ln in $lns) {
                                # add stock Owner
                                if($ln -match "^(TargetMbx|PermsExpire|Incident|Admin|BusinessOwner|ITOwner|Owner):.*$"){
                                    # it's part of a defined Info tag
                                    $matches=$null ;
                                    # ingest the matches and throw away the lines
                                    if($ln -match "(?<=TargetMbx:)\w+" ){ $oADOtherInfo.TargetMbx = $matches[0] } ; $matches=$null ;
                                    if($ln -match "(?<=PermsExpire:)\d+\/\d+/\d+" ) {$oADOtherInfo.PermsExpire = (get-date $matches[0]) ;   } ; ; $matches=$null ;
                                    if($ln -match "(?<=Incident:)^\d{5,6}$"){ $oADOtherInfo.Incident = $matches[0] ;  } ;  $matches=$null ;
                                    if($ln -match "(?<=Admin:)\w*"){ $oADOtherInfo.Admin = $matches[0] ;   } ; $matches=$null ;
                                    if($ln -match "(?<=BusinessOwner:)\w{2,20}"){ $oADOtherInfo.BusinessOwner = $matches[0] ;   } ; $matches=$null ;
                                    if($ln -match "(?<=ITOwner:)\w{2,20}"){ $oADOtherInfo.ITOwner = $matches[0] ;  } ; $matches=$null ;
                                    if($ln -match "(?<=Owner:)\w{2,20}$"){ $oADOtherInfo.Owner = $matches[0] ;  } ; $matches=$null ;
                                } else {
                                    $UpdInfo+="$($ln)`r`n" ;
                                } ;
                        }# loop-E ;

                        # do compare if detected existing owner: tag     $oADOtherInfo.Owner
                        if($oADOtherInfo.Owner){
                            if($oADOtherInfo.Owner -ne $InputSplat.Owner){
                                # preserve original owner, if change would update it
                                $UpdInfo+="`r`nOwner:$($oADOtherInfo.Owner)" ;
                                # update the $InputSplat Owner to reflect corrected owner
                                $smsg= "$((get-date).ToString("HH:mm:ss")):NOTE: $($oADUsr.Name) has an eisting Owner value, deferring ownership to existing value: $($oADOtherInfo.Owner)";
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                $InputSplat.Owner = $oADOtherInfo.Owner ;
                                $InputSplat.OwnerMbx=(get-mailbox -identity $($InputSplat.Owner) $(InputSplat.domaincontroller) -ea stop) ;
                            } else {
                                $UpdInfo+="`r`nOwner:$($InputSplat.Owner)" ;
                            }
                        } else {
                            #$UpdInfo+="`r`nOwner:$($oADOtherInfo.Owner)" ;
                            $UpdInfo+="`r`nOwner:$($InputSplat.Owner)" ;
                        }

                        if($bDebug){
                            $smsg= "$((get-date).ToString("HH:mm:ss")):New Info field:"; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                            $UpdInfo | out-string | out-default ;
                        } ;

                    } else {
                        # new notes
                        $UpdInfo="Owner:$($InputSplat.Owner)" ;
                    } # if-E populated notes ;

                    # prepend the new $ChangeLog to the top and assign to .notes
                    $UpdInfo="$($ChangeLog)`r`n$($UpdInfo)" ;
                    # strip empty lines in there too
                    $UpdInfo = $UpdInfo -replace '(?ms)(?:\r|\n)^\s*$' ;
                    # mailcontact
                    #$AContactSplat.notes="$($UpdInfo)" ;
                    # ADU
                    $ADOtherInfoProps.ADNotes="$($UpdInfo)" ;
                    #=========== ^ NOTES PARSER

                    if($Whatif){
                        $smsg= "Whatif $($Mbxsplat.DisplayName) Update..."; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        $ADSplat.whatif = $true ;
                        Set-ADUser @ADSplat -Replace @{info="$($UpdInfo)"} ;
                    } else {
                        $smsg= "Executing $($Mbxsplat.DisplayName) Update...";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        $ADSplat.whatif = $false ;
                        $smsg = "Set-ADUser w`n$(($ADSplat|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $Exit = 0 ; # zero out $exit each new cmd try/retried
                        Do {
                            Try {
                                Set-ADUser @ADSplat -Replace @{info="$($UpdInfo)"} -ErrorAction stop ;
                                $Exit = $Retries ;
                            } Catch {
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
                                #Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                                Start-Sleep -Seconds $RetrySleep ;
                                $Exit ++ ;
                                $smsg= "Failed to exec cmd because: $($Error[0])" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                $smsg= "Try #: $Exit" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                If ($Exit -eq $Retries) {$smsg= "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; } ;
                            } # try-E
                        } Until ($Exit -eq $Retries) # loop-E

                        $ADSplat.whatif = $true ;
                    }  ;

                    #region REPAIR_EMAILADDRESSES ; #*------v REPAIR_EMAILADDRESSES v------
                    #$OmBX = Repair-xopEmailAddresses -identity $oMbx.ExchangeGuid.guid -RestoreEAPPolicy:$true -whatif:$($whatif) -domaincontroller:$($domaincontroller) ; 
                    # DON'T USE -RestoreEAPPolicy:$true  IT REVERSES THE FIXES IT JUST DID! (BAD EOP AND RECIPIENTUPDATE CODE IN XOP IS THE SOURCE)
                    $OmBX = Repair-xopEmailAddresses -identity $oMbx.ExchangeGuid.guid -whatif:$($whatif) -domaincontroller:$($domaincontroller) ; 
                    #endregion REPAIR_EMAILADDRESSES ; #*------^ END REPAIR_EMAILADDRESSES ^------

                    #region UPNFromEmail ; # ------
                    # pull the SMTP: addr and use it to force the UPN

                    if($Whatif){
                        $smsg= "Whatif skipping UPN Update...";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                    } else {
                        $forestUPNSuffixes=get-adforest -ea stop | Select-Object -expand upnsuffixes ;
                        $forestdom=$forestUPNSuffixes |Where-Object{$_ -match $rgxTTCDomainsLegacy}
                        if($BrandUPN){
                            $smsg = "-BrandUPN:`$TRUE: USING PRIMARY SMTP IF A FORESTDOMAIN UPN SUFFIX" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            if($AduPrimeSMTP = ($OMbx.emailaddresses |?{$_ -cmatch '^SMTP:'}).replace('SMTP:','')){
                                $newUPN = $AduPrimeSMTP
                            } ;                            
                            if( $newUPN ){
                            }else{
                                $smsg = "No `$newUPN generated: Defaulting to Dirname@forestdomain UPN suffix" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                <#
                                Do {
                                    $smsg= "Waiting for ADUser to return email addresses" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                    write-host "." -NoNewLine;Start-Sleep -m (1000 * 5)
                                    $dirname=(Get-ADUser -identity $oMbx.samaccountname -server $InputSplat.domaincontroller -Properties proxyAddresses -ea 0 | Select-Object -Expand proxyAddresses | Where-Object {$_ -clike "SMTP:*"});
                                } Until (($dirname)) ;
                                #>
                                if($dirName = $AduPrimeSMTP.tostring().replace("SMTP:","").split("@")[0]){
                                    $newUPN="$($dirname)@$($forestdom)";                                   
                                } else {
                                    throw "invalid `$dirname$($dirname) type returned! (non-string)" ;
                                } ;                                 
                            }
                        }else{
                            $smsg = "-BrandUPN:`$FALSE: COERCING CENTRAL FORESTDOMAIN UPN SUFFIX" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            if($dirName = $AduPrimeSMTP.tostring().replace("SMTP:","").split("@")[0]){
                                $newUPN="$($dirname)@$($forestdom)";                                   
                            } else {
                                throw "invalid `$dirname$($dirname) type returned! (non-string)" ;
                            } ;  
                        } ; 
                        # strip chars allowed in eml, *NOT* allowed in upns:
                        <# Nov 06, 2004 05:27 PM
                        From MSDN: User account names are limited to 20 characters and group names are limited to 256 characters. In addition, account names cannot be terminated by a period and they cannot include commas or any of the following printable characters: ", /, \, [, ], :, |, <, >, +, =, ;, ?, *. Names also cannot include characters in the range 1-31, which are nonprintable

                        [Prepare to provision users through directory synchronization to Office 365 | Microsoft Docs](https://docs.microsoft.com/en-us/office365/enterprise/prepare-for-directory-synchronization):
                        maximum number of characters for the userPrincipalName attribute is 113.
                        Maximum number of characters for the user name that is in front of the at sign (@): 64
                        Maximum number of characters for the domain name following the at sign (@): 48

                        Invalid characters: \ % & * + / = ? { } | < > ( ) ; : , [ ] "
                        An umlaut is also an invalid character.
                        #>
                        # samaccountname & alias should be fine, I'm running an alphanumeric or underscore replacement on the sama, which also goes to the alias
                        $newUPN = Remove-StringDiacritic -string $newUPN ;
                        $newUPN = Remove-StringLatinCharacters -string $newUPN ;
                        # yank ampersand,apostrophe,hyphen,underscore,%, +
                        $newUPN = $newUPN.replace("&", "").replace("'","").replace("-","").replace("%","").replace("+","").replace("SMTP:","")
                        if($bDebug){$smsg= "$((get-date).ToString("HH:mm:ss")):`$newUPN:$($newUPN)"; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; } ;
                        $pltSetADU2 = @{
                            identity=$oMbx.samaccountname ;UserPrincipalName=$newUPN ;server=$InputSplat.domaincontroller ;ErrorAction='Stop' ;
                        } ; 
                        $smsg = "Set-ADUser w`n$(($pltSetADU2|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        # retry support
                        $Exit = 0 ;
                        Do {
                            Try {
                                Set-ADUser @pltSetADU2 ;
                                #-identity $oMbx.samaccountname -UserPrincipalName $newUPN -server $InputSplat.domaincontroller -ErrorAction Stop;
                                $Exit = $Retries ;
                            } Catch {
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
                                #Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                                Start-Sleep -Seconds $RetrySleep ;
                                $Exit ++ ;
                                $smsg= "Failed to exec cmd because: $($Error[0])" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                $smsg= "Try #: $($Exit)" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                If ($Exit -eq $Retries) {$smsg= "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; } ;
                            } # try-E
                        } Until ($Exit -eq $Retries) # loop-E

                    }; 
                
                    #endregion UPNFromEmail ; # ------
                    #endregion MAKECHANGES ; # ------

                    #region REPORTING ; # ------
                    do{
                        $smsg= "===REVIEW SETTINGS:=== " ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        $mbxo = get-mailbox -Identity $($InputSplat.samaccountname) -domaincontroller $($InputSplat.domaincontroller) ;
                        $cmbxo= Get-CASMailbox -Identity $($MbxSplat.samaccountname) -domaincontroller $($InputSplat.domaincontroller) ;
                        $aduprops="GivenName,Surname,Manager,Company,Office,Title,StreetAddress,City,StateOrProvince,c,co,countryCode,PostalCode,Phone,Fax,Description" ;
                        $ADu = get-ADuser -Identity $($InputSplat.samaccountname) -properties * -server $($InputSplat.domaincontroller)| Select-Object *;
                        $smsg= "User Email:`t$(($mbxo.WindowsEmailAddress.tostring()).trim())" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        $smsg= "Owner Email:`t$(($InputSplat.OwnerMbx.WindowsEmailAddress.tostring()).trim())" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        $smsg= "Mailbox Information:" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        $smsg= "$(($mbxo | select @{Name='LogonName';Expression={$_.SamAccountName }},Name,DisplayName,Alias,database,UserPrincipalName,RetentionPolicy,CustomAttribute5,CustomAttribute9,RecipientType,RecipientTypeDetails | out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        $smsg= "`n$(($Adu | select GivenName,Surname,Manager,Company,Office,Title,StreetAddress,City,StateOrProvince,c,co,countryCode,PostalCode,Phone,Fax,Description | out-string).trim())";
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        if($NonGeneric -eq $false){
                            $smsg= "ActiveSyncMailboxPolicy:$($cmbxo.ActiveSyncMailboxPolicy.tostring())" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        } ;
                        $smsg= "Description: $($Adu.Description.tostring())";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        $smsg= "Info: $($Adu.info.tostring())";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        if(!($MbxSplat.Shared -OR $MbxSplat.Room -OR $MbxSplat.Equipment  )){
                            $smsg= "Initial Password: $(($InputSplat.pass | out-string).trim())" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        } ;
                        $bRet=Read-Host "Enter Y to Refresh Review (replication latency)." ;
                    } until ($bRet -ne "Y");
                    $smsg= "$xxx`n";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                    if($InputSplat.NonGeneric){
                        $smsg= "(projected Permissions SecGrp name: $($InputSplat.PermGrpName))`n" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                    } ;

                    #endregion REPORTING ; # ------

                }  # if-E whatif/exec
                
                # opt emit new mbx object
                if(-not($NoOutput)){                    
                    $smsg = "-NoOutput:`$false: returning mailbox object to pipeline" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    $mbxo  | write-output 
                }else{
                    $smsg = "(-NoOutput:`$true: suppressing new mailbox return)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;                                        
                }
            }  # if-E $ombx

        } else { $smsg= "OU $(OU) not found. ABORTING!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; } ;



    } # PROC-E ;

    END {


    } # END-E
}
#endregion NEW_MAILBOXSHARED ; #*------^ END new-MailboxShared ^------
