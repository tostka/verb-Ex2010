#*------v new-MailboxShared.ps1 v------
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
    FileName    : 
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite:	URL
    AddedTwitter:	URL
    REVISIONS
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
        [Parameter(Mandatory=$true,HelpMessage="Incident number for the change request[[int]nnnnnn]")]
        [int]$Ticket,
        [Parameter(HelpMessage="Option to hardcode a specific DC [-domaincontroller xxxx]")]
        [string]$domaincontroller,
        [Parameter(HelpMessage="Calling script path (used for log construction)[-ParentPath c:\pathto\script.ps1]")]
        [string]$ParentPath,
        [Parameter(HelpMessage="Suppress YYY confirmation prompts [-NoPrompt]")]
        [switch] $NoPrompt,
        [Parameter(HelpMessage='Debugging Flag [$switch]')]
        [switch] $showDebug,
        [Parameter(HelpMessage='Whatif Flag [$switch]')]
        [switch] $whatIf
    ) ;

    BEGIN {
        $verbose = ($VerbosePreference -eq "Continue") ; 
        # Get the name of this function
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        # Get parameters this function was invoked with
        $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
        $continue = $true ;
        switch -regex ($env:COMPUTERNAME){
            ($rgxMyBoxW){ $LocalInclDir="c:\usr\work\exch\scripts" ; }
            ($rgxProdEx2010Servers){ $LocalInclDir="c:\scripts" ; }
            ($rgxLabEx2010Servers){ $LocalInclDir="c:\scripts" ; }
            ($rgxProdL13Servers){ $LocalInclDir="c:\scripts" ; }
            ($rgxLabL13Servers){ $LocalInclDir="c:\scripts" ; }
        } ;
        $Retries = 4 ; # number of re-attempts
        $RetrySleep = 5 ; # seconds to wait between retries
        # $rgxCU5 = [infra file]
        # OU that's used when can't find any baseuser for the owner's OU, default to a random shared from ($ADSiteCodeUS) (avoid crapping out):
        $FallBackBaseUserOU = "$($DomTORfqdn)/($ADSiteCodeUS)/Generic Email Accounts" ;

        # strings are: "[tModName];[tModFile];tModCmdlet"
        $tMods = @() ;
        #$tMods+="verb-Auth;C:\sc\verb-Auth\verb-Auth\verb-Auth.psm1;get-password" ;
        $tMods+="verb-logging;C:\sc\verb-logging\verb-logging\verb-logging.psm1;write-log";
        $tMods+="verb-IO;C:\sc\verb-IO\verb-IO\verb-IO.psm1;Add-PSTitleBar" ;
        $tMods+="verb-Mods;C:\sc\verb-Mods\verb-Mods\verb-Mods.psm1;check-ReqMods" ;
        #$tMods+="verb-Desktop;C:\sc\verb-Desktop\verb-Desktop\verb-Desktop.psm1;Speak-words" ;
        #$tMods+="verb-dev;C:\sc\verb-dev\verb-dev\verb-dev.psm1;Get-CommentBlocks" ;
        $tMods+="verb-Text;C:\sc\verb-Text\verb-Text\verb-Text.psm1;Remove-StringDiacritic" ;
        #$tMods+="verb-Automation.ps1;C:\sc\verb-Automation.ps1\verb-Automation.ps1\verb-Automation.ps1.psm1;Retry-Command" ;
        #$tMods+="verb-AAD;C:\sc\verb-AAD\verb-AAD\verb-AAD.psm1;Build-AADSignErrorsHash";
        $tMods+="verb-ADMS;C:\sc\verb-ADMS\verb-ADMS\verb-ADMS.psm1;load-ADMS";
        $tMods+="verb-Ex2010;C:\sc\verb-Ex2010\verb-Ex2010\verb-Ex2010.psm1;Connect-Ex2010";
        #$tMods+="verb-EXO;C:\sc\verb-EXO\verb-EXO\verb-EXO.psm1;Connect-Exo";
        #$tMods+="verb-L13;C:\sc\verb-L13\verb-L13\verb-L13.psm1;Connect-L13";
        $tMods+="verb-Network;C:\sc\verb-Network\verb-Network\verb-Network.psm1;Send-EmailNotif";
        #$tMods+="verb-Teams;C:\sc\verb-Teams\verb-Teams\verb-Teams.psm1;Connect-Teams";
        #$tMods+="verb-SOL;C:\sc\verb-SOL\verb-SOL\verb-SOL.psm1;Connect-SOL" ;
        #$tMods+="verb-Azure;C:\sc\verb-Azure\verb-Azure\verb-Azure.psm1;get-AADBearToken" ;
        foreach($tMod in $tMods){
            $tModName = $tMod.split(';')[0] ;             $tModFile = $tMod.split(';')[1] ;             $tModCmdlet = $tMod.split(';')[2] ;
            $smsg = "( processing `$tModName:$($tModName)`t`$tModFile:$($tModFile)`t`$tModCmdlet:$($tModCmdlet) )" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            if($tModName -eq 'verb-Network' -OR $tModName -eq 'verb-Text' -OR $tModName -eq 'verb-IO'){
                write-host "GOTCHA!:$($tModName)" ;
            } ;
            $lVers = get-module -name $tModName -ListAvailable -ea 0 ;
            if($lVers){                 $lVers=($lVers | sort version)[-1];                 try {                     import-module -name $tModName -RequiredVersion $lVers.Version.tostring() -force -DisableNameChecking -verbose:$false                }   catch {                      write-warning "*BROKEN INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;import-module -name $tModDFile -force -DisableNameChecking -verbose:$false                } ;
            } elseif (test-path $tModFile) {                 write-warning "*NO* INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;                 try {import-module -name $tModDFile -force -DisableNameChecking -verbose:$false}                 catch {                     write-error "*FAILED* TO LOAD MODULE*:$($tModName) VIA $(tModFile) !" ;                     $tModFile = "$($tModName).ps1" ;                     $sLoad = (join-path -path $LocalInclDir -childpath $tModFile) ;                     if (Test-Path $sLoad) {                         write-verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                         . $sLoad ;                         if ($showdebug) { write-verbose "Post $sLoad" };                     } else {                         $sLoad = (join-path -path $backInclDir -childpath $tModFile) ;                         if (Test-Path $sLoad) {                             write-verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                             . $sLoad ;                             if ($showdebug) { write-verbose "Post $sLoad" };                         } else {                             Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;                             exit;                         } ;                     } ;                 } ;             } ;
            if(!(test-path function:$tModCmdlet)){                 write-warning -verbose:$true  "UNABLE TO VALIDATE PRESENCE OF $tModCmdlet`nfailing through to `$backInclDir .ps1 version" ;                 $sLoad = (join-path -path $backInclDir -childpath "$($tModName).ps1") ;                 if (Test-Path $sLoad) {                     write-verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                     . $sLoad ;                     if ($showdebug) { write-verbose "Post $sLoad" };                     if(!(test-path function:$tModCmdlet)){                         write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO CONFIRM `$tModCmdlet:$($tModCmdlet) FOR $($tModName)" ;                     } else {                         write-verbose  "(confirmed $tModName loaded: $tModCmdlet present)"                     }                 } else {                     Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;                     exit;                 } ;
            } else {                 write-verbose  "(confirmed $tModName loaded: $tModCmdlet present)"             } ;
        } ;  # loop-E
        #*------^ END MOD LOADS ^------

        if($ParentPath){
            $rgxProfilePaths='(\\Documents\\WindowsPowerShell\\scripts|\\Program\sFiles\\windowspowershell\\scripts)' ; 
            if($ParentPath -match $rgxProfilePaths){
                $ParentPath = "$(join-path -path 'c:\scripts\' -ChildPath (split-path $ParentPath -leaf))" ; 
            } ; 
            $logspec = start-Log -Path ($ParentPath) -showdebug:$($showdebug) -whatif:$($whatif) ;
            if($logspec){
                $logging=$logspec.logging ;
                $logfile=$logspec.logfile ;
                $transcript=$logspec.transcript ;
            } else {$smsg = "Unable to configure logging!" ; write-warning "$((get-date).ToString('HH:mm:ss')):$($sMsg)" ; Exit ;} ;
        } else {$smsg = "No functional `$ParentPath found!" ; write-warning "$((get-date).ToString('HH:mm:ss')):$($sMsg)" ;  Exit ;} ;
        if($whatif){
            $logfile=$logfile.replace("-BATCH","-BATCH-WHATIF") ;
            $transcript=$transcript.replace("-BATCH","-BATCH-WHATIF") ;
        } else {
            $logfile=$logfile.replace("-BATCH","-BATCH-EXEC") ;
            $transcript=$transcript.replace("-BATCH","-BATCH-EXEC") ;
        } ;
        if($Ticket){
            $logfile=$logfile.replace("-BATCH","-$($Ticket)") ;
            $transcript=$transcript.replace("-BATCH","-$($Ticket)") ;
        } else {
            $logfile=$logfile.replace("-BATCH","-nnnnnn") ;
            $transcript=$transcript.replace("-BATCH","-nnnnnn") ;
        } ;

        $xxx="====VERB====";
        $xxx=$xxx.replace("VERB","NewMbx") ;
        $BARS=("="*10);

        $reqMods+="Add-PSTitleBar;Remove-PSTitleBar".split(";") ;
        $reqMods+="Test-TranscriptionSupported;Test-Transcribing;Stop-TranscriptLog;Start-IseTranscript;Start-TranscriptLog;get-ArchivePath;Archive-Log;Start-TranscriptLog".split(";") ;
        $reqMods=$reqMods| select -Unique ;

        #region SPLATDEFS ; # ------
        # dummy hashes
        if($host.version.major -ge 3){
            $InputSplat=[ordered]@{
                Dummy = $null ;
            } ;
            $MbxSplat=[ordered]@{
                Dummy = $null ;
            } ;
            $MbxSetSplat=[ordered]@{
                Dummy = $null ;
            } ;
            $MbxSetCASmbx=[ordered]@{
                Dummy = $null ;
            } ;
            $ADSplat=[ordered]@{
                Dummy = $null ;
            } ;
            $UsrSplat=[ordered]@{
                Dummy = $null ;
            } ;
        } else {
            $InputSplat=@{
                Dummy = $null ;
            } ;
            $MbxSplat=@{
                Dummy = $null ;
            } ;
            $MbxSetSplat=@{
                Dummy = $null ;
            } ;
            $MbxSetCASmbx=@{
                Dummy = $null ;
            } ;
            $ADSplat=@{
                Dummy = $null ;
            } ;
            $UsrSplat=@{
                Dummy = $null ;
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
        if($DisplayName){$InputSplat.DisplayName=$DisplayName};
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
        if ($Cu5){
            if ($Cu5 -match $rgxCU5 ) {
                # pulled switch out, it wasn't actually translating, just rgx of the final tags
                $InputSplat.Cu5=$Cu5;
            } else {
                $InputSplat.Cu5=$null ;
            }  ;
        }; #  # if-E Cu5


        if($SiteOverride){$InputSplat.SiteOverride=$SiteOverride};
        if($Ticket){$InputSplat.Ticket=$Ticket};

        #endregion SPLATDEFS ; # ------
        #region LOADMODS ; # ------
        $rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" ;
        #$rgxEx10HostName=[infra file]
        $rgxRemsPssName="^(Exchange2010|Session\sfor\simplicit\sremoting\smodule\sat\s.*)" ;
        $rgxSnapPssname="^Session\d{1}$" ;
        $rgxEx2010SnapinName="^Microsoft\.Exchange\.Management\.PowerShell\.E2010$";
        $Ex2010SnapinName="Microsoft.Exchange.Management.PowerShell.E2010" ;

        #
        #LEMS detect: IdleTimeout -ne -1
        if(get-pssession |?{($_.configurationname -eq 'Microsoft.Exchange') -AND ($_.ComputerName -match $rgxEx10HostName) -AND ($_.IdleTimeout -ne -1)} ){
            write-verbose  "$((get-date).ToString('HH:mm:ss')):LOCAL EMS detected" ;
            $Global:E10IsDehydrated=$false ;
        # REMS detect dleTimeout -eq -1
        } elseif(get-pssession |?{$_.configurationname -eq 'Microsoft.Exchange' -AND $_.ComputerName -match $rgxEx10HostName -AND ($_.IdleTimeout -eq -1)} ){
            write-verbose  "$((get-date).ToString('HH:mm:ss')):REMOTE EMS detected" ;
            $reqMods+="get-GCFast;Get-ExchangeServerInSite;connect-Ex2010;Reconnect-Ex2010;Disconnect-Ex2010;Disconnect-PssBroken".split(";") ;
            if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
            reconnect-ex2010 ;
            $Global:E10IsDehydrated=$true ;
        } else {
            write-verbose  "$((get-date).ToString('HH:mm:ss')):No existing Ex2010 Connection detected" ;
            # Server snapin defer
            if(($host.version.major -lt 3) -AND (get-service MSExchangeADTopology -ea SilentlyContinue)){
                write-verbose -verbose:$bshowVerbose  "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Loading Local Server EMS10 Snapin" ;
                $reqMods+="Load-EMSSnap;load-EMSLatest".split(";") ;
                if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
                Load-EMSSnap ;
                $Global:E10IsDehydrated=$false ;
            } else {
                # if you want REMS - (assumed on new scripts)
                $reqMods+="connect-Ex2010;Reconnect-Ex2010;Disconnect-Ex2010;Get-ExchangeServerInSite;Disconnect-PssBroken".split(";") ;
                if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
                reconnect-ex2010 ;
                $Global:E10IsDehydrated=$true ;
            } ;
        } ;
        #

        # load ADMS
        $reqMods+="load-ADMS;get-AdminInitials".split(";") ;
        if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
        write-host -foregroundcolor darkgray "$((get-date).ToString('HH:mm:ss')):(loading ADMS...)" ;
        load-ADMS -cmdlet get-aduser,Set-ADUser,Get-ADGroupMember,Get-ADDomainController,Get-ADObject,get-adforest | out-null ; 

        $AdminInits=get-AdminInitials ;

        #region LOADMODS ; # ------

    }  # BEG-E ;

    PROCESS {

        #region DATAPREP ; # ------
        if ( ($InputSplat.OwnerMbx=(get-mailbox -identity $($InputSplat.Owner) -ea 0)) -OR ($InputSplat.OwnerMbx=(get-remotemailbox -identity $($InputSplat.Owner) -ea 0)) ){

        } else {
          throw "Unable to resolve $($InputSplat.Owner) to any existing OP or EXO mailbox" ;
          Cleanup ; Exit ;
        } ;

        $InputSplat.Domain=$($InputSplat.OwnerMbx.identity.tostring().split("/")[0]) ;
        $InputSplat.SiteCode=($InputSplat.OwnerMbx.identity.tostring().split('/')[1]) ;

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

        If($InputSplat.NonGeneric) {
            if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU  -Sitecode $SiteCode -Generic $false)   ) {

            } else { Cleanup ; Exit ;}
        } elseIf($Room -OR $Equipement) {
            if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU  -Sitecode $SiteCode -Resource $true ) ) {
            } else { Cleanup ; Exit ;}
        } else {
            if ( $MbxSplat.OrganizationalUnit = (Get-SiteMbxOU  -Sitecode $SiteCode -Generic $true ) ) {
            } else { Cleanup ; Exit ;}
        }

        # add forced office designation, to match $SiteCode/OU
        # New-Mailbox doesn't support both -Shared & -Office in the same syntax set, move it to $MbxSetSplat
        $MbxSetSplat.Office = $SiteCode ; 
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
                            } else { Cleanup ; Exit ;}
                        } else {
                            throw "Unable to resolve $($InputSplat.ManagedBy) to any existing OP or EXO mailbox" ;
                            Cleanup ; Exit ;
                        } ;
                    }
                    "MailUser" {
                        if ( ($Tmbx=(get-remotemailbox -identity $($InputSplat.Owner) -ea stop)) ){
                            if($showdebug){ $smsg= "Owner MailUser detected" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;  } ;
                            # base users off of Owner
                            if ( $InputSplat.BaseUser=(get-Remotemailbox -identity $($InputSplat.Owner) -domaincontroller $($InputSplat.domaincontroller) -ea continue ) ) {
                            } else { Cleanup ; Exit ;}
                        } else {
                            # without the -ea stop, we need an explicit error
                            throw "Unable to resolve $($InputSplat.ManagedBy) to any existing OP or EXO mailbox" ;
                            Cleanup ; Exit ;
                        } ;
                    }
                    default {
                        throw "$($InputSplat.ManagedBy) Not found, or unrecognized RecipientType" ;
                        Cleanup ; Exit ;
                    }
                } ;

            } else {
                # strict shared acct, no FirstName
                # support for shared/room/equip
                if(!$Equip -AND !$Room){
                    # only use Shared when not Equip or Room
                    $smsg= "SHARED Mailbox specified" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                    $InputSplat.Add("Shared",$true ) ;
                    # tear out the unused
                    $InputSplat.Remove("Room") ;
                    $InputSplat.Remove("Equip") ;
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
                } ;
                $InputSplat.Add("FirstName",$null ) ;
                $InputSplat.Add("LastName",$InputSplat.DisplayName) ;
                # psv2 LACKS the foreign-lang cleanup funtions below, skip on psv2
                if($host.version.major -lt 3){
                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):Psv2 detected, skipping foreign character normalization!" ;
                     $DnameClean=$InputSplat.DisplayName
                } else {
                    $DnameClean=Remove-StringDiacritic -string $InputSplat.DisplayName
                    $DnameClean= Remove-StringLatinCharacters -string $DnameClean ;
                } ;
                # strip all nonalphnums from samacct!
                $InputSplat.samaccountname=$([System.Text.RegularExpressions.Regex]::Replace($DnameClean,"[^1-9a-zA-Z_]",""));
                if($InputSplat.samaccountname.length -gt 20) { $InputSplat.samaccountname=$InputSplat.samaccountname.tostring().substring(0,20) };

                # base generics off of baseuser
                # deter BaseUser as a random user in the $($MbxSplat.OrganizationalUnit)
                # leave the -BaseUser param in to force an override, but if blank, draw random from the OU
                if(!($InputSplat.BaseUser)){
                    # draw a random from the $($MbxSplat.OrganizationalUnit)
                    if($InputSplat.SiteCode -eq "($ADSiteCodeUS)"){
                        if ( $InputSplat.BaseUser=get-mailbox -OrganizationalUnit $($MbxSplat.OrganizationalUnit) -resultsize 50 | ?{($_.distinguishedname -notlike '*demo*') -AND (!$_.CustomAttribute5)} | get-random ) {

                        } elseif ( $InputSplat.BaseUser=get-remotemailbox -OnPremisesOrganizationalUnit  $($MbxSplat.OrganizationalUnit) -resultsize 50 | ?{($_.distinguishedname -notlike '*demo*') -AND (!$_.CustomAttribute5)} | get-random ) {

                        } else {
                            $smsg= "UNABLE TO FIND A BASEUSER - USE -BASEUSER TO SPECIFY A SUITABLE ACCT *SOMEWHERE*" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                            Cleanup ; Exit ;
                        } ;
                    } else {
                        if ( $InputSplat.BaseUser = get-remotemailbox -OnPremisesOrganizationalUnit $($MbxSplat.OrganizationalUnit) -resultsize 50 | ? { $_.distinguishedname -notlike '*demo*' } | get-random   ) {

                        }elseif ( $InputSplat.BaseUser=get-mailbox -OrganizationalUnit $($MbxSplat.OrganizationalUnit) -resultsize 50 | ?{$_.distinguishedname -notlike '*demo*'} | get-random
                          ) {
                        } elseif ( $InputSplat.BaseUser=get-remotemailbox -OnPremisesOrganizationalUnit  $FallBackBaseUserOU -resultsize 50 | ?{($_.distinguishedname -notlike '*demo*') -AND (!$_.CustomAttribute5)} | get-random ) {
                            # if all else fails, pull a random remotemailbox from ($ADSiteCodeUS) - we're losing comps moving ahead
                        } else {
                            $smsg= "UNABLE TO FIND A BASEUSER - USE -BASEUSER TO SPECIFY A SUITABLE ACCT *SOMEWHERE*" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                            Cleanup ; Exit ;
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
                                Cleanup ; Exit ;
                            } ;
                        } ;
                        "MailUser" {
                            if ( ($tmpBaseUser=(get-remotemailbox -identity $InputSplat.BaseUser -ea stop)) ){
                                $InputSplat.BaseUser=$tmpBaseUser ;
                            } else {
                                # without the -ea stop, we need an explicit error
                                throw "Unable to resolve $($InputSplat.BaseUser) to any existing OP or EXO mailbox" ;
                                Cleanup ; Exit ;
                            } ;
                        } ;
                        default {
                            throw "$($InputSplat.ManagedBy) Not found, or unrecognized RecipientType" ;
                            Cleanup ; Exit ;
                        }
                    } ;
                } ;

            };
            $InputSplat.Add("Phone","") ;
            # *** BREAKPOINT ;
            $InputSplat.Add("Title","") ;

            #region passwordgen #-----------
            # need to test complex, and if failed, pull another: (above doesn't consistently deliver Ad complexity req's)
            Do { $password = $([System.Web.Security.Membership]::GeneratePassword(8,2)) } Until (Validate-Password -pwd $password ) ;
            $InputSplat.Add("pass",$($password));
            #region passwordgen #-----------

            # 9:43 AM 4/1/2016 secgrp is only necc for Generics
            if(!$InputSplat.NonGeneric){
                $InputSplat.Add("PermGrpName",$($InputSplat.SiteName + "-Data-Email-" + $InputSplat.DisplayName + "-G")) ;
            } else {
                $InputSplat.Add("PermGrpName",$null) ;
            } ;

            if ( $InputSplat.BUserAD=(get-user -identity $($InputSplat.BaseUser.samaccountname) -domaincontroller $($InputSplat.domaincontroller) -ea continue)  ) {
            } else { Cleanup ; Exit ;} ;
            $InputSplat.ADDesc="$(get-date -format 'MM/dd/yyyy') for $($InputSplat.OwnerMbx.samaccountname) $($InputSplat.ticket) -tsk" ;

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


            # 64char limit on names
            if($InputSplat.DisplayName.length -gt 64){
                $smsg= "`n **** NOTE TRUNCATING NAME, -GT 64 CHARS!  ****`N" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                $MbxSplat.Name=$InputSplat.DisplayName.Substring(0,63) ;
                $MbxSplat.DisplayName=$InputSplat.DisplayName.Substring(0,63) ;
            } else {
                $MbxSplat.Name=$InputSplat.DisplayName;
                $MbxSplat.DisplayName=$InputSplat.DisplayName;
            };

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
                $smsg= "`n **** NOTE TRUNCATING FIRSTNAME, -GT 64 CHARS!  ****`N" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                $MbxSplat.LastName=$InputSplat.LastName.Substring(0,63) ;
            } else {
                $MbxSplat.LastName=$InputSplat.LastName;
            };

            if($MbxSplat.FirstName -AND $MbxSplat.LastName){
                # dot-separated addr
                $DirName="$([System.Text.RegularExpressions.Regex]::Replace($MbxSplat.LastName,'[^1-9a-zA-Z_]','')).$([System.Text.RegularExpressions.Regex]::Replace($MbxSplat.FirstName,'[^1-9a-zA-Z_]',''))" ;
            } else {
                # no-dot, just the fname+lname concat'd, trimmed
                $DirName=("$([System.Text.RegularExpressions.Regex]::Replace($MbxSplat.LastName,'[^1-9a-zA-Z_]',''))$([System.Text.RegularExpressions.Regex]::Replace($MbxSplat.FirstName,'[^1-9a-zA-Z_]',''))").trim() ;
            } ;

            switch ($Domain){
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
            $MbxSplat.Alias = $InputSplat.samaccountname;
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
                New-Mailbox @MbxSplat -ea Stop ;

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
                        # do loop until up to 4 retries...
                        Do {
                            Try {
                                New-Mailbox @MbxSplat -ea Stop ;
                                $Exit = $Retries ;
                            } Catch {
                                Start-Sleep -Seconds $RetrySleep ;
                                $Exit ++ ;
                                $smsg= "Failed to exec cmd because: $($Error[0])" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                $smsg= "Try #: $Exit" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                If ($Exit -eq $Retries) {$smsg= "Unable to exec cmd!"; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; } ;
                            } # try-E
                        } Until ($Exit -eq $Retries) # loop-E

                        $MbxSplat.Whatif=$true ;
                    } ;

                    if($Whatif){
                        $smsg= "SKIPPING EXEC: Whatif-only pass";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                    } else {
                        do {Write-Host "." -NoNewLine;Start-Sleep -s 1} until ($oMbx = (get-mailbox -identity  $($Mbxsplat.samaccountname) -domaincontroller $($Mbxsplat.DomainController) -ea silentlycontinue)) ;
                        if($bDebug){
                            $smsg= "`$oMbx:$($Mbxsplat.DisplayName)" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                            $smsg= "`$oMbx.DN:$($oMbx.DistinguishedName)" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        } ;
                    }  # if-E
                } else {
                    $smsg= "INVALID KEY ABORTING NO CHANGE!" ;  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                    Exit ;
                } ;

            } # if-E No oMbx

            if($oMbx){

                if($Whatif){
                        $smsg= "SKIPPING REMAINING AD CMDS - NO OBJECT YET: Whatif-only pass";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                } else {
                    $MbxSetSplat.Whatif=$false ;
                    $Exit = 0 ;
                    if($bDebug) {$smsg= "$((get-date).ToString("HH:mm:ss")):Updating Mbx" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; }
                    Do {
                        Try {
                            Set-Mailbox @MbxSetSplat -ea Stop  ;
                            $Exit = $Retries ;
                        } Catch {
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
                    $Exit = 0 ;
                    Do {
                        Try {
                            set-CASMailbox @MbxSetCASmbx -ea Stop ;
                            $Exit = $Retries ;
                        } Catch {
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
                        set-mailbox $($InputSplat.samaccountname) -domaincontroller $($InputSplat.domaincontroller) -EmailAddressPolicyEnabled $false  ;sleep 1; set-mailbox $($InputSplat.samaccountname) -domaincontroller $($InputSplat.domaincontroller) -EmailAddressPolicyEnabled $true;
                    } ;

                    #=========== V NOTES PARSER
                    <# This block takes an existing Adobject Notes field and parses it and updates fields with a new ChangeLog reference
                    #>
                    # we should _append_ the $InfoStr into any existing Info for the object

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
                        $Exit = 0 ; # zero out $exit each new cmd try/retried
                        Do {
                            Try {
                                Set-ADUser @ADSplat -Replace @{info="$($UpdInfo)"} -ErrorAction stop ;
                                $Exit = $Retries ;
                            } Catch {
                                Start-Sleep -Seconds $RetrySleep ;
                                $Exit ++ ;
                                $smsg= "Failed to exec cmd because: $($Error[0])" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                $smsg= "Try #: $Exit" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                If ($Exit -eq $Retries) {$smsg= "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; } ;
                            } # try-E
                        } Until ($Exit -eq $Retries) # loop-E

                        $ADSplat.whatif = $true ;
                    }  ;

                    #region UPNFromEmail ; # ------
                    # pull the SMTP: addr and use it to force the UPN
                    if($Whatif){
                        $smsg= "Whatif skipping UPN Update...";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                    } else {
                        # dynm pull the forestdom from the forest, match @DOMAIN.COM
                        $forestdom=get-adforest -ea stop | select -expand upnsuffixes |?{$_ -match $rgxTTCDomainsLegacy}
                        if($forestdom -is [string]){
                            # pull primary SMTP:, verify -is [string]/non-array
                            Do {
                                $smsg= "Waiting for ADUser to return email addresses" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                write-host "." -NoNewLine;Start-Sleep -m (1000 * 5)
                                $dirname=(Get-ADUser -identity $oMbx.samaccountname -server $InputSplat.domaincontroller -Properties proxyAddresses -ea 0 | Select -Expand proxyAddresses | Where {$_ -clike "SMTP:*"});
                            } Until (($dirname)) ;

                            if($dirname -is [string]){
                                # convert the $dirname to string and strip proto and take first element
                                $dirname=$dirname.tostring().replace("SMTP:","").split("@")[0]  ;
                                $newUPN="$($dirname)@$($forestdom)";
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
                                $newUPN = $newUPN.replace("&", "").replace("'","").replace("-","").replace("%","").replace("+","")  ;

                                if($bDebug){$smsg= "$((get-date).ToString("HH:mm:ss")):`$newUPN:$($newUPN)"; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; } ;

                                # retry support
                                $Exit = 0 ;
                                Do {
                                    Try {
                                        Set-ADUser -identity $oMbx.samaccountname -UserPrincipalName $newUPN -server $InputSplat.domaincontroller -ErrorAction Stop;
                                        $Exit = $Retries ;
                                    } Catch {
                                        Start-Sleep -Seconds $RetrySleep ;
                                        $Exit ++ ;
                                        $smsg= "Failed to exec cmd because: $($Error[0])" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                        $smsg= "Try #: $($Exit)" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                        If ($Exit -eq $Retries) {$smsg= "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; } ;
                                    } # try-E
                                } Until ($Exit -eq $Retries) # loop-E

                            } else {
                                throw "invalid `$dirname$($dirname) type returned! (non-string)" ;
                            } ;
                        } else {
                            throw "invalid `$forestdom:$($forestdom) returned! (non-string)" ;
                        } ;
                    } ;
                    #endregion UPNFromEmail ; # ------
                    #endregion MAKECHANGES ; # ------

                    #region REPORTING ; # ------
                    do{
                        $smsg= "===REVIEW SETTINGS:=== " ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        $mbxo = get-mailbox -Identity $($InputSplat.samaccountname) -domaincontroller $($InputSplat.domaincontroller) ;
                        $cmbxo= Get-CASMailbox -Identity $($MbxSplat.samaccountname) -domaincontroller $($InputSplat.domaincontroller) ;
                        $aduprops="GivenName,Surname,Manager,Company,Office,Title,StreetAddress,City,StateOrProvince,c,co,countryCode,PostalCode,Phone,Fax,Description" ;
                        $ADu = get-ADuser -Identity $($InputSplat.samaccountname) -properties * -server $($InputSplat.domaincontroller)| select *;
                        $smsg= "User Email:`t$(($mbxo.WindowsEmailAddress.tostring()).trim())" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        $smsg= "Owner Email:`t$(($InputSplat.OwnerMbx.WindowsEmailAddress.tostring()).trim())" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        $smsg= "Mailbox Information:" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        $smsg= "$(($mbxo | select @{Name='LogonName';Expression={$_.SamAccountName }},Name,DisplayName,Alias,database,UserPrincipalName,RetentionPolicy,CustomAttribute5,CustomAttribute9,RecipientType,RecipientTypeDetails | out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                        $smsg= "$(($Adu | select GivenName,Surname,Manager,Company,Office,Title,StreetAddress,City,StateOrProvince,c,co,countryCode,PostalCode,Phone,Fax,Description | out-string).trim())";
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



            }  # if-E $ombx

        } else { $smsg= "OU $(OU) not found. ABORTING!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; } ;



    } # PROC-E ;

    END {


    } # END-E
}

#*------^ new-MailboxShared.ps1 ^------