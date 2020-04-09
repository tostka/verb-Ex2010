#*------v add-MailboxAccessGrant.ps1 v------
function add-MailboxAccessGrant {
    <#
    .SYNOPSIS
    add-MailboxAccessGrant.ps1 - Grant access to a mailbox, via middle-man AD Security Group
    .NOTES
    Version     : 1.0.
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
    # 2:35 PM 4/3/2020 genericized for pub, moved material into infra, updated hybrid mod loads, cleaned up comments/remmed material ; updated to use start-log, debugged to funciton on jumpbox, w divided modules
    # 8:48 AM 11/26/2019 new-MailboxShared():moved the Office spec from $MbxSplat => $MbxSetSplat. New-Mailbox syntax set that supports -Shared, doesn't support -Office
    # 12:10 PM 10/10/2019 default $mbxsplat.Office to $SiteCode value - they should always match, no reason to have a blank offic, if you know the OU & Site. Updated load-ADMS call to specific -cmdlets (speed)
    # 12:51 PM 10/4/2019 passing initial whatif on -Room
    # 2:22 PM 10/1/2019 2076 & 1549 added: $FallBackBaseUserOU, OU that's used when can't find any baseuser for the owner's OU, default to a random shared from SITECODE (avoid crapping out):
    # 9:48 AM 9/27/2019 new-MailboxShared:added `a 'beep' to YYY prompt
    # 1:47 PM 9/20/2019 switched to auto mailbox distribution (ALL TO SITECODE dbs)
    # 10:44 AM 6/18/2019 #2779, rem'd assign repl'd with explicit add
    # 12:52 PM 6/13/2019 repl get-timestamp with expanded cmd in remaining
    # 11:48 AM 6/13/2019 repl function cleanup() with the 10/2/18 vers
    # 11:46 AM 6/13/2019 updated write-log() with deferring verb-transcript.ps1 code
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
    # 10:28 AM 6/27/2018 add $domaincontroller param option - skips dc discovery process and uses the spec, also updated $findOU code to work with torolab dom
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
    # 10:38 AM 6/7/2016 tested debugged 378194, generic creation. Now has new UPN set based on Primary SMTP dirname@toro.com.
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
    add-MailboxAccessGrant.ps1 - Create New Generic Mbx
    .DESCRIPTION
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
    .\new-MailboxGenericTOR.ps1 -ticket "355925" -DisplayName "XXX Confirms"  -MInitial ""  -Owner "LOGON" -NonGeneric $true -showDebug -whatIf ;
    Testing syntax with explicit BaseUSer specified, Whatif test & Debug messages displayed:
    .EXAMPLE
    .\new-MailboxGenericTOR.ps1 -ticket "355925" -DisplayName "XXX Confirms"  -MInitial ""  -Owner "LOGON" -BaseUser "AccountsReceivable" -NonGeneric -showDebug -whatIf ;
    .EXAMPLE
    .\new-MailboxGenericTOR.ps1 -Ticket 99999 -DisplayName "TestScriptMbxRoom" -MInitial "" -Owner LOGON -NonGeneric $false -Room $true -SiteOverride SITECODE -Vscan YES -showDebug -whatif ;
    Testing syntax with Room resource mbx type specified, Whatif test & Debug messages displayed:
    .EXAMPLE
    .\new-MailboxGenericTOR.ps1 -Ticket 99999 -DisplayName "TestScriptMbxEquip" -MInitial "" -Owner LOGON -NonGeneric $false -Equip  $true -SiteOverride SITECODE -Vscan YES -showDebug -whatif ;
    Testing syntax with Equipment resource mbx type specified, Whatif test & Debug messages displayed:

    .LINK
    #>

    # 10:13 AM 10/21/2015 switch IsGeneric from switch to boolean, retitle as NonGeneric, and default to $false unless explicitly $true
    # 1:31 PM 2/27/2017 default vscan true
    Param(
        [Parameter(Mandatory = $true, HelpMessage = "Display Name for mailbox [fname lname,genericname]")]
        [string]$DisplayName,
        [Parameter(HelpMessage = "Middle Initial for mailbox (for non-Generic)[a]")]
        [string]$MInitial,
        [Parameter(Mandatory = $true, HelpMessage = "Specify the userid to be responsible for access-grant-approvals[name,emailaddr,alias]")]
        [string]$Owner,
        [Parameter(HelpMessage = "Optionally a specific existing mailbox upon which to base the new mailbox settings (default is to draw a random mbx from the target OU)[name,emailaddr,alias]")]
        [string]$BaseUser,
        [Parameter(HelpMessage = "Optional parameter indicating new mailbox Is Room-type[-Room `$true]")]
        [bool]$Room,
        [Parameter(HelpMessage = "Optional parameter indicating new mailbox Is Equipment-type[-Equip `$true]")]
        [bool]$Equip,
        [Parameter(HelpMessage = "Optional parameter indicating new mailbox Is NonGeneric-type[-NonGeneric `$true]")]
        [bool]$NonGeneric,
        [Parameter(HelpMessage = "Optional parameter indicating new mailbox belongs to a Contractor[-IsContractor switch]")]
        [switch]$IsContractor,
        [Parameter(HelpMessage = "Optional parameter controlling Vscan (CU9) access (prompts if not specified)[-Vscan YES|NO|NULL]")]
        [string]$Vscan = "YES",
        [Parameter(Mandatory = $false, HelpMessage = "Optionally force CU5 (variant domain assign) [-Cu5 Exmark]")]
        [string]$Cu5,
        [Parameter(HelpMessage = "Optionally specify a 3-letter Site Code o force OU placement to vary from Owner's current site[3-letter Site code]")]
        [string]$SiteOverride,
        [Parameter(Mandatory = $true, HelpMessage = "Incident number for the change request[[int]nnnnnn]")]
        [int]$Ticket,
        [Parameter(HelpMessage = "Option to hardcode a specific DC [-domaincontroller xxxx]")]
        [string]$domaincontroller,
        [Parameter(HelpMessage = "Suppress YYY confirmation prompts [-NoPrompt]")]
        [switch] $NoPrompt,
        [Parameter(HelpMessage = 'Debugging Flag [$switch]')]
        [switch] $showDebug,
        [Parameter(HelpMessage = 'Whatif Flag [$switch]')]
        [switch] $whatIf
    ) ;

    BEGIN {
        $verbose = ($VerbosePreference -eq "Continue") ;
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
            $tModName = $tMod.split(';')[0] ;
            $tModFile = $tMod.split(';')[1] ;
            $tModCmdlet = $tMod.split(';')[2] ;
            $smsg = "( processing `$tModName:$($tModName)`t`$tModFile:$($tModFile)`t`$tModCmdlet:$($tModCmdlet) )" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            if($tModName -eq 'verb-Network' -OR $tModName -eq 'verb-Text' -OR $tModName -eq 'verb-IO'){
                write-host "GOTCHA!:$($tModName)" ;
            } ;
            $lVers = get-module -name $tModName -ListAvailable -ea 0 ;
            if($lVers){
                $lVers=($lVers | sort version)[-1];
                try {
                    import-module -name $tModName -RequiredVersion $lVers.Version.tostring() -force -DisableNameChecking
                }   catch {
                     write-warning "*BROKEN INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;import-module -name $tModDFile -force -DisableNameChecking
                } ;
            } elseif (test-path $tModFile) {
                write-warning "*NO* INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;
                try {import-module -name $tModDFile -force -DisableNameChecking}
                catch {
                    write-error "*FAILED* TO LOAD MODULE*:$($tModName) VIA $(tModFile) !" ;
                    $tModFile = "$($tModName).ps1" ;
                    $sLoad = (join-path -path $LocalInclDir -childpath $tModFile) ;
                    if (Test-Path $sLoad) {
                        Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;
                        . $sLoad ;
                        if ($showdebug) { Write-Verbose -verbose "Post $sLoad" };
                    } else {
                        $sLoad = (join-path -path $backInclDir -childpath $tModFile) ;
                        if (Test-Path $sLoad) {
                            Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;
                            . $sLoad ;
                            if ($showdebug) { Write-Verbose -verbose "Post $sLoad" };
                        } else {
                            Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;
                            exit;
                        } ;
                    } ;
                } ;
            } ;
            if(!(test-path function:$tModCmdlet)){
                write-warning -verbose:$true  "UNABLE TO VALIDATE PRESENCE OF $tModCmdlet`nfailing through to `$backInclDir .ps1 version" ;
                $sLoad = (join-path -path $backInclDir -childpath "$($tModName).ps1") ;
                if (Test-Path $sLoad) {
                    Write-Verbose -verbose:$true ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;
                    . $sLoad ;
                    if ($showdebug) { Write-Verbose -verbose "Post $sLoad" };
                    if(!(test-path function:$tModCmdlet)){
                        write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO CONFIRM `$tModCmdlet:$($tModCmdlet) FOR $($tModName)" ;
                    } else {
                        write-verbose -verbose:$true  "(confirmed $tModName loaded: $tModCmdlet present)"
                    }
                } else {
                    Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;
                    exit;
                } ;
            } else {
                write-verbose -verbose:$true  "(confirmed $tModName loaded: $tModCmdlet present)"
            } ;
        } ;  # loop-E
        #*------^ END MOD LOADS ^------

        #  add password generator
        [Reflection.Assembly]::LoadWithPartialName("System.Web") | out-null ;

        if($ParentPath){
            $logspec = start-Log -Path ($ParentPath) -showdebug:$($showdebug) -whatif:$($whatif) ;
            if($logspec){
                $logging=$logspec.logging ;
                $logfile=$logspec.logfile ;
                $transcript=$logspec.transcript ;
            } else {throw "Unable to configure logging!" } ;
        } else {

        } ;

        <#$transcript = join-path -path $PSScriptRoot -ChildPath "logs" ;
        if(!(test-path -path $transcript)){ "Creating missing log dir $($transcript)..." ; mkdir $transcript  ; } ;
        $transcript=join-path -path $transcript -childpath $ScriptNameNoExt  ;
        $transcript+= "-Transcript-BATCH-$(get-date -format 'yyyyMMdd-HHmmtt')-trans-log.txt"  ;
        #>
        # add log file variant as target of Write-Log:
        $logfile=$transcript.replace("-Transcript","-LOG").replace("-trans-log","-log")
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
        $logging = $True ;

        $xxx="====VERB====";
        $xxx=$xxx.replace("VERB","NewMbx") ;
        $BARS=("="*10);

        $reqMods+="Add-PSTitleBar;Remove-PSTitleBar".split(";") ;
        $reqMods+="Test-TranscriptionSupported;Test-Transcribing;Stop-TranscriptLog;Start-IseTranscript;Start-TranscriptLog;get-ArchivePath;Archive-Log;Start-TranscriptLog".split(";") ;
        $reqMods=$reqMods| select -Unique ;

        #region SPLATDEFS ; # ------

        if (($host.version.major) -lt 3) {
            $InputSplat = @{
                TargetID     = "TARGETMBX";
                SecGrpName   = "";
                Owner        = "LYNCTEST2"
                PermsDays    = 60;
                SiteOverride = "";
                Members      = "LYNCTEST3"
            } ;
            # 2:02 PM 5/10/2016 owner isn't a param on NEW-ADGroup => ManagedBy
            # 7:28 AM 9/6/2018 switch to EXO-compatible group type: Univ, mail-enable
            $SGSplat = @{
                Name            = "";
                DisplayName     = "";
                SamAccountName  = "";
                GroupScope      = "Global";
                GroupCategory   = "Universal";
                ManagedBy       = "";
                Description     = "";
                OtherAttributes = "";
                Path            = "";
                Server          = ""
            };
            $SGUpdtSplat = @{
                Identity = "";
                Server   = ""
            };
            $DGEnableSplat = @{
                Identity         = "";
                DomainController = "" ;
            } ;
            $DGUpdtSplat = @{
                Identity                      = "";
                HiddenFromAddressListsEnabled = $true ;
                DomainController              = "" ;
            } ;
            $GrantSplat = @{
                Identity        = "" ;
                User            = "" ;
                AccessRights    = "FullAccess";
                InheritanceType = "All";
            };
            # 8:05 AM 10/14/2015 add for AD SendAs perms grant
            <#$ADMbxGrantSplat=@{
	          Identity="" ;
	          User="" ;
	          ExtendedRights="Send As" ;
	        };#>
            #8:59 AM 10/14/2015 try pulling id, pipeline it in
            $ADMbxGrantSplat = @{
                User           = "" ;
                ExtendedRights = "Send As" ;
            };
        } else {
            # 12:04 PM 2/10/2016 psv3 code
            $InputSplat = [ordered]@{
                TargetID     = "TARGETMBX";
                SecGrpName   = "";
                Owner        = "LYNCTEST2"
                PermsDays    = 60;
                SiteOverride = "";
                Members      = "LYNCTEST3"
            } ;
            $SGSplat = [ordered]@{
                Name            = "";
                DisplayName     = "";
                SamAccountName  = "";
                GroupScope      = "Universal";
                GroupCategory   = "Security";
                ManagedBy       = "";
                Description     = "";
                OtherAttributes = "";
                Path            = "";
                Server          = ""
            };
            $SGUpdtSplat = [ordered]@{
                Identity = "";
                Server   = ""
            };
            $DGEnableSplat = [ordered]@{
                Identity         = "";
                DomainController = ""
            };
            $DGUpdtSplat = [ordered]@{
                Identity                      = "";
                HiddenFromAddressListsEnabled = $true ;
                DomainController              = "" ;
            } ;
            $GrantSplat = [ordered]@{
                Identity        = "" ;
                User            = "" ;
                AccessRights    = "FullAccess";
                InheritanceType = "All";
            };
            # 8:05 AM 10/14/2015 add for AD SendAs perms grant
            <#$ADMbxGrantSplat=[ordered]@{
	          Identity="" ;
	          User="" ;
	          ExtendedRights="Send As" ;
	        };#>
            #8:59 AM 10/14/2015 try pulling id, pipeline it in
            $ADMbxGrantSplat = [ordered]@{
                User           = "" ;
                ExtendedRights = "Send As" ;
            };
        }

        if ($PSCommandPath) { $pltInput.add("ParentPath", $PSCommandPath) } ;
        if ($TargetID) { $InputSplat.TargetID = $TargetID };
        if ($SecGrpName) { $InputSplat.SecGrpName = $SecGrpName };
        if ($Owner) { $InputSplat.Owner = $Owner };
        if ($PermsDays) { $InputSplat.PermsDays = $PermsDays };
        if ($SiteOverride) { $InputSplat.SiteOverride = $SiteOverride };
        if ($Members) { $InputSplat.Members = $Members };

        $smsg = "`nSpecified Target Email: $($InputSplat.TargetID)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

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
            write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):LOCAL EMS detected" ;
            $Global:E10IsDehydrated=$false ;
        # REMS detect dleTimeout -eq -1
        } elseif(get-pssession |?{$_.configurationname -eq 'Microsoft.Exchange' -AND $_.ComputerName -match $rgxEx10HostName -AND ($_.IdleTimeout -eq -1)} ){
            write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):REMOTE EMS detected" ;
            $reqMods+="get-GCFast;Get-ExchangeServerInSite;connect-Ex2010;Reconnect-Ex2010;Disconnect-Ex2010;Disconnect-PssBroken".split(";") ;
            if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
            reconnect-ex2010 ;
            $Global:E10IsDehydrated=$true ;
        } else {
            write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):No existing Ex2010 Connection detected" ;
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
        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):(loading ADMS...)" ;
        load-ADMS -cmdlet get-aduser,Set-ADUser,Get-ADGroupMember,Get-ADDomainController,Get-ADObject,get-adforest | out-null ;

        $AdminInits=get-AdminInitials ;

        #region LOADMODS ; # ------

    }  # BEG-E ;

    PROCESS {

        #region DATAPREP ; # ------
        $Tmbx = (get-mailbox $($InputSplat.TargetID) -domaincontroller (Get-ADDomainController).Name.tostring() -ea stop) ;
        $GrantSplat.Identity = $($Tmbx.samaccountname);
        $domain = $Tmbx.identity.tostring().split("/")[0]
        $InputSplat.Add("Domain", $($domain) ) ;
        if (!$domaincontroller) {
            $domaincontroller = (get-gcfast -domain $domain) ;
        } else {
            $smsg = "Using hard-coded domaincontroller:$($domaincontroller)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
        } ;

        $InputSplat.Add("DomainController", $domaincontroller) ;
        $SGUpdtSplat.Server = $($InputSplat.DomainController);
        $DGEnableSplat.DomainController = $($domaincontroller);
        $DGUpdtSplat.DomainController = $($domaincontroller);
        $InputSplat.Site = ($Tmbx.identity.tostring().split('/')[1]) ;

        switch ((get-recipient -Identity $Inputsplat.Owner).RecipientType ) {
            "UserMailbox" {
                #$Inputsplat.OwnerMbx = get-mailbox -identity $Inputsplat.Owner -ea stop | select -expand SamAccountname ;
                # 1:15 PM 11/15/2017 no, this needs to be the obj
                #$Inputsplat.OwnerMbx = get-mailbox -identity $Inputsplat.Owner -ea stop  ;

                if ( ($InputSplat.OwnerMbx = (get-mailbox -identity $($InputSplat.Owner) -ea stop)) ) {
                    if ($showdebug) { $smsg = "UserMailbox detected" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; } ;
                } else {
                    # 11:54 AM 11/15/2017 without the -ea stop, we need an explicit error
                    throw "Unable to resolve $($InputSplat.Owner) to any existing OP or EXO mailbox" ;
                    Cleanup ; Exit ;
                } ;
            }
            "MailUser" {
                if ( ($InputSplat.OwnerMbx = (get-remotemailbox -identity $($InputSplat.Owner) -ea stop)) ) {
                    if ($showdebug) {
                        $smsg = "MailUser detected" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; #Error|Warn|Debug
                    } ;
                } else {
                    # 11:54 AM 11/15/2017 without the -ea stop, we need an explicit error
                    throw "Unable to resolve $($InputSplat.Owner) to any existing OP or EXO mailbox" ;
                    Cleanup ; Exit ;
                } ;
            }
            default {
                throw "$($InputSplat.Owner) Not found, or unrecognized RecipientType" ;
                Cleanup ; Exit ;
            }
        } ;

        # owner needs to be samaccountname or DN - can't use email addresses!
        if ($Inputsplat.Owner -match $rgxEmailAddr) {
            $smsg = "Converting Owner email:$($Inputsplat.Owner) to logon..." ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
            #$Inputsplat.Owner = get-mailbox -identity $Inputsplat.Owner -ea stop | select -expand SamAccountname ;
            $Inputsplat.Owner = $InputSplat.OwnerMbx.alias ;
        } ;

        # if no permsdays, default it to 60d
        if (!($InputSplat.PermsDays)) { "defaulting PermsDays to 60"; $InputSplat.PermsDays = 60 };
        if ($InputSplat.PermsDays -eq 999) {
            [string]$PermsExp = (get-date "12/31/2099" -format "MM/dd/yyyy") ;
        } else {
            [string]$PermsExp = (get-date (Get-Date).AddDays($InputSplat.PermsDays + 1) -format "MM/dd/yyyy") ;
        } ;

        $Infostr = "TargetMbx:$($Tmbx.samaccountname)`r`nPermsExpire:$($PermsExp)`r`nIncident:$($Ticket)`r`nAdmin:$($AdminInits)`r`nBusinessOwner:$($InputSplat.Owner);`r`nITOwner:$($InputSplat.Owner)" ;

        $smsg = "Site:$($InputSplat.Site)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
        $smsg = "`nTLogon: $($Tmbx.samaccountname )" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

        $Adu = (get-aduser $($tmbx.SamAccountName) -server $($InputSplat.DomainController) -ea stop -properties manager)  ;
        if ($Adu.Manager) {
            $Mgr = ((get-aduser ($Adu.manager)).samaccountname) ;
            $smsg = "MgrLogon: $($Mgr)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
        } else {
            $smsg = "$($Tmbx.displayname) has a blank AD Manager field.`nAsserting Owner from inputs:$($InputSplat.Owner) " ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
            $Mgr = $($InputSplat.Owner);
        } ;
        # 10:52 AM 6/20/2016 fixed typo $InputSplatSiteOverride => $InputSplat.SiteOverride (broke -SiteOverride function)
        if ($InputSplat.SiteOverride) {
            $SiteCode = $($InputSplat.SiteOverride);
        } else {
            # 12:28 PM 11/15/2017 we need to use the OwnerMbx - Owner currently is the alias, we want the object with it's dn
            $SiteCode = $InputSplat.OwnerMbx.identity.tostring().split("/")[1]  ;
        } ;
        $Domain = "global.ad.toro.com" ;
        if ($env:USERDOMAIN -eq 'TORO') {
            $FindOU = "^OU=Email\sAccess,OU=SEC\sGroups,OU=Managed\sGroups,";
        } ELSEif ($env:USERDOMAIN -eq 'TORO-LAB') {
            # CN=Lab-SEC-Email-Thomas Jefferson,OU=Email Access,OU=SEC Groups,OU=Managed Groups,OU=LYN,DC=global,DC=ad,DC=torolab,DC=com
            $FindOU = "^OU=Email\sAccess,OU=SEC\sGroups,OU=Managed\sGroups,"; ;
        } else {
            throw "UNRECOGNIZED USERDOMAIN:$($env:USERDOMAIN)" ;
        } ;

        $SGSplat.DisplayName = "$($SiteCode)-SEC-Email-$($Tmbx.DisplayName)-G";

        TRY {
            $OU = (Get-ADObject -filter { ObjectClass -eq 'organizationalunit' } -server $($DomainController) | ? { $_.distinguishedname -match "^$($FindOU).*OU=$($SiteCode),.*,DC=ad,DC=toro((lab)*),DC=com$" } | select distinguishedname).distinguishedname.tostring() ;
        } CATCH {
            $ErrTrpd = $_ ;
            $smsg = "UNABLE TO LOCATE $($FindOU) BELOW SITECODE $($SiteCode)!. EXITING!" ; $smsg = "MESSAGE" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; #Error|Warn
            $smsg = "Failed processing $($ErrTrpd.Exception.ItemName). `nError Message: $($ErrTrpd.Exception.Message)`nError Details: $($ErrTrpd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; #Error|Warn
            Exit
        } ;
        If ($OU -isnot [string]) {
            $smsg = "WARNING AD OU SEARCH SITE:$($InputSplat.SiteCode), FindOU:$($FindOU), FAILED TO RETURN A SINGLE OU...`n$($OU.distinguishedname)`nEXITING!";
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; #Error|Warn
            Exit ;
        } ;

        $smsg = "$SiteCode" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
        $smsg = "$OU" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

        $SGSplat.Path = $OU ;
        $smsg = "Checking specified SecGrp Members..." ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
        $SGMembers = ($InputSplat.members.split(",") | foreach { get-recipient $_ -ea stop })
        $smsg = "Checking for existing $($SGSplat.DisplayName)..." ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

        if ($bDebug) {
            $smsg = "`$SGSrchName:$($SGSrchName)`n`$SGSplat.DisplayName: $($SGSplat.DisplayName)"; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ;
        } ;

        $SGSrchName = $($SGSplat.DisplayName);
        $oSG = Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -server $($InputSplat.DomainController) -ErrorAction stop;

        if ($oSG) {
            if ($bDebug) {
                $smsg = "`$oSG:$($oSG.SamAccountname)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; #Error|Warn
                $smsg = "`$oSG.DN:$($oSG.DistinguishedName)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; #Error|Warn
            } ;
            # 2:38 PM 3/21/2016 we should _append_ the $InfoStr into any existing Info for the object
            # 11:58 AM 2/27/2017 can't use [ordered] on psv2 if we must have these in order use a psv2 OrderedDictionary
            if (($host.version.major) -lt 3) {
                <#
                $ADOtherInfoProps=@{
                TargetMbx=$null ;
                PermsExpire=$null ;
                Incident=$null ;
                Admin=$null ;
                BusinessOwner=$null ;
                ITOwner=$null ;
                } ;
                #>
                $ADOtherInfoProps = New-Object Collections.Specialized.OrderedDictionary ;
                #$ADOtherInfoProps.Add('One',1) ;
                $ADOtherInfoProps.Add('TargetMbx', $null) ;
                $ADOtherInfoProps.Add('PermsExpire', $null) ;
                $ADOtherInfoProps.Add('Incident', $null) ;
                $ADOtherInfoProps.Add('Admin', $null) ;
                $ADOtherInfoProps.Add('BusinessOwner', $null) ;
                $ADOtherInfoProps.Add('ITOwner', $null) ;

            } else {
                $ADOtherInfoProps = [ordered]@{
                    TargetMbx     = $null ;
                    PermsExpire   = $null ;
                    Incident      = $null ;
                    Admin         = $null ;
                    BusinessOwner = $null ;
                    ITOwner       = $null ;
                } ;
            } ;
            #$Infostr="TargetMbx:$($Tmbx.samaccountname)`r`nPermsExpire:$($PermsExp)`r`nIncident:$($Ticket)`r`nAdmin:$($AdminInits)`r`nBusinessOwner:$($InputSplat.Owner);`r`nITOwner:$($InputSplat.Owner)" ;
            if ($oSG.info) {
                # existing info tag
                # update the splat
                # 12:19 PM 3/22/2016 just loop each line split on `n: (Get-ADUser lynctest9 -Properties info).info.split("`n")| foreach{"Ln:$_"}
                $oADOtherInfo = New-Object PSObject -Property $ADOtherInfoProps ;

                #( $ln in ($oSG.info.tostring().split("`n") )  {
                $ilines = $oSG.info.tostring().split("`n").count ;
                $iIter = 0 ;
                foreach ( $ln in $oSG.info.tostring().split("`n") ) {
                    $iIter++

                    if ($iIter -eq 1) { $UpdInfo = $null; } ;

                    if ($ln -match "^(TargetMbx|PermsExpire|Incident|Admin|BusinessOwner|ITOwner):.*$") {
                        # it's part of a defined Info tag
                        $matches = $null ;
                        # ingest the matches and throw away the lines
                        if ($ln -match "(?<=TargetMbx:)\w+" ) { $oADOtherInfo.TargetMbx = $matches[0] } ; $matches = $null ;
                        if ($ln -match "(?<=PermsExpire:)\d+\/\d+/\d+" ) { $oADOtherInfo.PermsExpire = (get-date $matches[0]) ; } ; ; $matches = $null ;
                        # 12:44 PM 10/18/2016 update rgx for ticket to accommodate 5-digit (or 6) CW numbers "^\d{6}$"=>^\d{5,6}$
                        if ($ln -match "(?<=Incident:)^\d{5,6}$") { $oADOtherInfo.Incident = $matches[0] ; } ; $matches = $null ;
                        if ($ln -match "(?<=Admin:)\w*") { $oADOtherInfo.Admin = $matches[0] ; } ; $matches = $null ;
                        if ($ln -match "(?<=BusinessOwner:)\w{2,20}") { $oADOtherInfo.BusinessOwner = $matches[0] ; } ; $matches = $null ;
                        if ($ln -match "(?<=ITOwner:)\w{2,20}") { $oADOtherInfo.ITOwner = $matches[0] ; } ; $matches = $null ;
                    } else {
                        $UpdInfo += "$($ln)`r`n" ;
                    } ;

                    if ($iIter -eq $iLines) {
                        $smsg = "`$uinfo:`n$uinfo" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                        if ($oADOtherInfo) {
                            $smsg = "Updating existing Info tag:`n$(($oADOtherInfo |out-string).trim())";
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                        }
                        $UpdInfo += "`r`nTargetMbx:$($Tmbx.samaccountname)`r`nPermsExpire:$($PermsExp)`r`nIncident:$($Ticket)`r`nAdmin:$($AdminInits)`r`nBusinessOwner:$($InputSplat.Owner);`r`nITOwner:$($InputSplat.Owner)" ;
                        if ($bDebug) {
                            $smsg = "New Info field:`n$(($UpdInfo |out-string).trim())";
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                        } ;

                        #Set-ADUser -identity $tusr -Replace @{info="$($uinfo)"} -server LYNMS811 -whatif  ;
                        Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -server $($InputSplat.DomainController) -ErrorAction stop | Set-ADUser -Replace @{info = "$($UpdInfo)" }  -whatif ; ;
                    }


                } # loop-E $lines

            } ; # if-E $osg

        } else {
            $smsg = "$($SGSplat.DisplayName) Not found. Testing Create with the following paraemters..."  ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
            # create the secgrp
            $SGSplat.Name = $($SGSplat.DisplayName);
            $SGSplat.SamAccountName = $($SGSplat.DisplayName);
            $SGSplat.ManagedBy = $($InputSplat.Owner);
            $SGSplat.Description = "Email - access to $($Tmbx.displayname)'s mailbox";
            $SGSplat.Server = $($InputSplat.DomainController) ;
            # build the Notes/Info field as a hashcode: OtherAttributes=@{    info="TargetMbx:kadrits`r`nPermsExpire:6/19/2015"  } ;
            $SGSplat.OtherAttributes = @{info = $($Infostr) } ;


            $smsg = "`$SGSplat:`n---"; $smsg = "MESSAGE" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
            foreach ($row in $SGSplat) {
                foreach ($key in $row.keys) {
                    if ($key -eq "OtherAttributes") {
                        $smsg = "==v OtherAttributes: v==" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                        $SGSplat.OtherAttributes.GetEnumerator() | Foreach-Object {
                            $smsg = "==$($_.Key ):==`n$(($_.Value|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                        } ;
                        $smsg = "==^ OtherAttributes: ^==" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                    } else {
                        $smsg = "$($key): $($row[$key])" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                    } ;
                }
            } ;

            $smsg = "---`nWhatif $($SGSplat.Name) creation...";
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

            New-AdGroup @SGSplat -whatif -ea stop;
            $DGEnableSplat.identity = $SGSplat.SamAccountName ;
            $DGUpdtSplat.identity = $SGSplat.SamAccountName ;

            $smsg = "`$DGEnableSplat:`n---`n$(($DGEnableSplat|out-string).trim())`n---`n" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

            write-host -foregroundcolor yellow "$((get-date).ToString("HH:mm:ss")):Continue with $($SGSplat.Name) creation?...";
            if ($NoPrompt) { $bRet = "YYY" } else { $bRet = Read-Host "Enter YYY to continue. Anything else will exit`a" ; } ;
            if ($bRet.ToUpper() -eq "YYY") {


                if ($bWhatif) {
                    $smsg = "-Whatif pass, skipping exec." ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                } else {
                    $smsg = "Executing...";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                    New-AdGroup @SGSplat -ea stop ;
                    Do { write-host "." -NoNewLine; Start-Sleep -s 1 } Until (Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -server $($InputSplat.DomainController)) ;
                    #$oSG= (get-adgroup "$($SGSplat.DisplayName)" -server $($InputSplat.Domain) -ea stop );
                    $oSG = Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -server $($InputSplat.DomainController) -ErrorAction stop;
                    if ($bDebug) {
                        $smsg = "`$oSG:$($oSG.SamAccountname)`n`$oSG.DN:$($oSG.DistinguishedName)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; #Error|Warn
                    } ;
                    $smsg = "Enable-DistributionGroup w`n$(($DGEnableSplat|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    Enable-DistributionGroup @DGEnableSplat ;
                    $smsg = "Set HiddenFromAddressListsEnabled:Set-DistributionGroup w`n$(($DGUpdtSplat|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    Set-DistributionGroup @DGUpdtSplat ;
                    $oSG = Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -prop * -server $($InputSplat.DomainController) -ErrorAction stop;
                    $smsg = "Final SecGrp Config:$($oSG.SamAccountname)`n:$(($oSG | fl Name,GroupCategory,GroupScope,msExchRecipientDisplayType,showInAddressBook,mail|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                } ;
            } else { $smsg = "INVALID KEY ABORTING NO CHANGE!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; Exit ; } ;
        } # if-E $osg

        $smsg = "`nTesting SecGrp Members Add `nto group: $($oSG.Name)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
        # *** BREAKPOINT ;
        if ($oSG) {
            $ExistMbrs = @() ;
            # 11:27 AM 6/23/2017 typo, vari with no leading $
            $oSG | Get-ADGroupMember -server $($DomainController) | Select-Object -ExpandProperty sAMAccountName | ForEach-Object { $ExistMbrs += $_ } ;
            $SGUpdtSplat.Identity = $($oSG.samaccountname) ;
            $DGEnableSplat.Identity = $($oSG.samaccountname) ;
            $DGUpdtSplat.Identity = $($oSG.samaccountname) ;
            $GrantSplat.User = $($oSG.SamAccountName);
            #8:41 AM 10/14/2015 add adp
            $ADMbxGrantSplat.User = $($oSG.SamAccountName);
            $SGUpdtSplat.Server = $($InputSplat.DomainController) ;
            $DGEnableSplat.DomainController = $($InputSplat.DomainController) ;
            $DGUpdtSplat.DomainController = $($InputSplat.DomainController) ;
            # 12:47 PM 10/6/2015 add dc
            $GrantSplat.Add("DomainController", $($InputSplat.domaincontroller)) ;
            #8:41 AM 10/14/2015 add adp
            $ADMbxGrantSplat.Add("DomainController", $($InputSplat.domaincontroller)) ;

            if ($bWhatif) {
                $smsg = "-Whatif pass, skipping exec." ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
            } else {
                foreach ($Mbr in $SGMembers) {
                    If ($ExistMbrs -notcontains $Mbr.sAMAccountName) {
                        $smsg = "Test ADD:$($mbr.samaccountname)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                        Add-ADGroupMember @SGUpdtSplat -members $($mbr.samaccountname) -ea stop -whatif ;
                        <# 8:27 AM 6/27/2019 after win10ug, the above is now throwing:
                        ...
                        Testing SecGrp Member Add
                        to group: LYN-SEC-Email-wpaIRR-G
                        08:21:12:Test ADD:antoidx
                        add-MbxAccessGrant : Parameter cannot be processed because the parameter name 'member' is ambiguous. Possible matches include: -Members -MemberTimeToLive.
                        At C:\usr\work\exch\scripts\add-MbxAccessGrant.ps1:2510 char:1
                        + add-MbxAccessGrant @pltInput
                        Looks like they've added a new param, and it's broken partials,
                        curr doc: https://docs.microsoft.com/en-us/powershell/module/addsadministration/add-adgroupmember?view=win10-ps
                        Add-ADGroupMember  [-WhatIf]  [-Confirm]  [-AuthType <ADAuthType>]  [-Credential <PSCredential>]  [-Identity] <ADGroup>  [-Members] <ADPrincipal[]>  [-MemberTimeToLive <TimeSpan>]  [-Partition <String>]  [-PassThru]  [-Server <String>]  [<CommonParameters>]
                        # yea the param is named -members, not -member, with the addition it broke the auto-resolution
                        #>
                    } else {
                        $smsg = "SKIPPING:$($mbr.samaccountname) is already a member of $($oSG.samaccountname)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    } ;
                }  # loop-E ;
                $smsg = "Continue with Member Addition?...";
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                if ($NoPrompt) { $bRet = "YYY" } else { $bRet = Read-Host "Enter YYY to continue. Anything else will exit`a" ; } ;
                if ($bRet.ToUpper() -eq "YYY") {
                    $smsg = "Exec Update";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    foreach ($Mbr in $SGMembers) {
                        If ($ExistMbrs -notcontains $Mbr.sAMAccountName) {
                            "Exec ADD:$($mbr.samaccountname)"
                            if ($whatif) {
                                # 11:17 AM 6/22/2015 whatif-only pass
                                $smsg = "SKIPPING EXEC: Whatif-only pass";
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                            } else {
                                # 8:33 AM 6/27/2019 fix latest ADmod, added a conflicting param, autoresolve fails, typo -member -> proper -members
                                Add-ADGroupMember @SGUpdtSplat -members $($mbr.samaccountname) -ea stop ;
                            } ;
                        } else {
                            "SKIPPING:$($mbr.samaccountname) is already a member of $($oSG.samaccountname)"
                        } ;
                    } #  # loop-E;
                } ;
            } # if-E whatif ;
            $mbxp = $Tmbx | get-mailboxpermission -user ($oSG).Name -domaincontroller $InputSplat.domaincontroller -ea silentlycontinue | ? { $_.user -match ".*-(SEC|Data)-Email-.*$" }
            $smsg = "`nChecking Mailbox Permission on $($Tmbx.samaccountname) mailbox to accessing user:`n $($oSG.Name)...`n(blank if none)`n---`n$(($mbxp | select user,AccessRights,IsInhertied,Deny | format-list|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug


            # 8:10 AM 10/14/2015 AD SendAs too

            $mbxadp = $Tmbx | Get-ADPermission -domaincontroller $($InputSplat.domaincontroller) -ea Silentlycontinue | where { ($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and ($_.user -match ".*-(SEC|Data)-Email-.*$") };

            $smsg = "`nChecking AD SendAs Permission on $($Tmbx.samaccountname) mailbox to accessing user:`n $($oSG.Name)...`n(blank if none)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

            $smsg = "`n$(($mbxadp | select identity,User,ExtendedRights,Deny,Inherited | format-list|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
            # format-table -wrap ;

            $smsg = "`n---`nExisting $($oSG.Name) Membership...`n(blank if none)`n$((Get-ADGroupMember -identity $oSG.samaccountname -server $($DomainController) | select distinguishedName|out-string).trim())`n" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

            $smsg = "Testing Permissions Grant Update...`nAdd-MailboxPermission -whatif w`n$(($GrantSplat|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

            # 2:53 PM 5/18/2016 add retry code:
            $Exit = 0 ;
            # do loop until up to 4 retries...
            Do {
                Try {

                    add-mailboxpermission @GrantSplat -whatif ;
                    $Exit = $Retries ;
                } Catch {
                    $ErrTrapd = $Error[0] ;

                    Start-Sleep -Seconds $RetrySleep ;
                    $Exit ++ ;
                    $smsg = "Failed to exec add-mailboxpermission -whatif cmd because: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    $smsg = "Try #: $($Exit)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    If ($Exit -eq $Retries) { $smsg = "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ; } ;
                    # 11:15 AM 11/26/2019 add Cont - doesn't seem to be retrying
                    Continue ;
                } # try-E
            } Until ($Exit -eq $Retries) # loop-E

            $smsg = "Add-ADPermission -whatif... w`n$(($ADMbxGrantSplat|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

            $Exit = 0 ;
            Do {
                Try {
                    add-adpermission -identity $($TMbx.Identity) @ADMbxGrantSplat -whatif ;
                    $Exit = $Retries ;
                } Catch {
                    $ErrTrapd = $Error[0] ;
                    Start-Sleep -Seconds $RetrySleep ;
                    $Exit ++ ;
                    $smsg = "Failed to exec add-mailboxpermission -whatif cmd because: $($ErrTrpd)`nTry #: $($Exit)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ;
                    If ($Exit -eq $Retries) { $smsg = "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ; } ;
                    # 11:15 AM 11/26/2019 add Cont - doesn't seem to be retrying
                    Continue ;
                } # try-E
            } Until ($Exit -eq $Retries) # loop-E

            write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):Exec Permissions Grant Update";
            if ($whatif) {
                # 11:17 AM 6/22/2015 whatif-only pass
                write-verbose -verbose:$true "SKIPPING EXEC: Whatif-only pass";
            } else {
                write-host -foregroundcolor red "$((get-date).ToString("HH:mm:ss")):EXEC Add-MailboxPermission...";
                #add-mailboxpermission @GrantSplat ;

                $Exit = 0 ;
                # do loop until up to 4 retries...
                Do {
                    Try {

                        add-mailboxpermission @GrantSplat ;

                        $Exit = $Retries ;
                    } Catch {
                        $ErrTrapd = $Error[0] ;

                        Start-Sleep -Seconds $RetrySleep ;
                        $Exit ++ ;
                        $smsg = "Failed to exec add-mailboxpermission EXEC cmd because: $($ErrTrapd)`nTry #: $($Exit)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ; #Error|Warn|Debug
                        If ($Exit -eq $Retries) { $smsg = "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; } ;
                        # 11:15 AM 11/26/2019 add Cont - doesn't seem to be retrying
                        Continue ;
                    } # try-E
                } Until ($Exit -eq $Retries) # loop-E

                $smsg = "Add-ADPermission -whatif:identity $($TMbx.Identity) w`n$(($ADMbxGrantSplat|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                $Exit = 0 ;
                Do {
                    Try {
                        add-adpermission -identity $($TMbx.Identity) @ADMbxGrantSplat ;
                        $Exit = $Retries ;
                    } Catch {
                        $ErrTrapd = $Error[0] ;

                        Start-Sleep -Seconds $RetrySleep ;
                        $Exit ++ ;
                        $smsg = "Failed to exec add-adpermission EXEC cmd because: $($ErrTrapd)`nTry #: $($Exit)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ; #Error|Warn|Debug
                        If ($Exit -eq $Retries) { $smsg = "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; } ;
                        # 11:15 AM 11/26/2019 add Cont - doesn't seem to be retrying
                        Continue ;
                    } # try-E
                } Until ($Exit -eq $Retries) # loop-E

                # generics don't need this, test the OU path and only add folks below users
                # we're only hiding folks matching:
                #$rgxBannedOUs="^.*,OU=Disabled,OU=Users,.*OU=\w*,DC=(global|china),DC=ad,DC=toro,DC=com$" ;
                # and unhiding folks matching
                if ($Tmbx.distinguishedname -match $rgxUserOUs) {
                    # block that adds the $tmbx to the maintain-offboards.ps1 target AccGrant group for the region
                    $smsg = "Add TMBX $($tMbx.samaccountname) to AccGrant Group`n$(($TMbx | select -expand distinguishedname |?{$_ -match "DC=((global|china)),DC=ad,DC=toro,DC=com"}|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    switch ($matches[0]) {
                        "DC=global,DC=ad,DC=toro,DC=com" { $grpN = "LYN-DL-Exch-AGUnHide" ; } ;
                        "DC=china,DC=ad,DC=toro,DC=com" { $grpN = "XIA-DL-Exch-AGUnHide" ; } ;
                        default {
                            $smsg = "domain:NO MATCH!"; $smsg = "MESSAGE" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ;
                            EXIT ;
                        } ;
                    } ; # switch-E
                    $smsg = "==TGroup:$($grpN)";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                    if ($tdl = get-DistributionGroup -identity $grpN -domaincontroller $($InputSplat.domaincontroller) ) {
                        "==Add $($TMbx.name) to $($tdl.alias):" ;

                        $Exit = 0 ;
                        # do loop until up to 4 retries...
                        Do {
                            Try {
                                add-DistributionGroupMember -identity $tdl.alias -Member $TMbx.distinguishedname -domaincontroller $($InputSplat.domaincontroller) -whatif:$($whatif) ;

                                $Exit = $Retries ;
                            } Catch {
                                $ErrTrapd = $Error[0] ;

                                Start-Sleep -Seconds $RetrySleep ;
                                $Exit ++ ;
                                $smsg = "Failed to exec add-DistributionGroupMember EXEC cmd because: $($ErrTrapd)`nTry #: $($Exit)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ; #Error|Warn|Debug
                                If ($Exit -eq $Retries) { $smsg = "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; } ;
                                # 11:15 AM 11/26/2019 add Cont - doesn't seem to be retrying
                                Continue ;
                            } # try-E
                        } Until ($Exit -eq $Retries) # loop-E

                    } else {
                        "$($grpN): NOT FOUND" ;
                    }  ;
                } else {
                    $smsg = "TMBX $($tMbx.samaccountname) is in a non-User OU: Term Hide/Unhide groups do not apply...";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                }

            } ;
            write-verbose -verbose:$true "$(Get-Date -Format 'HH:mm:ss'):Waiting 5secs to refresh";
            Start-Sleep -s 5 ;

            # secgrp membership seldom comes through clean, add a refresh loop
            do {
                $smsg = "===REVIEW SETTINGS:===`n----Updated Permissions:`n`nChecking Mailbox/AD Permission on $($Tmbx.samaccountname) mailbox `n to accessing user:`n $($oSG.SamAccountName)`n---" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                # 12:52 PM 9/25/2017 what if we want it trimmed & common layout'd:
                $smsg = "`n$((get-mailboxpermission -identity $($TMbx.Identity) -user $(($oSG).Name) -domaincontroller $($InputSplat.domaincontroller) | ?{$_.user -match ".*-(SEC|Data)-Email-.*$"} | format-list user,AccessRights,IsInhertied,Deny |out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                $smsg = "`n$((get-mailboxpermission -identity $($TMbx.Identity) -user $(($oSG).Name) -domaincontroller $($InputSplat.domaincontroller) | ?{$_.user -match ".*-(SEC|Data)-Email-.*$"} | format-list user,AccessRights,IsInhertied,Deny|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                # 12:52 PM 9/25/2017 what if we want it trimmed & common layout'd:
                $smsg = "`n==User mbx grant: Confirming $($TMbx.name) member of $($grpN):" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                # 10:04 AM 11/22/2017 put the accgrant confirmation into the output:
                if ($Tmbx.distinguishedname -match $rgxUserOUs) {
                    $smsg = "$((Get-ADPermission -identity $($TMbx.Identity) -domaincontroller $($InputSplat.domaincontroller) -user "$($oSG.SamAccountName)"|  format-list User,ExtendedRights,Inherited,Deny | out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                } else {
                    $smsg = "TMBX $($tMbx.samaccountname) is in a non-User OU: Term Hide/Unhide groups do not apply...";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                }  ;

                $smsg = "`nUpdated $($oSG.Name) Membership...`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):---";
                if ($mbrs = Get-ADGroupMember -identity $oSG.samaccountname -server $($DomainController) | select distinguishedName ) {
                    $smsg = "$(($mbrs | out-string).trim() | out-default)`n-----------------------" ;
                } else {
                    $smsg = "(NO MEMBERS RETURNED)`n-----------------------" ;
                } ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                $bRet = Read-Host "Enter Y to Refresh Review (replication latency)." ;
            } until ($bRet -ne "Y");

        } else { $smsg = "$($InputSplat.SecGrpName) not found.`n" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; };


    } # PROC-E ;

    END {


    } # END-E

}

#*------^ add-MailboxAccessGrant.ps1 ^------

#*----------------v Function add-MbxAccessGrant v------
function add-MbxAccessGrant {
    <#
    .SYNOPSIS
    add-MbxAccessGrant.ps1 - Configure fmailbox permissions, via middle-man AD Security Grp
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-
    FileName    :
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell,Exchange,Permissions,Exchange2010
    REVISIONS   :
    # 9:57 AM 9/27/2019 added `a beep to all "YYY" prompts to draw attn
    # 12:50 PM 6/13/2019 repl get-timestamp() -> Get-Date -Format 'HH:mm:ss' throughout
    # 11:05 AM 6/13/2019 updated get-admininitials()
    # 2:19 PM 4/29/2019 add global.ad.torolab.com to the domain param validateset on get-gcfast copy (sync'd in from verb-ex2010.ps1 vers)
    # 11:43 AM 2/15/2019 debugged update through a prod revision
    # 11:15 AM 2/15/2019 copied in bug-fixed write-log() with fixed debug support
    # 10:41 AM 2/15/2019 updated write-log to latest deferring version
    # 10:39 AM 2/15/2019 added full write-log logging support
    # 3:24 PM 2/6/2019 #1416:needs -prop * to pull msExchRecipientDisplayType,showInAddressBook,mail etc
    # 8:36 AM 9/6/2018 switched secgrp to Global->Universal scope, mail-enabled as DG, and hiddenfromaddressbook, debugged out issues, used in prod creation
    # 10:28 AM 6/27/2018 add $domaincontroller param option - skips dc discovery process and uses the spec, also updated $findOU code to work with torolab dom
    # 11:05 AM 3/29/2018 #1116: added trycatch, UST lacked the secgrp ou and was failing ou lookup
    # 10:31 AM 11/22/2017 shifted a block of "User mbx grant:" confirmation into review block, also tightened up the formatted whitespace to make the material pasted into cw reflect all that you need to know on the grant status. also added distrib code
    # 1:17 PM 11/15/2017 949: no, this needs to be the obj (was extracting samaccountname)
    # 12:35 PM 11/15/2017 debugged EXO-hosted Owner code to function. worked granting lynctest18 (exo) access to shared 'SharedTestEXOOwner' OP
    # 11:58 AM 11/15/2017 945: accommodate EXO-hosted Owners by testing with get-remotemailbox -AND get-mailbox on the owner spec.
    # 12:57 PM 9/25/2017 standardized mbxp & adperm field output and format i the review settings block
    # 11:29 AM 6/23/2017 fix typos, added 'DomainControler', without the vari-spec $
    # 8:16 AM 6/23/2017 we're getting mismatches/misses on AD work - prolly because we're using the -server $domain, rather than a SPECIFIC GC! replic lag is killing us!
    #   switch to the same gc the Ex cmds are using
    #   rplc -server $($InputSplat.Domain)  -> -server $($InputSplat.DomainController)
    #   rplc -server $Domain -> -server $($DomainController)
    # 1:41 PM 6/13/2017 spliced in latest 3/16/16 get-gcfast()
    # 1:37 PM 5/30/2017 855: pretest get-command, fails if it doesn't even have it at all
    # 1:21 PM 5/30/2017 1484: block that adds the $tmbx to the maintain-offboards.ps1 target AccGrant group for the region
    # 1:20 PM 5/30/2017 pencil in new AG group add when grant is done.
    # 11:27 AM 5/22/2017 add $NoPrompt
    # 9:44 AM 3/2/2017 suppress ID null errors on close BP
    # 9:07 AM 3/2/2017: Example code:  Don't force load EMS/Add-EMSRemote if there's an existing functional get-exchangeserver command (suppress clobber errors)
    # 9:41 AM 3/2/2017 merged in updated Add-EMSRemote Set
    # 9:07 AM 3/2/2017: Add-EMSRemote set: Example code:  Don't force load EMS/Add-EMSRemote if there's an existing functional get-exchangeserver command (suppress clobber errors)
    # 9:07 AM 3/2/2017 Don't force load EMS/Add-EMSRemote if there's an existing functional get-exchangeserver command (suppress clobber errors)
    # 12:15 PM 2/27/2017 trailing membership test was still failing, tore out blocks of new code and recycled what's used up in the Existing quote who cares if it's a DN vs a name
    # 12:11 PM 2/27/2017 fixed compat in SPB (prolly ADL too) - resolved any Owner entered as email, to the samaccountname ; #1081: drop the pipe!
     $oSG | Get-ADGroupMember -server $Domain | select distinguishedName ; #1263 threw up trying to do the get-aduser on the members, # 1283 replace user lookup with this (skip the getxxx member, pull members right out of properties)
    # 1:04 PM 2/24/2017 tweak below
    #12:56 PM 2/24/2017 doesn't run worth a damn LYN-> adl/spb, force it to abort (avoid half-built remote objects that take too long to replicate back to lyn)
    # 12:24 PM 2/24/2017 fixed updated membership report bug - pulled pipe, probably dehydrated object issue sin remote ps
    # 12:11 PM 2/24/2017 fix vscode/code.exe char set damage: It replaced dashes (-) with "ï¿½"
    # fix -join typo/damage
    # 12:44 PM 10/18/2016 update rgx for ticket to accommodate 5-digit (or 6) CW numbers "^\d{6}$"=>^\d{5,6}$
    # 9:11 AM 9/30/2016 added pretest if(get-command -name set-AdServerSettings -ea 0)
    # # 12:22 PM 6/21/2016 secgrp membership seldom comes through clean, add a refresh loop
    # 10:52 AM 6/20/2016 fixed typo $InputSplatSiteOverride => $InputSplat.SiteOverride (broke -SiteOverride function)
    # 11:02 AM 6/7/2016 updated get-aduser review cmds to use the same dc, not the -domain global.ad.toro.com etc
    # 1:34 PM 5/26/2016 confirmed/verified works fine with SPB-hosted mbx under 376336 issitjx
    # 11:45 AM 5/19/2016 corrected $tmbx ref's to use $tmbx.identity v. $tmbx.samaccountname, now working. Retry code in place for SPB, but it didn't trigger during testing
    # 2:37 PM 5/18/2016 implmented Secgrp OU and Secgrp stnd name
    # 2:28 PM 5/18/2016 support dmg's latest unilateral changes: With the recent AD changes, all email access groups should be named         XXX-SEC-Email-firstname lastname-G     and stored in XXX\Managed Groups\SEC Groups\Email Access.
    # 2:17 PM 5/10/2016 used successfully to set a LYN manager perm's on an SPBMS640Mail02-hosted user. didn't time out, Set-MailboxPermission command completed after ~3 secs
    #     fixed bad param example, remmed out non-functional Owner in the SGSplat (nosuch param), and re-enabled the ManagedBy on the SG - it's not a mbx,
    #     so why not set ManagedBy, doesn't get used by the org chart in SP
    # 2:38 PM 3/17/2016 stop populating anything into any managed-by; it's an OrgChart political value now. Rename ManagedBy param and object names in here to 'Owner'
    # 1:12 PM 2/11/2016 fixed new bug in get-GCFast, wasn't detecting blank $site
    # 12:20 PM 2/11/2016 updated to standard EMS/AD Call block & Add-EMSRemote()
    # 9:36 AM 2/11/2016 just shifting to a single copy, with no # Requires at all, losing the -psv2.ps1 version
    # 2:23 PM 2/10/2016 debugged mismatched {}, working from SPB now
    # 1:54 PM 2/10/2016 recoded to work on SPB and ADL, this version just needs the #Requires -Version 3 for psv2 enabled to be a psv3 version
    #         added fundemental upgrade to the AD Site detection, to work from SPB/Spellbrook and ADL Adeliade
    # 12:07 PM 2/10/2016 Psv2 variant - at this point, the only real diff is the rem'd rem'd #Requires -Version 3 for psv2
    # 7:40 AM Add-EMSRemote: 2/5/2016 another damn cls REM IT! I want to see all the connectivity info, switched wh->wv, added explicit echo's of what it's doing.
    # 10:41 AM 1/13/2016 updated Add-EMSRemote set & removed Clear-Host's
    # 10:02 AM 1/13/2016: fixed cls bug due to spurious ";cls" included in the try/catch boilerplate: Write-Error "$((get-date).ToString('HH:mm:ss')): Command: $($_.InvocationInfo.MyCommand)" ;cls => Write-Error "$((get-date).ToString('HH:mm:ss')): Command: $($_.InvocationInfo.MyCommand)" ;
    # 1:02 PM 12/18/2015 missing SYD as well
    # 11:42 AM 12/18/2015: building a Psv2-compliant version (-psv2.ps1):
    - sub out all [ordered] (-psv2 only)
    - rem'd #Requires -Version 3 (-psv2 only)
    - added explicit .tostring() in front of all string handlers (.substring() etc) (added to both versions)
    # 3:08 PM 10/29/2015 added in XIA aware from other recent script updates, and -server xxx to all get-ad* that didn't have it
    #2:49 PM 10/29/2015 add entire MEL site, nothing in the OU or Secgrp name switch blocks
    # 9:08 AM 10/14/2015 added debugpref maint code to get write-debug to work
    # 8:04 AM 10/14/2015 add sendAS adperms
    # 7:31 AM 10/14/2015 added -dc specs to all *-user & *-mailbox cmds, to ensure we're pulling back data from same dc that was updated in the set-* commands
    # 7:19 AM 10/14/2015 fixed some typos, made sure all $InputSplat.domaincontroller were $()'d
    # 9:13 AM 10/12/2015 force $Grantsplat=$Tmbx to use $Tmbx.Samacctname, defaulting to displayname which isn't consistently resolvable, samacctname should be.
    # 1:17 PM 10/6/2015 update to spec, seems to work
    # splice in Add-EMSRemote set & get-gcfast
    # 10:49 AM 10/6/2015: updated vers of Get-AdminInitials
    # 2:49 PM 10/2/2015 updated catch block to be detailed
    # 10:57 AM 8/14/2015 defaulted PermsDays to 60 (was going to 999)
    # 10:46 AM 8/14/2015 add param examples for the PermsDays spec
    # 9:35 AM 8/14/2015 updated params examples to reflect use of -ticket & -siteoverride
    # 11:00 AM 8/12/2015 also add an Info ref for the admin doing the work
    # 10:37 AM 8/12/2015 I see from dumping all matching secgrps ...
      $AllSGs = Get-ADGroup  -filter {GroupCategory  -eq "Security"  -and GroupScope -eq "Global"} -properties info,description;
      $AllSGs = $AllSgs |?{$_.Name -match "^\w{3}-(SEC|Data)-Email-.*$"} ; ($AllSGs | measure).count ;
      # have to 2-stage filter as the get-adgroup -filter has no -match regex operator support
      ... that dawn used to use these, oldest record I see, was from 2005, most recent appears to have been 2011. But she was recording sometimes recording the incident req# - which is a useful item to include (esp if you don't want folks monkeyin with the Notes/Info
      value and breaking automation to clean these up!).
      So we need to add incident number to the add-MbxAccessGrant.ps1, that means a parameter and a 3rd line in the Info append
    #11:32 AM 8/5/2015 fixed trailing ) in Updated IRO-SEC-Email-Jodie Gilroy-G)
    #11:43 AM 7/20/2015 line 197added a :space after displayname:
    # 12:18 PM 7/17/2015 added -ea silentlycontinue to get-mailboxpermission - it was causing it to bomb script when no match found
    # 1:55 PM 6/15/2015 initial version
    .DESCRIPTION
    add-MbxAccessGrant.ps1 - Configure fmailbox permissions, via middle-man AD Security Grp
    .PARAMETER TargetID
    Target Mailbox for Access Grant[name,emailaddr,alias]
    .PARAMETER SecGrpName
    Custom override default generated name for Perm-hosting Security Group[[SIT]-SEC-Email-[DisplayName]-G]
    .PARAMETER Owner
    Specify the userid to be responsible for access-grant-approvals[name,emailaddr,alias]
    .PARAMETER SiteOverride
    Specify a 3-letter Site Code. Used to force DL name/placement to vary from TargetID's current site[3-letter Site code]
    .PARAMETER PermsDays
    Specify the number of day's the access-grant should be in place. (60 default. 999=permanent)[30-60,999]")]
    .PARAMETER Members
    Comma-delimited string of potential users to be granted access[name,emailaddr,alias]
    .PARAMETER ticket
    Incident number for the change request[[int]nnnnnn]
    .PARAMETER NoPrompt
    Suppress YYY confirmation prompts
    .PARAMETER domaincontroller
    Option to hardcode a specific DC [-domaincontroller xxxx]
    .PARAMETER Whatif
    Parameter to run a Test no-change pass, and log results [-Whatif switch]
    .PARAMETER ShowProgress
    Parameter to display progress meter [-ShowProgress switch]
    .PARAMETER ShowDebug
    Parameter to display Debugging messages [-ShowDebug switch]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    .\add-MbxAccessGrant.ps1 -ticket 123456 -SiteOverride LYN -TargetID lynctest13 -Owner kadrits -PermsDays 999 -members "lynctest16,lynctest18" -showDebug -whatIf ;
    Parameter Whatif test with Debug messages displayed
    .LINK
    *----------^ END Comment-based Help  ^---------- #>
    [CmdletBinding()]
    Param(
        [Parameter(HelpMessage = "Target Mailbox for Access Grant[name,emailaddr,alias]")]
        [string]$TargetID,
        [Parameter(HelpMessage = "Custom override default generated name for Perm-hosting Security Group[[SIT]-SEC-Email-[DisplayName]-G]")]
        [string]$SecGrpName,
        [Parameter(HelpMessage = "Specify the userid to be responsible for access-grant-approvals[name,emailaddr,alias]")]
        [string]$Owner,
        [Parameter(HelpMessage = "Specify the number of day's the access-grant should be in place. (60 default. 999=permanent)[30-60,999]")]
        [ValidateRange(7, 999)]
        [int]$PermsDays,
        [Parameter(HelpMessage = "Specify a 3-letter Site Code. Used to force DL name/placement to vary from TargetID's current site[3-letter Site code]")]
        [string]$SiteOverride,
        [Parameter(HelpMessage = "Comma-delimited string of potential users to be granted access[name,emailaddr,alias]")]
        [string]$Members,
        [Parameter(HelpMessage = "Incident number for the change request[[int]nnnnnn]")]
        [int]$Ticket,
        [Parameter(HelpMessage = "Suppress YYY confirmation prompts [-NoPrompt]")]
        [switch] $NoPrompt,
        [Parameter(HelpMessage = "Option to hardcode a specific DC [-domaincontroller xxxx]")]
        [string]$domaincontroller,
        [Parameter(HelpMessage = 'Debugging Flag [$switch]')]
        [switch] $showDebug,
        [Parameter(HelpMessage = 'Whatif Flag [$switch]')]
        [switch] $whatIf
    ) ;

    # NoPrompt Suppress YYY confirmation prompts

    BEGIN {
        $verbose = ($VerbosePreference -eq "Continue") ;


        #region SPLATDEFS ; # ------

        if (($host.version.major) -lt 3) {
            $InputSplat = @{
                TargetID     = "TARGETMBX";
                SecGrpName   = "";
                Owner        = "LYNCTEST2"
                PermsDays    = 60;
                SiteOverride = "";
                Members      = "LYNCTEST3"
            } ;
            # 2:02 PM 5/10/2016 owner isn't a param on NEW-ADGroup => ManagedBy
            # 7:28 AM 9/6/2018 switch to EXO-compatible group type: Univ, mail-enable
            $SGSplat = @{
                Name            = "";
                DisplayName     = "";
                SamAccountName  = "";
                GroupScope      = "Global";
                GroupCategory   = "Universal";
                ManagedBy       = "";
                Description     = "";
                OtherAttributes = "";
                Path            = "";
                Server          = ""
            };
            $SGUpdtSplat = @{
                Identity = "";
                Server   = ""
            };
            $DGEnableSplat = @{
                Identity         = "";
                DomainController = "" ;
            } ;
            $DGUpdtSplat = @{
                Identity                      = "";
                HiddenFromAddressListsEnabled = $true ;
                DomainController              = "" ;
            } ;
            $GrantSplat = @{
                Identity        = "" ;
                User            = "" ;
                AccessRights    = "FullAccess";
                InheritanceType = "All";
            };
            # 8:05 AM 10/14/2015 add for AD SendAs perms grant
            <#$ADMbxGrantSplat=@{
	          Identity="" ;
	          User="" ;
	          ExtendedRights="Send As" ;
	        };#>
            #8:59 AM 10/14/2015 try pulling id, pipeline it in
            $ADMbxGrantSplat = @{
                User           = "" ;
                ExtendedRights = "Send As" ;
            };
        } else {
            # 12:04 PM 2/10/2016 psv3 code
            $InputSplat = [ordered]@{
                TargetID     = "TARGETMBX";
                SecGrpName   = "";
                Owner        = "LYNCTEST2"
                PermsDays    = 60;
                SiteOverride = "";
                Members      = "LYNCTEST3"
            } ;
            $SGSplat = [ordered]@{
                Name            = "";
                DisplayName     = "";
                SamAccountName  = "";
                GroupScope      = "Universal";
                GroupCategory   = "Security";
                ManagedBy       = "";
                Description     = "";
                OtherAttributes = "";
                Path            = "";
                Server          = ""
            };
            $SGUpdtSplat = [ordered]@{
                Identity = "";
                Server   = ""
            };
            $DGEnableSplat = [ordered]@{
                Identity         = "";
                DomainController = ""
            };
            $DGUpdtSplat = [ordered]@{
                Identity                      = "";
                HiddenFromAddressListsEnabled = $true ;
                DomainController              = "" ;
            } ;
            $GrantSplat = [ordered]@{
                Identity        = "" ;
                User            = "" ;
                AccessRights    = "FullAccess";
                InheritanceType = "All";
            };
            # 8:05 AM 10/14/2015 add for AD SendAs perms grant
            <#$ADMbxGrantSplat=[ordered]@{
	          Identity="" ;
	          User="" ;
	          ExtendedRights="Send As" ;
	        };#>
            #8:59 AM 10/14/2015 try pulling id, pipeline it in
            $ADMbxGrantSplat = [ordered]@{
                User           = "" ;
                ExtendedRights = "Send As" ;
            };
        }

        if ($PSCommandPath) { $pltInput.add("ParentPath", $PSCommandPath) } ;
        if ($TargetID) { $InputSplat.TargetID = $TargetID };
        if ($SecGrpName) { $InputSplat.SecGrpName = $SecGrpName };
        if ($Owner) { $InputSplat.Owner = $Owner };
        if ($PermsDays) { $InputSplat.PermsDays = $PermsDays };
        if ($SiteOverride) { $InputSplat.SiteOverride = $SiteOverride };
        if ($Members) { $InputSplat.Members = $Members };

        $smsg = "`nSpecified Target Email: $($InputSplat.TargetID)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

        #endregion SPLATDEFS ; # ------
        #region LOADMODS ; # ------
        $rgxExoPsHostName = "^(ps\.outlook\.com|outlook\.office365\.com)$" ;
        $rgxRemsPssName = "^(Exchange2010|Session\sfor\simplicit\sremoting\smodule\sat\s.*)" ;
        $rgxSnapPssname = "^Session\d{1}$" ;
        $rgxEx2010SnapinName = "^Microsoft\.Exchange\.Management\.PowerShell\.E2010$";
        $Ex2010SnapinName = "Microsoft.Exchange.Management.PowerShell.E2010" ;

        #
        #LEMS detect: IdleTimeout -ne -1
        if (get-pssession | ? { ($_.configurationname -eq 'Microsoft.Exchange') -AND ($_.ComputerName -match $rgxEx10HostName) -AND ($_.IdleTimeout -ne -1) } ) {
            write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):LOCAL EMS detected" ;
            $Global:E10IsDehydrated = $false ;
            # REMS detect dleTimeout -eq -1
        } elseif (get-pssession | ? { $_.configurationname -eq 'Microsoft.Exchange' -AND $_.ComputerName -match $rgxEx10HostName -AND ($_.IdleTimeout -eq -1) } ) {
            write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):REMOTE EMS detected" ;
            $reqMods += "get-GCFast;Get-ExchangeServerInSite;connect-Ex2010;Reconnect-Ex2010;Disconnect-Ex2010;Disconnect-PssBroken".split(";") ;
            if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ; }  ;
            reconnect-ex2010 ;
            $Global:E10IsDehydrated = $true ;
        } else {
            write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):No existing Ex2010 Connection detected" ;
            # Server snapin defer
            if (($host.version.major -lt 3) -AND (get-service MSExchangeADTopology -ea SilentlyContinue)) {
                write-verbose -verbose:$bshowVerbose  "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Loading Local Server EMS10 Snapin" ;
                $reqMods += "Load-EMSSnap;load-EMSLatest".split(";") ;
                if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ; }  ;
                Load-EMSSnap ;
                $Global:E10IsDehydrated = $false ;
            } else {
                # if you want REMS - (assumed on new scripts)
                $reqMods += "connect-Ex2010;Reconnect-Ex2010;Disconnect-Ex2010;Get-ExchangeServerInSite;Disconnect-PssBroken".split(";") ;
                if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ; }  ;
                reconnect-ex2010 ;
                $Global:E10IsDehydrated = $true ;
            } ;
        } ;
        #

        # load ADMS
        $reqMods += "load-ADMS".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ; }  ;
        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):(loading ADMS...)" ;
        load-ADMS ;

        # EXO connection
        #$reqMods+="connect-exo;Reconnect-exo;Disconnect-exo".split(";") ;
        <#if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):(loading EXO...)" ;
        reconnect-exo ;
        #>

        <# RLMS connection
        $reqMods+="Get-LyncServerInSite;load-LMS;Disconnect-LMSR".split(";") ;
        if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):(loading LMS...)" ;
        Reconnect-L13 ;
        #>

        $AdminInits = get-AdminInitials ;

        #region LOADMODS ; # ------

    }  # BEG-E ;

    PROCESS {

        #region DATAPREP ; # ------

        $Tmbx = (get-mailbox $($InputSplat.TargetID) -domaincontroller (Get-ADDomainController).Name.tostring() -ea stop) ;
        $GrantSplat.Identity = $($Tmbx.samaccountname);
        $domain = $Tmbx.identity.tostring().split("/")[0]
        $InputSplat.Add("Domain", $($domain) ) ;
        if (!$domaincontroller) {
            switch ($domain) {
                "global.ad.toro.com" {
                    # 10:42 AM 2/11/2016 no above is redundant code: get-GcFast will do same lookup on local computer's site if not specified, use that
                    #$domaincontroller =(get-gcfast -domain global.ad.toro.com -site $ADSite)
                    $domaincontroller = (get-gcfast -domain global.ad.toro.com)
                } # global block end
                "china.ad.toro.com" {
                    if (($env:USERDOMAIN -eq "TORO-LAB") -AND ($domain -eq "china.ad.toro.com")) { break }
                    $domaincontroller = (get-gcfast -domain china.ad.toro.com)
                } # cn block end ;
                "global.ad.torolab.com" {
                    $domaincontroller = (get-gcfast -domain global.ad.torolab.com) ;
                } ;
            } # switch-E $domain ;
        } else {
            $smsg = "Using hard-coded domaincontroller:$($domaincontroller)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
        } ;

        $InputSplat.Add("DomainController", $domaincontroller) ;
        $SGUpdtSplat.Server = $($InputSplat.DomainController);
        $DGEnableSplat.DomainController = $($domaincontroller);
        $DGUpdtSplat.DomainController = $($domaincontroller);
        $InputSplat.Site = ($Tmbx.identity.tostring().split('/')[1]) ;

        switch ((get-recipient -Identity $Inputsplat.Owner).RecipientType ) {
            "UserMailbox" {
                #$Inputsplat.OwnerMbx = get-mailbox -identity $Inputsplat.Owner -ea stop | select -expand SamAccountname ;
                # 1:15 PM 11/15/2017 no, this needs to be the obj
                #$Inputsplat.OwnerMbx = get-mailbox -identity $Inputsplat.Owner -ea stop  ;

                if ( ($InputSplat.OwnerMbx = (get-mailbox -identity $($InputSplat.Owner) -ea stop)) ) {
                    if ($showdebug) { $smsg = "UserMailbox detected" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; } ;
                } else {
                    # 11:54 AM 11/15/2017 without the -ea stop, we need an explicit error
                    throw "Unable to resolve $($InputSplat.Owner) to any existing OP or EXO mailbox" ;
                    Cleanup ; Exit ;
                } ;
            }
            "MailUser" {
                if ( ($InputSplat.OwnerMbx = (get-remotemailbox -identity $($InputSplat.Owner) -ea stop)) ) {
                    if ($showdebug) {
                        $smsg = "MailUser detected" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; #Error|Warn|Debug
                    } ;
                } else {
                    # 11:54 AM 11/15/2017 without the -ea stop, we need an explicit error
                    throw "Unable to resolve $($InputSplat.Owner) to any existing OP or EXO mailbox" ;
                    Cleanup ; Exit ;
                } ;
            }
            default {
                throw "$($InputSplat.Owner) Not found, or unrecognized RecipientType" ;
                Cleanup ; Exit ;
            }
        } ;

        # owner needs to be samaccountname or DN - can't use email addresses!
        if ($Inputsplat.Owner -match $rgxEmailAddr) {
            $smsg = "Converting Owner email:$($Inputsplat.Owner) to logon..." ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
            #$Inputsplat.Owner = get-mailbox -identity $Inputsplat.Owner -ea stop | select -expand SamAccountname ;
            $Inputsplat.Owner = $InputSplat.OwnerMbx.alias ;
        } ;

        # if no permsdays, default it to 60d
        if (!($InputSplat.PermsDays)) { "defaulting PermsDays to 60"; $InputSplat.PermsDays = 60 };
        if ($InputSplat.PermsDays -eq 999) {
            [string]$PermsExp = (get-date "12/31/2099" -format "MM/dd/yyyy") ;
        } else {
            [string]$PermsExp = (get-date (Get-Date).AddDays($InputSplat.PermsDays + 1) -format "MM/dd/yyyy") ;
        } ;

        $Infostr = "TargetMbx:$($Tmbx.samaccountname)`r`nPermsExpire:$($PermsExp)`r`nIncident:$($Ticket)`r`nAdmin:$($AdminInits)`r`nBusinessOwner:$($InputSplat.Owner);`r`nITOwner:$($InputSplat.Owner)" ;

        $smsg = "Site:$($InputSplat.Site)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
        $smsg = "`nTLogon: $($Tmbx.samaccountname )" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

        $Adu = (get-aduser $($tmbx.SamAccountName) -server $($InputSplat.DomainController) -ea stop -properties manager)  ;
        if ($Adu.Manager) {
            $Mgr = ((get-aduser ($Adu.manager)).samaccountname) ;
            $smsg = "MgrLogon: $($Mgr)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
        } else {
            $smsg = "$($Tmbx.displayname) has a blank AD Manager field.`nAsserting Owner from inputs:$($InputSplat.Owner) " ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
            $Mgr = $($InputSplat.Owner);
        } ;
        # 10:52 AM 6/20/2016 fixed typo $InputSplatSiteOverride => $InputSplat.SiteOverride (broke -SiteOverride function)
        if ($InputSplat.SiteOverride) {
            $SiteCode = $($InputSplat.SiteOverride);
        } else {
            # 12:28 PM 11/15/2017 we need to use the OwnerMbx - Owner currently is the alias, we want the object with it's dn
            $SiteCode = $InputSplat.OwnerMbx.identity.tostring().split("/")[1]  ;
        } ;
        $Domain = "global.ad.toro.com" ;
        if ($env:USERDOMAIN -eq 'TORO') {
            $FindOU = "^OU=Email\sAccess,OU=SEC\sGroups,OU=Managed\sGroups,";
        } ELSEif ($env:USERDOMAIN -eq 'TORO-LAB') {
            # CN=Lab-SEC-Email-Thomas Jefferson,OU=Email Access,OU=SEC Groups,OU=Managed Groups,OU=LYN,DC=global,DC=ad,DC=torolab,DC=com
            $FindOU = "^OU=Email\sAccess,OU=SEC\sGroups,OU=Managed\sGroups,"; ;
        } else {
            throw "UNRECOGNIZED USERDOMAIN:$($env:USERDOMAIN)" ;
        } ;

        $SGSplat.DisplayName = "$($SiteCode)-SEC-Email-$($Tmbx.DisplayName)-G";

        TRY {
            $OU = (Get-ADObject -filter { ObjectClass -eq 'organizationalunit' } -server $($DomainController) | ? { $_.distinguishedname -match "^$($FindOU).*OU=$($SiteCode),.*,DC=ad,DC=toro((lab)*),DC=com$" } | select distinguishedname).distinguishedname.tostring() ;
        } CATCH {
            $ErrTrpd = $_ ;
            $smsg = "UNABLE TO LOCATE $($FindOU) BELOW SITECODE $($SiteCode)!. EXITING!" ; $smsg = "MESSAGE" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; #Error|Warn
            $smsg = "Failed processing $($ErrTrpd.Exception.ItemName). `nError Message: $($ErrTrpd.Exception.Message)`nError Details: $($ErrTrpd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; #Error|Warn
            Exit
        } ;
        If ($OU -isnot [string]) {
            $smsg = "WARNING AD OU SEARCH SITE:$($InputSplat.SiteCode), FindOU:$($FindOU), FAILED TO RETURN A SINGLE OU...`n$($OU.distinguishedname)`nEXITING!";
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; #Error|Warn
            Exit ;
        } ;

        $smsg = "$SiteCode" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
        $smsg = "$OU" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

        $SGSplat.Path = $OU ;
        $smsg = "Checking specified SecGrp Members..." ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
        $SGMembers = ($InputSplat.members.split(",") | foreach { get-recipient $_ -ea stop })
        $smsg = "Checking for existing $($SGSplat.DisplayName)..." ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

        if ($bDebug) {
            $smsg = "`$SGSrchName:$($SGSrchName)`n`$SGSplat.DisplayName: $($SGSplat.DisplayName)"; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ;
        } ;

        $SGSrchName = $($SGSplat.DisplayName);
        $oSG = Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -server $($InputSplat.DomainController) -ErrorAction stop;

        if ($oSG) {
            if ($bDebug) {
                $smsg = "`$oSG:$($oSG.SamAccountname)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; #Error|Warn
                $smsg = "`$oSG.DN:$($oSG.DistinguishedName)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; #Error|Warn
            } ;
            # 2:38 PM 3/21/2016 we should _append_ the $InfoStr into any existing Info for the object
            # 11:58 AM 2/27/2017 can't use [ordered] on psv2 if we must have these in order use a psv2 OrderedDictionary
            if (($host.version.major) -lt 3) {
                <#
                $ADOtherInfoProps=@{
                TargetMbx=$null ;
                PermsExpire=$null ;
                Incident=$null ;
                Admin=$null ;
                BusinessOwner=$null ;
                ITOwner=$null ;
                } ;
                #>
                $ADOtherInfoProps = New-Object Collections.Specialized.OrderedDictionary ;
                #$ADOtherInfoProps.Add('One',1) ;
                $ADOtherInfoProps.Add('TargetMbx', $null) ;
                $ADOtherInfoProps.Add('PermsExpire', $null) ;
                $ADOtherInfoProps.Add('Incident', $null) ;
                $ADOtherInfoProps.Add('Admin', $null) ;
                $ADOtherInfoProps.Add('BusinessOwner', $null) ;
                $ADOtherInfoProps.Add('ITOwner', $null) ;

            } else {
                $ADOtherInfoProps = [ordered]@{
                    TargetMbx     = $null ;
                    PermsExpire   = $null ;
                    Incident      = $null ;
                    Admin         = $null ;
                    BusinessOwner = $null ;
                    ITOwner       = $null ;
                } ;
            } ;
            #$Infostr="TargetMbx:$($Tmbx.samaccountname)`r`nPermsExpire:$($PermsExp)`r`nIncident:$($Ticket)`r`nAdmin:$($AdminInits)`r`nBusinessOwner:$($InputSplat.Owner);`r`nITOwner:$($InputSplat.Owner)" ;
            if ($oSG.info) {
                # existing info tag
                # update the splat
                # 12:19 PM 3/22/2016 just loop each line split on `n: (Get-ADUser lynctest9 -Properties info).info.split("`n")| foreach{"Ln:$_"}
                $oADOtherInfo = New-Object PSObject -Property $ADOtherInfoProps ;

                #( $ln in ($oSG.info.tostring().split("`n") )  {
                $ilines = $oSG.info.tostring().split("`n").count ;
                $iIter = 0 ;
                foreach ( $ln in $oSG.info.tostring().split("`n") ) {
                    $iIter++

                    if ($iIter -eq 1) { $UpdInfo = $null; } ;

                    if ($ln -match "^(TargetMbx|PermsExpire|Incident|Admin|BusinessOwner|ITOwner):.*$") {
                        # it's part of a defined Info tag
                        $matches = $null ;
                        # ingest the matches and throw away the lines
                        if ($ln -match "(?<=TargetMbx:)\w+" ) { $oADOtherInfo.TargetMbx = $matches[0] } ; $matches = $null ;
                        if ($ln -match "(?<=PermsExpire:)\d+\/\d+/\d+" ) { $oADOtherInfo.PermsExpire = (get-date $matches[0]) ; } ; ; $matches = $null ;
                        # 12:44 PM 10/18/2016 update rgx for ticket to accommodate 5-digit (or 6) CW numbers "^\d{6}$"=>^\d{5,6}$
                        if ($ln -match "(?<=Incident:)^\d{5,6}$") { $oADOtherInfo.Incident = $matches[0] ; } ; $matches = $null ;
                        if ($ln -match "(?<=Admin:)\w*") { $oADOtherInfo.Admin = $matches[0] ; } ; $matches = $null ;
                        if ($ln -match "(?<=BusinessOwner:)\w{2,20}") { $oADOtherInfo.BusinessOwner = $matches[0] ; } ; $matches = $null ;
                        if ($ln -match "(?<=ITOwner:)\w{2,20}") { $oADOtherInfo.ITOwner = $matches[0] ; } ; $matches = $null ;
                    } else {
                        $UpdInfo += "$($ln)`r`n" ;
                    } ;

                    if ($iIter -eq $iLines) {
                        $smsg = "`$uinfo:`n$uinfo" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                        if ($oADOtherInfo) {
                            $smsg = "Updating existing Info tag:`n$(($oADOtherInfo |out-string).trim())";
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                        }
                        $UpdInfo += "`r`nTargetMbx:$($Tmbx.samaccountname)`r`nPermsExpire:$($PermsExp)`r`nIncident:$($Ticket)`r`nAdmin:$($AdminInits)`r`nBusinessOwner:$($InputSplat.Owner);`r`nITOwner:$($InputSplat.Owner)" ;
                        if ($bDebug) {
                            $smsg = "New Info field:`n$(($UpdInfo |out-string).trim())";
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                        } ;

                        #Set-ADUser -identity $tusr -Replace @{info="$($uinfo)"} -server LYNMS811 -whatif  ;
                        Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -server $($InputSplat.DomainController) -ErrorAction stop | Set-ADUser -Replace @{info = "$($UpdInfo)" }  -whatif ; ;
                    }


                } # loop-E $lines

            } ; # if-E $osg

        } else {
            $smsg = "$($SGSplat.DisplayName) Not found. Testing Create with the following paraemters..."  ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
            # create the secgrp
            $SGSplat.Name = $($SGSplat.DisplayName);
            $SGSplat.SamAccountName = $($SGSplat.DisplayName);
            $SGSplat.ManagedBy = $($InputSplat.Owner);
            $SGSplat.Description = "Email - access to $($Tmbx.displayname)'s mailbox";
            $SGSplat.Server = $($InputSplat.DomainController) ;
            # build the Notes/Info field as a hashcode: OtherAttributes=@{    info="TargetMbx:kadrits`r`nPermsExpire:6/19/2015"  } ;
            $SGSplat.OtherAttributes = @{info = $($Infostr) } ;


            $smsg = "`$SGSplat:`n---"; $smsg = "MESSAGE" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
            foreach ($row in $SGSplat) {
                foreach ($key in $row.keys) {
                    if ($key -eq "OtherAttributes") {
                        $smsg = "==v OtherAttributes: v==" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                        $SGSplat.OtherAttributes.GetEnumerator() | Foreach-Object {
                            $smsg = "==$($_.Key ):==`n$(($_.Value|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                        } ;
                        $smsg = "==^ OtherAttributes: ^==" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                    } else {
                        $smsg = "$($key): $($row[$key])" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                    } ;
                }
            } ;

            $smsg = "---`nWhatif $($SGSplat.Name) creation...";
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

            New-AdGroup @SGSplat -whatif -ea stop;
            $DGEnableSplat.identity = $SGSplat.SamAccountName ;
            $DGUpdtSplat.identity = $SGSplat.SamAccountName ;

            $smsg = "`$DGEnableSplat:`n---`n$(($DGEnableSplat|out-string).trim())`n---`n" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

            write-host -foregroundcolor yellow "$((get-date).ToString("HH:mm:ss")):Continue with $($SGSplat.Name) creation?...";
            if ($NoPrompt) { $bRet = "YYY" } else { $bRet = Read-Host "Enter YYY to continue. Anything else will exit`a" ; } ;
            if ($bRet.ToUpper() -eq "YYY") {


                if ($bWhatif) {
                    $smsg = "-Whatif pass, skipping exec." ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                } else {
                    $smsg = "Executing...";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                    New-AdGroup @SGSplat -ea stop ;
                    Do { write-host "." -NoNewLine; Start-Sleep -s 1 } Until (Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -server $($InputSplat.DomainController)) ;
                    #$oSG= (get-adgroup "$($SGSplat.DisplayName)" -server $($InputSplat.Domain) -ea stop );
                    $oSG = Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -server $($InputSplat.DomainController) -ErrorAction stop;
                    if ($bDebug) {
                        $smsg = "`$oSG:$($oSG.SamAccountname)`n`$oSG.DN:$($oSG.DistinguishedName)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; #Error|Warn
                    } ;
                    $smsg = "Enable-DistributionGroup w`n$(($DGEnableSplat|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    Enable-DistributionGroup @DGEnableSplat ;
                    $smsg = "Set HiddenFromAddressListsEnabled:Set-DistributionGroup w`n$(($DGUpdtSplat|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    Set-DistributionGroup @DGUpdtSplat ;
                    $oSG = Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -prop * -server $($InputSplat.DomainController) -ErrorAction stop;
                    $smsg = "Final SecGrp Config:$($oSG.SamAccountname)`n:$(($oSG | fl Name,GroupCategory,GroupScope,msExchRecipientDisplayType,showInAddressBook,mail|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                } ;
            } else { $smsg = "INVALID KEY ABORTING NO CHANGE!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; Exit ; } ;
        } # if-E $osg

        $smsg = "`nTesting SecGrp Members Add `nto group: $($oSG.Name)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
        # *** BREAKPOINT ;
        if ($oSG) {
            $ExistMbrs = @() ;
            # 11:27 AM 6/23/2017 typo, vari with no leading $
            $oSG | Get-ADGroupMember -server $($DomainController) | Select-Object -ExpandProperty sAMAccountName | ForEach-Object { $ExistMbrs += $_ } ;
            $SGUpdtSplat.Identity = $($oSG.samaccountname) ;
            $DGEnableSplat.Identity = $($oSG.samaccountname) ;
            $DGUpdtSplat.Identity = $($oSG.samaccountname) ;
            $GrantSplat.User = $($oSG.SamAccountName);
            #8:41 AM 10/14/2015 add adp
            $ADMbxGrantSplat.User = $($oSG.SamAccountName);
            $SGUpdtSplat.Server = $($InputSplat.DomainController) ;
            $DGEnableSplat.DomainController = $($InputSplat.DomainController) ;
            $DGUpdtSplat.DomainController = $($InputSplat.DomainController) ;
            # 12:47 PM 10/6/2015 add dc
            $GrantSplat.Add("DomainController", $($InputSplat.domaincontroller)) ;
            #8:41 AM 10/14/2015 add adp
            $ADMbxGrantSplat.Add("DomainController", $($InputSplat.domaincontroller)) ;

            if ($bWhatif) {
                $smsg = "-Whatif pass, skipping exec." ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
            } else {
                foreach ($Mbr in $SGMembers) {
                    If ($ExistMbrs -notcontains $Mbr.sAMAccountName) {
                        $smsg = "Test ADD:$($mbr.samaccountname)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                        Add-ADGroupMember @SGUpdtSplat -members $($mbr.samaccountname) -ea stop -whatif ;
                        <# 8:27 AM 6/27/2019 after win10ug, the above is now throwing:
                        ...
                        Testing SecGrp Member Add
                        to group: LYN-SEC-Email-wpaIRR-G
                        08:21:12:Test ADD:antoidx
                        add-MbxAccessGrant : Parameter cannot be processed because the parameter name 'member' is ambiguous. Possible matches include: -Members -MemberTimeToLive.
                        At C:\usr\work\exch\scripts\add-MbxAccessGrant.ps1:2510 char:1
                        + add-MbxAccessGrant @pltInput
                        Looks like they've added a new param, and it's broken partials,
                        curr doc: https://docs.microsoft.com/en-us/powershell/module/addsadministration/add-adgroupmember?view=win10-ps
                        Add-ADGroupMember  [-WhatIf]  [-Confirm]  [-AuthType <ADAuthType>]  [-Credential <PSCredential>]  [-Identity] <ADGroup>  [-Members] <ADPrincipal[]>  [-MemberTimeToLive <TimeSpan>]  [-Partition <String>]  [-PassThru]  [-Server <String>]  [<CommonParameters>]
                        # yea the param is named -members, not -member, with the addition it broke the auto-resolution
                        #>
                    } else {
                        $smsg = "SKIPPING:$($mbr.samaccountname) is already a member of $($oSG.samaccountname)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    } ;
                }  # loop-E ;
                $smsg = "Continue with Member Addition?...";
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                if ($NoPrompt) { $bRet = "YYY" } else { $bRet = Read-Host "Enter YYY to continue. Anything else will exit`a" ; } ;
                if ($bRet.ToUpper() -eq "YYY") {
                    $smsg = "Exec Update";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    foreach ($Mbr in $SGMembers) {
                        If ($ExistMbrs -notcontains $Mbr.sAMAccountName) {
                            "Exec ADD:$($mbr.samaccountname)"
                            if ($whatif) {
                                # 11:17 AM 6/22/2015 whatif-only pass
                                $smsg = "SKIPPING EXEC: Whatif-only pass";
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                            } else {
                                # 8:33 AM 6/27/2019 fix latest ADmod, added a conflicting param, autoresolve fails, typo -member -> proper -members
                                Add-ADGroupMember @SGUpdtSplat -members $($mbr.samaccountname) -ea stop ;
                            } ;
                        } else {
                            "SKIPPING:$($mbr.samaccountname) is already a member of $($oSG.samaccountname)"
                        } ;
                    } #  # loop-E;
                } ;
            } # if-E whatif ;
            $mbxp = $Tmbx | get-mailboxpermission -user ($oSG).Name -domaincontroller $InputSplat.domaincontroller -ea silentlycontinue | ? { $_.user -match ".*-(SEC|Data)-Email-.*$" }
            $smsg = "`nChecking Mailbox Permission on $($Tmbx.samaccountname) mailbox to accessing user:`n $($oSG.Name)...`n(blank if none)`n---`n$(($mbxp | select user,AccessRights,IsInhertied,Deny | format-list|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug


            # 8:10 AM 10/14/2015 AD SendAs too

            $mbxadp = $Tmbx | Get-ADPermission -domaincontroller $($InputSplat.domaincontroller) -ea Silentlycontinue | where { ($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and ($_.user -match ".*-(SEC|Data)-Email-.*$") };

            $smsg = "`nChecking AD SendAs Permission on $($Tmbx.samaccountname) mailbox to accessing user:`n $($oSG.Name)...`n(blank if none)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

            $smsg = "`n$(($mbxadp | select identity,User,ExtendedRights,Deny,Inherited | format-list|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
            # format-table -wrap ;

            $smsg = "`n---`nExisting $($oSG.Name) Membership...`n(blank if none)`n$((Get-ADGroupMember -identity $oSG.samaccountname -server $($DomainController) | select distinguishedName|out-string).trim())`n" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

            $smsg = "Testing Permissions Grant Update...`nAdd-MailboxPermission -whatif w`n$(($GrantSplat|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

            # 2:53 PM 5/18/2016 add retry code:
            $Exit = 0 ;
            # do loop until up to 4 retries...
            Do {
                Try {

                    add-mailboxpermission @GrantSplat -whatif ;
                    $Exit = $Retries ;
                } Catch {
                    $ErrTrapd = $Error[0] ;

                    Start-Sleep -Seconds $RetrySleep ;
                    $Exit ++ ;
                    $smsg = "Failed to exec add-mailboxpermission -whatif cmd because: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    $smsg = "Try #: $($Exit)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    If ($Exit -eq $Retries) { $smsg = "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ; } ;
                    # 11:15 AM 11/26/2019 add Cont - doesn't seem to be retrying
                    Continue ;
                } # try-E
            } Until ($Exit -eq $Retries) # loop-E

            $smsg = "Add-ADPermission -whatif... w`n$(($ADMbxGrantSplat|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

            $Exit = 0 ;
            Do {
                Try {
                    add-adpermission -identity $($TMbx.Identity) @ADMbxGrantSplat -whatif ;
                    $Exit = $Retries ;
                } Catch {
                    $ErrTrapd = $Error[0] ;
                    Start-Sleep -Seconds $RetrySleep ;
                    $Exit ++ ;
                    $smsg = "Failed to exec add-mailboxpermission -whatif cmd because: $($ErrTrpd)`nTry #: $($Exit)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ;
                    If ($Exit -eq $Retries) { $smsg = "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ; } ;
                    # 11:15 AM 11/26/2019 add Cont - doesn't seem to be retrying
                    Continue ;
                } # try-E
            } Until ($Exit -eq $Retries) # loop-E

            write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):Exec Permissions Grant Update";
            if ($whatif) {
                # 11:17 AM 6/22/2015 whatif-only pass
                write-verbose -verbose:$true "SKIPPING EXEC: Whatif-only pass";
            } else {
                write-host -foregroundcolor red "$((get-date).ToString("HH:mm:ss")):EXEC Add-MailboxPermission...";
                #add-mailboxpermission @GrantSplat ;

                $Exit = 0 ;
                # do loop until up to 4 retries...
                Do {
                    Try {

                        add-mailboxpermission @GrantSplat ;

                        $Exit = $Retries ;
                    } Catch {
                        $ErrTrapd = $Error[0] ;

                        Start-Sleep -Seconds $RetrySleep ;
                        $Exit ++ ;
                        $smsg = "Failed to exec add-mailboxpermission EXEC cmd because: $($ErrTrapd)`nTry #: $($Exit)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ; #Error|Warn|Debug
                        If ($Exit -eq $Retries) { $smsg = "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; } ;
                        # 11:15 AM 11/26/2019 add Cont - doesn't seem to be retrying
                        Continue ;
                    } # try-E
                } Until ($Exit -eq $Retries) # loop-E

                $smsg = "Add-ADPermission -whatif:identity $($TMbx.Identity) w`n$(($ADMbxGrantSplat|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                $Exit = 0 ;
                Do {
                    Try {
                        add-adpermission -identity $($TMbx.Identity) @ADMbxGrantSplat ;
                        $Exit = $Retries ;
                    } Catch {
                        $ErrTrapd = $Error[0] ;

                        Start-Sleep -Seconds $RetrySleep ;
                        $Exit ++ ;
                        $smsg = "Failed to exec add-adpermission EXEC cmd because: $($ErrTrapd)`nTry #: $($Exit)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ; #Error|Warn|Debug
                        If ($Exit -eq $Retries) { $smsg = "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; } ;
                        # 11:15 AM 11/26/2019 add Cont - doesn't seem to be retrying
                        Continue ;
                    } # try-E
                } Until ($Exit -eq $Retries) # loop-E

                # generics don't need this, test the OU path and only add folks below users
                # we're only hiding folks matching:
                #$rgxBannedOUs="^.*,OU=Disabled,OU=Users,.*OU=\w*,DC=(global|china),DC=ad,DC=toro,DC=com$" ;
                # and unhiding folks matching
                if ($Tmbx.distinguishedname -match $rgxUserOUs) {
                    # block that adds the $tmbx to the maintain-offboards.ps1 target AccGrant group for the region
                    $smsg = "Add TMBX $($tMbx.samaccountname) to AccGrant Group`n$(($TMbx | select -expand distinguishedname |?{$_ -match "DC=((global|china)),DC=ad,DC=toro,DC=com"}|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    switch ($matches[0]) {
                        "DC=global,DC=ad,DC=toro,DC=com" { $grpN = "LYN-DL-Exch-AGUnHide" ; } ;
                        "DC=china,DC=ad,DC=toro,DC=com" { $grpN = "XIA-DL-Exch-AGUnHide" ; } ;
                        default {
                            $smsg = "domain:NO MATCH!"; $smsg = "MESSAGE" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ;
                            EXIT ;
                        } ;
                    } ; # switch-E
                    $smsg = "==TGroup:$($grpN)";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                    if ($tdl = get-DistributionGroup -identity $grpN -domaincontroller $($InputSplat.domaincontroller) ) {
                        "==Add $($TMbx.name) to $($tdl.alias):" ;

                        $Exit = 0 ;
                        # do loop until up to 4 retries...
                        Do {
                            Try {
                                add-DistributionGroupMember -identity $tdl.alias -Member $TMbx.distinguishedname -domaincontroller $($InputSplat.domaincontroller) -whatif:$($whatif) ;

                                $Exit = $Retries ;
                            } Catch {
                                $ErrTrapd = $Error[0] ;

                                Start-Sleep -Seconds $RetrySleep ;
                                $Exit ++ ;
                                $smsg = "Failed to exec add-DistributionGroupMember EXEC cmd because: $($ErrTrapd)`nTry #: $($Exit)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ; #Error|Warn|Debug
                                If ($Exit -eq $Retries) { $smsg = "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; } ;
                                # 11:15 AM 11/26/2019 add Cont - doesn't seem to be retrying
                                Continue ;
                            } # try-E
                        } Until ($Exit -eq $Retries) # loop-E

                    } else {
                        "$($grpN): NOT FOUND" ;
                    }  ;
                } else {
                    $smsg = "TMBX $($tMbx.samaccountname) is in a non-User OU: Term Hide/Unhide groups do not apply...";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                }

            } ;
            write-verbose -verbose:$true "$(Get-Date -Format 'HH:mm:ss'):Waiting 5secs to refresh";
            Start-Sleep -s 5 ;

            # secgrp membership seldom comes through clean, add a refresh loop
            do {
                $smsg = "===REVIEW SETTINGS:===`n----Updated Permissions:`n`nChecking Mailbox/AD Permission on $($Tmbx.samaccountname) mailbox `n to accessing user:`n $($oSG.SamAccountName)`n---" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                # 12:52 PM 9/25/2017 what if we want it trimmed & common layout'd:
                $smsg = "`n$((get-mailboxpermission -identity $($TMbx.Identity) -user $(($oSG).Name) -domaincontroller $($InputSplat.domaincontroller) | ?{$_.user -match ".*-(SEC|Data)-Email-.*$"} | format-list user,AccessRights,IsInhertied,Deny |out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                $smsg = "`n$((get-mailboxpermission -identity $($TMbx.Identity) -user $(($oSG).Name) -domaincontroller $($InputSplat.domaincontroller) | ?{$_.user -match ".*-(SEC|Data)-Email-.*$"} | format-list user,AccessRights,IsInhertied,Deny|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                # 12:52 PM 9/25/2017 what if we want it trimmed & common layout'd:
                $smsg = "`n==User mbx grant: Confirming $($TMbx.name) member of $($grpN):" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                # 10:04 AM 11/22/2017 put the accgrant confirmation into the output:
                if ($Tmbx.distinguishedname -match $rgxUserOUs) {
                    $smsg = "$((Get-ADPermission -identity $($TMbx.Identity) -domaincontroller $($InputSplat.domaincontroller) -user "$($oSG.SamAccountName)"|  format-list User,ExtendedRights,Inherited,Deny | out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                } else {
                    $smsg = "TMBX $($tMbx.samaccountname) is in a non-User OU: Term Hide/Unhide groups do not apply...";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                }  ;

                $smsg = "`nUpdated $($oSG.Name) Membership...`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):---";
                if ($mbrs = Get-ADGroupMember -identity $oSG.samaccountname -server $($DomainController) | select distinguishedName ) {
                    $smsg = "$(($mbrs | out-string).trim() | out-default)`n-----------------------" ;
                } else {
                    $smsg = "(NO MEMBERS RETURNED)`n-----------------------" ;
                } ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                $bRet = Read-Host "Enter Y to Refresh Review (replication latency)." ;
            } until ($bRet -ne "Y");

        } else { $smsg = "$($InputSplat.SecGrpName) not found.`n" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } ; };


    } # PROC-E ;

    END {


    } # END-E
} #*----------------^ END Function add-MbxAccessGrant ^--------