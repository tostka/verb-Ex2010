#*------v new-MailboxGenericTOR.ps1 v------
function new-MailboxGenericTOR {

    <#
    .SYNOPSIS
    new-MailboxGenericTOR.ps1 - Wrapper/pre-processor function to create New shared Mbx (leverages new-MailboxShared generic function)
    .NOTES
    Version     : 1.1.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : http://twitter.com/tostka
    CreatedDate : 2021-05-11
    FileName    : new-MailboxGenericTOR.ps1
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Exchange,ExchangeOnPremises,Mailbox,Creation,Maintenance,UserMailbox
    REVISIONS
    * 3:55 PM 4/28/2026 revised -nooutput support to pass through
    * 8:45 AM 4/23/2026 spliced over latest begin/svc conn, no-trans block
    * 6:45 PM 4/9/2026 added temp code to force load this variant over mod copy
    * 2:24 PM 4/3/2026 added -CU5 registered test
    * 5:24 PM 1/28/2026 supress passstatusZ_tenorg error; Implement missing $SiteOverride passthrough ; REQUIRED FOR MIGRATIONS DOMAINS, THEY DON'T RESOLVE TO A FUNCTIONAL SITEOU (COMES BACK _MIGRATE)
    * 12:41 PM 1/27/2026 latest conn_svcs block updated
    * 2:48 PM 1/19/2026 -whatif's find ; bugfix: $pltCcOPSvcs.UserRole (postfilter, not match test)
    # 12:20 PM 1/16/2026 check out against AAD->MG migr mandate, bring in latest logging & SERVICE_CONNECTIONS blocks
    # 2:46 PM 1/24/2025 add support for $OfficeOverride = 'Pune, IN' ; support for Office that doesn't match SITE OU code: $OfficeOverride = 'Pune, IN' ; 
    # 1:15 PM 9/6/2023 updated CBH, pulled in expls from 7PSnMbxG/psb-PSnewMbxG.cbp. Works with current cba auth etc. 
    # 10:30 AM 10/13/2021 pulled [int] from $ticket , to permit non-numeric & multi-tix
    * 11:37 AM 9/16/2021 string
    * 8:55 AM 5/11/2021 functionalized into verb-ex2010 ; ren: internal helper func Cleanup() -> _cleanup() ; subbed some wv -v:v,=> wh (silly to use wv, w force display; should use it solely for optional verbose details
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
    #10:49 AM 2/11/2016: updated get-GCFast to current spec, updated any calls for "-site 'SITENAME'" to just default to local machine lookup
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
    new-MailboxGenericTOR.ps1 - Wrapper/pre-processor function to create New shared Mbx (leverages new-MailboxShared generic function)
    No service connectivity or module dependancies: Preprocesses the specified inputs into values suitable for the service-specific functions
    Typical intputs (can splat):
    ticket="355925";
    DisplayName="XXX Confirms" ;
    MInitial="" ;
    Owner="LOGON";
    BaseUser="AccountsReceivable";
    IsContractor=$false;
    NonGeneric=$true

    # splat entries
    ticket "355925";  DisplayName "XXX Confirms" ;  MInitial "" ;  Owner "LOGON";  BaseUser "AccountsReceivable";  IsContractor $false;  NonGeneric $true
    # equiv params use:
    -ticket "355925" -DisplayName "XXX Confirms"  -MInitial ""  -Owner "LOGON" -BaseUser "AccountsReceivable" -NonGeneric

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
    .PARAMETER TenOrg 
	TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']
	.PARAMETER Credential
	Credential to use for this connection [-credential [credential obj variable]")][System.Management.Automation.PSCredential]
	.PARAMETER UserRole
	Role of account (SID|CSID|UID|B2BI|CSVC|ESvc|LSvc)[-UserRole SID]
    .PARAMETER NoPrompt
    Suppress YYY confirmation prompts
    .PARAMETER Whatif
    Parameter to run a Test no-change pass, and log results [-Whatif switch]
    .PARAMETER ShowDebug
    Parameter to display Debugging messages [-ShowDebug switch]
    .PARAMETER NoOutput
    Switch to enable output (success/fail), defaults false, but adding to support tested function execution.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE 
    write-verbose "==SPLAT-DRIVEN, CREATE & IMMED GRANT & MIGRATE WRAPPER==" ; 
    write-verbose "For CU5 support, add to `$spltINPUTS hash: CU5='Exmark'" 
    write-verbose "CU5 OPTIONS: Exmark|Irritrol|IrritrolEurope|Lawn-boy|Lawngenie|Toro.be|TheToroCompany|Hayter|Toroused|EZLinkSupport|TheToroCo|Torodistributor|Dripirrigation|Toro.hu|Toro.co.uk|Torohosted|Uniquelighting|ToroExmark|RainMaster|Boss|perrotde|perrotpl"
    $whatIf=$true ;
    $spltINPUTS=@{
        ticket="TICKET" ;
        DisplayName="DNAME"  ;
        MInitial="" ;
        Owner="OWNER" ;
        showDebug=$true ;
        PermsDays=999 ;
        members="GRANTEE1,GRANTEE2";
    } ;
    if(!$dc){$dc=get-gcfast} ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):using common `$dc:$($dc)" ;
    $pltNmbx=[ordered]@{ ticket=$pltINPUTS.ticket ; DisplayName=$pltINPUTS.DisplayName  ; MInitial="" ; Owner=$pltINPUTS.Owner ; NonGeneric=$false  ; Vscan="YES" ; domaincontroller=$dc ; showDebug=$true ; whatIf=$($whatif) ; } ;
    if($pltINPUTS.Cu5){$pltNmbx.add("CU5",$pltINPUTS.CU5)} ;
    write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):`$pltINPUTS:`n$(($pltINPUTS|out-string).trim())" ;
    write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):new-MailboxGenericTOR w`n$(($pltNmbx|out-string).trim())" ;
    new-MailboxGenericTOR @pltNmbx;
    if(!$whatif){
        Do {write-host "." -NoNewLine; Start-Sleep -m (1000 * 5)} Until (($tmbx = get-mailbox $pltINPUTS.DisplayName -domaincontroller $dc )) ;
        if($tmbx){
            $pltGrant=[ordered]@{ ticket=$pltINPUTS.ticket  ; TargetID=$tmbx.samaccountname ; Owner=$pltINPUTS.Owner ; PermsDays=$pltINPUTS.PermsDays ; members=$pltINPUTS.members ; domaincontroller=$dc ; showDebug=$true ; whatIf=$whatif ; } ;
            write-host -foregroundcolor green "n$((get-date).ToString('HH:mm:ss')):add-MbxAccessGrant w`n$(($pltGrant|out-string).trim())" ;
            add-MbxAccessGrant @pltGrant ;
            caad ;
            write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):PREPARING DAWDLE LOOP!($($tmbx.PrimarySmtpAddress))`nAADLastSync:`n$((get-AADLastSync| ft -a TimeGMT,TimeLocal|out-string).trim())" ;
            Do {rxo ;  write-host "." -NoNewLine;  Start-Sleep -s 30} Until ((get-xorecipient $tmbx.PrimarySmtpAddress -EA 0)) ;
            write-host "`n*READY TO MOVE*!`a" ; sleep -s 1 ; write-host "*READY TO MOVE*!`a" ; sleep -s 1 ; write-host "*READY TO MOVE*!`a`n" ;
            write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):Running:`nmove-EXOmailboxNow.ps1 -TargetMailboxes $($tmbx.ALIAS) -showDebug -whatIf`n`n" ;
            . move-EXOmailboxNow.ps1 -TargetMailboxes $tmbx.ALIAS -showDebug -whatIf ;
            $strMoveCmd="move-EXOmailboxNow.ps1 -TargetMailboxes $($tmbx.ALIAS) -showDebug -NoTEST -whatIf:`$false" ;
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Move Command (copied to cb):`n`n$($strMoveCmd)`n" ;
            $strMoveCmd | out-clipboard ;
            $strCleanCmd="get-xomoverequest -BatchName ExoMoves-* | ?{`$_.status -eq 'Completed'} | Remove-xoMoveRequest -whatif" ;
            write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):Post-completion Cleanup Command :`n`n$($strCleanCmd)`n" ;
        } else { write-warning "No mbx found matching $($pltINPUTS.DisplayName). ABORTING"} ;
    } else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(WHATIF skipping AMPG & move)" } ;
    BP Scriptblock that pre-parses basic $spltINPUTS inputs and feeds new-MailboxGenericTOR.ps1, add-MbxAccessGrant.ps1 & move-EXOmailboxNow.ps1 (normally stored in psb-PSnewMbxG.cbp)
    .EXAMPLE
    write-verbose "==BULK CREATE WITH CU5 SUPPORT + MBPG==" ; 
    write-verbose "- For dnames from email address: use '!' where you want a space to appear (is collapsed out for emailaddr, expanded to space for dname)" ; 
    write-verbose "- Also where any Capitals appear in email addr, it will trigger replacement of any ! with space, and it will use the address dirname _as Capitalized_, for the final Dname.
    write-verbose "   Otherwise it replaces all underscores & periods in the dname with spaces, and converts to TitleCase, to create final Dname"; 
    write-verbose "- If put period (.) in dname, this will use it in new-sharedmailbox to split fname/lname to drive requested address (where explicitly asked for an email w a fname.lname period).
    write-verbose "    Other wise, the displayname is pushed into the LName of the mailbox" ; 
    write-verbose "`$mbxs specs an array of semicolon-delim'd data PER NEW MAILBOX: [email@domain.com];[FwdContactAddr];[CU5Spec];[OWNERUPN];[COMMA-DELIM'D GRANTEE ADDRESSES]"  ; 
    write-verbose "Note: The FWDCONTACTADDR value ISN'T USED IN INITIAL MBX CREATE (CAN BE USED MANUALLY IF RECYCLING THE SAME SAME ARRAY TO LATER CREATE MAILCONTACTS)" ; 
    $whatif=$true ;
    [array]$mbxs="ADDR1@toro.com; NOFWD@NOTHING.COM; TORO; OWNER1; GRANTEE1A@toro.com,GRANTEE1B@toro.com" ;
    $mbxs+="ADDR2@toro.com; NOFWD@NOTHING.COM; TORO; OWNER2; GRANTEE2A@toro.com,GRANTEE2B@toro.com" ;
    $ticket="TICKETNO" ;
    $SiteOverride="LYN" ;
    if(!$dc){$dc=get-gcfast} ;
    $moveTargets=@() ;
    $sQot = [char]34 ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):using common `$dc:$($dc)" ;
    $cultTxt=$((Get-Culture).TextInfo) ;
    $ttl = ($mbxs|measure).count ;
    $Procd=0 ;
    foreach($mbx in $mbxs) {
        $dname,$fwd,$cu5,$owner,$grantees= $mbx.split(';').trim() ;
        write-verbose "detect periods, go into email address" ;
        if($dname.split('@')[0].contains('.')){
            write-host "Dname contains periods, preserving for inclusion in new-MailboxShared as period-delimtied email addr" ;
            write-verbose "if Dname (emladdr) contains any caps: Split at @ & take 1st half as Dname, replacing any ! with a space (no other capitalization chgs are made to the specified emladdr string" ; 
            write-verbose "if Dname (emladdr) does *not* contain any caps, split at @, take 1st half as Dname, and recapitalize as TitleCase (Fname Lname)" ; 
            if($dname -cmatch '([A-Z])'){$dname = $dname.split("@")[0].replace('!',' ')} else {  $dname =  $cultTxt.ToTitleCase(($dname.split("@")[0].replace("_"," ").toLower())) } ;
        }else{ ;
            if($dname -cmatch '([A-Z])'){$dname = $dname.split("@")[0].replace('!',' ')} else {  $dname =  $cultTxt.ToTitleCase(($dname.split("@")[0].replace("_"," ").replace("."," ").toLower()))} ;
        } ;
        $sBnr="#*======v ($($Procd)/$($ttl)):$($dname) v======" ;
        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
        $pltNmbx=[ordered]@{  ticket=$ticket ; DisplayName="$($dname)"  ; MInitial="" ; Owner=$owner ; SiteOverride=$SiteOverride ; NonGeneric=$false  ; Vscan="YES" ; NoPrompt=$true ; domaincontroller=$dc ; showDebug=$true ; whatIf=$($whatif) ; } ;
        if($cu5 -AND ($cu5 -ne "toro")){
            write-host -fore yellow "CU5:$($cu5): CUSTOM DOMAIN SPEC" ;
            $pltNmbx.add("CU5",$CU5) ;
        } ;
        write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):new-MailboxGenericTOR w`n$(($pltNmbx|out-string).trim())" ;
        new-MailboxGenericTOR @pltNmbx ;
        if(!($whatif)){
            write-host "waiting 10 secs..." ;
            start-sleep -seconds 10 ;
            Do {write-host "." -NoNewLine; Start-Sleep -m (1000 * 5)} Until (($tmbx = get-mailbox "$($dname)" -domaincontroller $dc -ea 0)) ;
            $pltGrant=[ordered]@{  ticket=$ticket  ; TargetID=$tmbx.samaccountname ; Owner=$owner ; PermsDays=999 ; members=$grantees ; NoPrompt=$true ; domaincontroller=$dc ; showDebug=$true  ; whatIf=$whatif ; } ;
            write-host -foregroundcolor green "n$((get-date).ToString('HH:mm:ss')):===add-MbxAccessGrant w`n$(($pltGrant|out-string).trim())" ;
            add-MbxAccessGrant @pltGrant ;
            $moveTargets+= $tmbx.alias ;
        } else {write-host -foregroundcolor green "(-whatif, skipping acc grant)" };
        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
    } ;
    if(!($whatif)){
        write-host -foregroundcolor green "===$((get-date).ToString('HH:mm:ss')):CONFIRMING PERMISSIONS:" ;
        foreach($movetarget in $movetargets) {
            write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):`Alias:$($movetarget):" ;
            get-mailboxpermission -identity "$($movetarget)" |?{$_.user -like 'toro*'}| select user;
        } ;
    } ;
    if(!($whatif) -AND $tmbx){
        caad ;
        write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):PREPARING DAWDLE LOOP!($($tmbx.PrimarySmtpAddress))`nAADLastSync:`n$((get-AADLastSync| ft -a TimeGMT,TimeLocal|out-string).trim())" ;
        Do {rxo ; write-host "." -NoNewLine; Start-Sleep -s 30} Until ((get-xorecipient $tmbx.PrimarySmtpAddress -EA 0)) ;
        write-host "`n*READY TO MOVE*!`a" ; sleep -s 1 ; write-host "*READY TO MOVE*!`a" ; sleep -s 1 ; write-host "*READY TO MOVE*!`a`n" ;
        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Running:`n`nmove-EXOmailboxNow.ps1 -TargetMailboxes $($sQot + ($moveTargets -join '","') + $sQot) -showDebug -whatIf`n`n" ;
        . move-EXOmailboxNow.ps1 -TargetMailboxes $moveTargets -showDebug -whatIf ;
        $strMoveCmd="move-EXOmailboxNow.ps1 -TargetMailboxes `$moveTargets -showDebug -NoTEST  -whatIf:`$false" ;
        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Move Command (copied to cb):`n`n$($strMoveCmd)`n" ;
        $strMoveCmd | out-clipboard ;
        $strCleanCmd="get-xomoverequest -BatchName ExoMoves-* | ?{`$_.status -eq 'Completed'} | Remove-xoMoveRequest -whatif" ;
        write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):Post-completion Cleanup Command :`n`n$($strCleanCmd)`n" ;
    } else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(WHATIF skipping AMPG & move)" } ;
    BP Scriptblock that permits bulk creation of a series of mailboxes, with explicit email addresses. DisplayName is generated as a variant of the email address (see verbose comments above for details). As per the prior expl, also waits for ADC repliction to complete, and then mocks up mailbox move to cloud.
    .EXAMPLE
    write-verbose "==MAILBOXES ARRAY AD-HOC REPORT (runs for each address in the $tmbxs array)==" ; 
    $tmbxs="MBX1@toro.com","MBX2@toro.com" ;
    foreach($tmbx in $tmbxs){
        write-host  "==$($tmbx)" ;
        $mbxo = get-mailbox -Identity $tmbx  ;
        $cmbxo= Get-CASMailbox -Identity $mbxo.samaccountname ;
        $aduprops="GivenName,Surname,Manager,Company,Office,Title,StreetAddress,City,StateOrProvince,c,co,countryCode,PostalCode,Phone,Fax,Description" ;
        $ADu = get-ADuser -Identity $mbxo.samaccountname-properties * | select *;
        write-host -foregroundcolor green "User Email:`t$(($mbxo.WindowsEmailAddress.tostring()).trim())" ;
        write-host -foregroundcolor green "Mailbox Information:" ;
        write-host -foregroundcolor green "$(($mbxo | select @{Name='LogonName';
        Expression={$_.SamAccountName }},Name,DisplayName,Alias,database,UserPrincipalName,RetentionPolicy,CustomAttribute5,CustomAttribute9,RecipientType,RecipientTypeDetails | out-string).trim())" ;
        write-host -foregroundcolor green "$(($Adu | select GivenName,Surname,Manager,Company,Office,Title,StreetAddress,City,StateOrProvince,c,co,countryCode,PostalCode,Phone,Fax,Description | out-string).trim())";
        write-host -foregroundcolor green "ActiveSyncMailboxPolicy:$($cmbxo.ActiveSyncMailboxPolicy.tostring())" ;
        write-host -foregroundcolor green "Description: $($Adu.Description.tostring())";
        write-host -foregroundcolor green "Info: $($Adu.info.tostring())";
        write-host -foregroundcolor green "Initial Password: $(($pltINPUTS.pass | out-string).trim())" ;
        $tmbx=$mbxo=$cmbxo=$aduprops=$ADu=$null;
        write-host "===========" ;
    } ;
    BP Scriptblock that outputs a summary report for each mailbox in the array (output resembles the output for the new-MailboxGenericTOR function)
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
    ## hosted w/in verb-Ex2010, 
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
    	[Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
	        [ValidateNotNullOrEmpty()]
	        $TenOrg = 'TOR',
	    [Parameter(HelpMessage="Credential to use for cloud actions [-credential [credential obj variable]")][System.Management.Automation.PSCredential]
	        $Credential,
	    [Parameter(Mandatory = $false, HelpMessage = "Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]")]
            [ValidateSet('SID','CSID','UID','B2BI','CSVC')]
	        [string]$UserRole='SID',
        [Parameter(HelpMessage="Suppress YYY confirmation prompts [-NoPrompt]")]
            [switch] $NoPrompt,
        [Parameter(HelpMessage='Debugging Flag [$switch]')]
            [switch] $showDebug,
        [Parameter(HelpMessage='Whatif Flag [$switch]')]
            [switch] $whatIf,
        [Parameter(HelpMessage='NoOutput Flag [$switch]')]
            [switch] $NoOutput=$true
    ) ;
    
    BEGIN {
        #region ENVIRO_DISCOVER_NODEP ; #*------v ENVIRO_DISCOVER_NODEP v------
        $Verbose = [boolean]($VerbosePreference -eq 'Continue') ;
        # USE: -whatif:$($whatifswitch) # proxy for -whatif when SupportsShouldProcess
        [boolean]$whatIfSwitch = ($WhatIf.IsPresent -or $whatif -eq $true -OR $WhatIfPreference -eq $true); $smsg = "-Verbose:$($Verbose)`t-Whatif:$($whatifswitch) " ; write-host -foregroundcolor yellow $smsg
        $rPSCmdlet = $PSCmdlet ; # an object that represents the cmdlet or advanced function that's being run. Available on functions w CmdletBinding (& $args will not be available). (Blank on non-CmdletBinding/Non-Adv funcs).
        $rPSScriptRoot = $PSScriptRoot ; # the full path of the executing script's parent directory., PS2: valid only in script modules (.psm1). PS3+:it's valid in all scripts. (Funcs: ParentDir of the file that hosts the func)
        $rPSCommandPath = $PSCommandPath ; # the full path and filename of the script that's being run, or file hosting the funct. Valid in all scripts.
        $rMyInvocation = $MyInvocation ; # populated only for scripts, function, and script blocks.
        # - $MyInvocation.MyCommand.Name returns name of a function, to identify the current command,  name of the current script (pop'd w func name, on Advfuncs)
        # - Ps3+:$MyInvocation.PSScriptRoot : full path to the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        # - Ps3+:$MyInvocation.PSCommandPath : full path and filename of the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        #     ** note: above pair contain information about the _invoker or calling script_, not the current script
        $rPSBoundParameters = $PSBoundParameters ;
        #endregion ENVIRO_DISCOVER_NODEP ; #*------^ END ENVIRO_DISCOVER_NODEP ^------
        #region RV_TRANSCRIPT_NODEP ; #*------v RV_TRANSCRIPT_NODEP v------
        if($rMyInvocation.mycommand.commandtype -eq 'ExternalScript'){
            $cmdType = 'Script' ; $isScript = $true ; $isfunc = $false ; $isFuncAdv = $false  ;
            if($rMyInvocation.InvocationName){
                if($rMyInvocation.InvocationName -AND $rMyInvocation.InvocationName -ne '.'){
                    $CmdPathed = (resolve-path $rMyInvocation.InvocationName -ea STOP).path ;
                }elseif($rMyInvocation.mycommand.definition -and (test-path -path $rMyInvocation.mycommand.definition -pathtype Leaf)){
                    $CmdPathed = (resolve-path $rMyInvocation.mycommand.definition -ea STOP).path ;
                }
                $CmdName= split-path $CmdPathed -ea 0 -leaf ;
                $CmdParentDir = split-path $CmdPathed -ea 0 ;
                $CmdParentPathedExt = $null ; 
                $CmdNameNoExt = [system.io.path]::GetFilenameWithoutExtension($CmdPathed) ;                
            } else{
                throw "Unpopulated dependant:`$rMyInvocation.InvocationName!" ;
            } ;
        }elseif($rMyInvocation.mycommand.commandtype -eq 'Function'){
            $cmdType = 'Function' ; $isScript = $false ; $isfunc = $true ; $isFuncAdv = $false  ;
            if($rMyInvocation.mycommand.name){
                $CmdName = $rMyInvocation.mycommand.name ;
                $CmdParentPathedExt = [regex]::match($CmdName,'(\.\w+$)').value ; 
                $CmdNameNoExt = [system.io.path]::GetFilenameWithoutExtension($CmdName) ;
            } ;
            if($rMyInvocation.ScriptName){
                if($CmdParentPathed = (resolve-path -path $rMyInvocation.ScriptName).path){
                    $CmdParentDir = split-path $CmdParentPathed ;
                    write-verbose "Hosted function, mock it up as a proxy path for transcription name" ;
                    $CmdParentPathedExt = [regex]::match($CmdParentPathed,'(\.\w+$)').value
                    $CmdPathed = (join-path -Path $CmdParentDir -ChildPath "$($CmdName)$($CmdParentPathedExt)")
                } else{
                    throw "emtpy `$rMyInvocation.ScriptName!, unable to calculate isFunc: `$CmdParentDir!" ;
                }
            }elseif($rPSCmdlet.MyInvocation.mycommand.commandtype -eq 'Function' -AND $rPSCmdlet.MyInvocation.mycommand.modulename -AND $rPSCmdlet.MyInvocation.mycommand.Module.path){
                # cover function, pathed into the Module.path
                if($CmdParentPathed = (resolve-path -path $rPSCmdlet.MyInvocation.mycommand.Module.path).path){
                    $CmdParentDir = split-path $CmdParentPathed ;
                    write-verbose "Hosted function, mock it up as a proxy path for transcription name" ;
                    $CmdParentPathedExt = [regex]::match($CmdParentPathed,'(\.\w+$)').value
                    $CmdPathed = (join-path -Path $CmdParentDir -ChildPath "$($CmdName)$($CmdParentPathedExt)")
                } else{
                    throw "emtpy `$rMyInvocation.ScriptName!, unable to calculate isFunc: `$CmdParentDir!" ;
                }
            } ;
            if($isFunc -AND ((gv rPSCmdlet -ea 0).value -eq $null)){
                $isFuncAdv = $false
            } elseif($isFunc) {
                $isFuncAdv = [boolean]($isFunc -AND $rMyInvocation.InvocationName -AND ($CmdName -eq $rMyInvocation.InvocationName))
            }            
        } else{
            throw "unrecognized environment combo: unable to resolve Script v Function v FuncAdv!" ;
        } ;
        $smsg = $null ; 
        if($isScript){$smsg += "`n`$isScript : $(($isScript|out-string).trim())" ;}
        if($isfunc){$smsg += "`n`$isfunc : $(($isfunc|out-string).trim())"} ;
        if($isFuncAdv){$smsg += "`n`$isFuncAdv : $(($isFuncAdv|out-string).trim())"} ;    
        $smsg += "`n`$CmdName  : $(($CmdName|out-string).trim())" ;
        $smsg += "`n`$CmdPathed  : $(($CmdPathed|out-string).trim())" ;
        $smsg += "`n`$CmdParentDir  : $(($CmdParentDir|out-string).trim())" ;
        if($CmdParentPathedExt){$smsg += "`n`$CmdParentPathedExt  : $(($CmdParentPathedExt|out-string).trim())"};
        $smsg += "`n`$CmdNameNoExt  : $(($CmdNameNoExt|out-string).trim())" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
        if($CmdParentDir -AND $CmdNameNoExt){
            $transcript = (join-path -path $CmdParentDir -ChildPath 'logs') ;
            if(-not (test-path $transcript  -PathType Container -ea 0)){ mkdir $transcript -verbose } ;
            $transcript = join-path -path $transcript -childpath $CmdNameNoExt ;
            #if($ticket){$transcript += "-$($ticket)" }
            #if($whatif -OR $WhatIf.IsPresent){$transcript += "-WHATIF"}ELSE{$transcript += "-EXEC"} ;
            #if($thisXoMbx.userprincipalname){$transcript += "-$($thisXoMbx.userprincipalname)"} ;
            #$transcript += "-LASTPASS-trans-log.txt" ;
            $transcript += "-$(get-date -format 'yyyyMMdd-HHmmtt')-trans-log.txt" ;
            write-verbose "RV_TRANSCRIPT: `$transcript: $($transcript)" ; 
        }else{ write-warning "INSUFFICIENT ENVIRONMENT INPUTS TO AUTOBUILD A STRANSCRIPT!" } ; 
        #endregion RV_TRANSCRIPT_NODEP ; #*------^ END RV_TRANSCRIPT_NODEP ^------
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
        
        $dbgDate = '4/27/2026'; # debugging ipmo force loads variants not in modules

        #  add password generator
        [Reflection.Assembly]::LoadWithPartialName("System.Web") | out-null ;

        # add retry support
        $Retries = 4 ; # number of re-attempts
        $RetrySleep = 5 ; # seconds to wait between retries
        # add CU5 validator 
        #$rgxCU5 = [moved to infra file]     
    
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
    } ;  # BEGIN-E
    PROCESS{

        #$logging = $True ; # need to set in scope outside of functions
        $pltInput=[ordered]@{
            Verbose = $($VerbosePreference -eq 'Continue') ; 
        } ;

        if($PSCommandPath){$pltInput.add("ParentPath",$PSCommandPath) } ; 
        if($DisplayName){$pltInput.add("DisplayName",$DisplayName) } ;
        if($MInitial){$pltInput.add("MInitial",$MInitial) } ;
        if($Owner){$pltInput.add("Owner",$Owner) } ;
        if($BaseUser){$pltInput.add("BaseUser",$BaseUser) } ;
        if($IsContractor){$pltInput.add("IsContractor",$IsContractor) } ;
        # add room/equip support
        if($Room){$pltInput.add("Room",$Room) } ;
        if($Equip){$pltInput.add("Equip",$Equip) } ;
        if($Ticket){$pltInput.add("Ticket",$Ticket) } ;
        if($domaincontroller){$pltInput.add("domaincontroller",$domaincontroller) } ;
        if($NoPrompt){$pltInput.add("NoPrompt",$NoPrompt) } ;
        if($showDebug){$pltInput.add("showDebug",$showDebug) } ;
        if($verbose){
            if($pltInput.keys -contains 'verbose'){
                $pltInput.verbose = $($VerbosePreference -eq "Continue") ; 
            }else{
                $pltInput.add("verbose",$(($VerbosePreference -eq "Continue"))) } ;
            }
        #if($whatIf){$pltInput.add("whatIf",$whatIf) } ;
        if(get-variable -name whatif){$pltInput.add("whatIf",$whatIf) } ;
        # 2:59 PM 1/24/2025 new OfficeOverride
        if($OfficeOverride){
            if($pltInput.keys -contains 'OfficeOverride'){
                $pltInput.OfficeOverride=$OfficeOverride ; 
            }else{
                $pltInput.add('OfficeOverride',$OfficeOverride) ;  
            }
        };
        # *5:17 PM 1/28/2026 unimplemented siteoverride!
        if($SiteOverride){
            if($pltInput.keys -contains 'SiteOverride'){
                $pltInput.SiteOverride=$SiteOverride ; 
            }else{
                $pltInput.add('SiteOverride',$SiteOverride) ;  
            }
        };
        # only reset from defaults on explicit -NonGeneric $true param
        if($NonGeneric -eq $true){
            # switching over generics to real 'shared' mbxs: "Shared" = $True
        } else {
            # force it if not true
            $NonGeneric=$false;
            #$pltInput.NonGeneric=$false
            # rem'd above in favor of below
        } ;
        if($NonGeneric){$pltInput.add("NonGeneric",$NonGeneric) } ;

        # vscan
        if ($Vscan){
            if ($Vscan -match "(?i:^(YES|NO)$)" ) {
                $Vscan = $Vscan.ToString().ToUpper() ;
                if($Vscan){$pltInput.add("Vscan",$Vscan) } ;
            } else {
                $Vscan = $null ;
                #$pltInput.Vscan=$Vscan ;
                # force em  on all, no reason not to have external email!
                if($Vscan){$pltInput.add("Vscan","YES") } ;
            }  ;
        }; # If not explicit yes/no, prompt for input

        # Cu5 override support (normally inherits from assigned owner/manager)
        if ($Cu5){
            # CONFIRM IT'S SUPPORTED TAG:
            TRY{
                $eaps = get-emailaddresspolicy -ea STOP ; 
                $CU5Supported = $eaps |foreach-object{
                    if($_.recipientfilter -match "\(CustomAttribute5\s-eq\s'([\w\.]+)'\)"){
                        $matches[1] | write-output  ; 
                    }
                }
                if($CU5Supported){
                    if($CU5Supported -contains $cu5){
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):-CU5:$($CU5) supported" ; 
                    }else{
                        $smsg = "-CU5:$($CU5) NOT SUPPORTED" ; 
                        $smsg += "`n$(($eaps|?{$_.recipientfilter -match 'CustomAttribute5'}|ft -a name,recipientfilter |out-string).trim())" ; 
                        write-warning $smsg ;
                        BREAK ; 
                    } ; 
                } ; 
            } CATCH {$ErrTrapd=$Error[0] ;
               write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
               $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
               write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
             } ;
            
            #$pltInput.Cu5=$Cu5;
            # looks like it's adding on assign (?.?)7
            if($Cu5){$pltInput.add("Cu5",$Cu5) } ;
        } else {
            $pltInput.add("Cu5",$null) ;
        }  ;
    
        $tCmdlet = 'new-MailboxShared' ; $BMod = 'VERB-ex2010' ; 
        if($psISE -AND ((get-date ).tostring() -match $dbgDate)){ 
            if((gcm $tCmdlet).source -eq $BMod){
                Do{
                    gci "D:\scripts\$($tCmdlet)_func.ps1" -ea STOP | ipmo -fo -verb  ;
                }until((gcm $tCmdlet).source -ne $BMod)
            } ;
        } ; 
        #if($NoOutput){
        if($pltInput.keys -contains 'NoOutput'){
            $pltInput.NoOutput=[boolean]($NoOutput) ; 
        }else{
            $pltInput.add('NoOutput',[boolean]($NoOutput)) ;  
        }
        #};
        if(-not($NoOutput)){
            $bRet = new-MailboxShared @pltInput ; 
            $bRet | write-output ;
        } else { 
            new-MailboxShared @pltInput
        } ; 
        # shift exit work into _cleanup function ;
        #_cleanup ;
        #Exit ;

        #*======^ END SUB MAIN ^======
        #endregion SUBMAIN ; # ------
    }  # PROC-E
}

#*------^ new-MailboxGenericTOR.ps1 ^------
