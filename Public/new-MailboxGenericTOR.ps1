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
    
    BEGIN{
        $Verbose = [boolean]($VerbosePreference -eq 'Continue') ; 

        #region FUNCTIONS_INTERNAL ; #*======v FUNCTIONS_INTERNAL v======
        # Pull the CUser mod dir out of psmodpaths:
        #$CUModPath = $env:psmodulepath.split(';')|?{$_ -like '*\Users\*'} ;
    
        # 2b4() 2b4c() & fb4() are located up in the CONSTANTS_AND_ENVIRO\ENCODED_CONTANTS block ( to convert Constant assignement strings)

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

        #region CONSTANTS_AND_ENVIRO ; #*======v CONSTANTS_AND_ENVIRO v======
        #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
        push-TLSLatest
        $Verbose = [boolean]($VerbosePreference -eq 'Continue') ; 
        $rPSCmdlet = $PSCmdlet ; # an object that represents the cmdlet or advanced function that's being run. Available on functions w CmdletBinding (& $args will not be available). (Blank on non-CmdletBinding/Non-Adv funcs).
        $rPSScriptRoot = $PSScriptRoot ; # the full path of the executing script's parent directory., PS2: valid only in script modules (.psm1). PS3+:it's valid in all scripts. (Funcs: ParentDir of the file that hosts the func)
        $rPSCommandPath = $PSCommandPath ; # the full path and filename of the script that's being run, or file hosting the funct. Valid in all scripts.
        $rMyInvocation = $MyInvocation ; # populated only for scripts, function, and script blocks.
        # - $MyInvocation.MyCommand.Name returns name of a function, to identify the current command,  name of the current script (pop'd w func name, on Advfuncs)
        # - Ps3+:$MyInvocation.PSScriptRoot : full path to the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        # - Ps3+:$MyInvocation.PSCommandPath : full path and filename of the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        #     ** note: above pair contain information about the _invoker or calling script_, not the current script
        $rPSBoundParameters = $PSBoundParameters ; 
        #region PREF_VARI_DUMP ; #*------v PREF_VARI_DUMP v------
        <#$script:prefVaris = @{
            whatifIsPresent = $whatif.IsPresent
            whatifPSBoundParametersContains = $rPSBoundParameters.ContainsKey('WhatIf') ; 
            whatifPSBoundParameters = $rPSBoundParameters['WhatIf'] ;
            WhatIfPreferenceIsPresent = $WhatIfPreference.IsPresent ; # -eq $true
            WhatIfPreferenceValue = $WhatIfPreference;
            WhatIfPreferenceParentScopeValue = (Get-Variable WhatIfPreference -Scope 1).Value ;
            ConfirmPSBoundParametersContains = $rPSBoundParameters.ContainsKey('Confirm') ; 
            ConfirmPSBoundParameters = $rPSBoundParameters['Confirm'];
            ConfirmPreferenceIsPresent = $ConfirmPreference.IsPresent ; # -eq $true
            ConfirmPreferenceValue = $ConfirmPreference ;
            ConfirmPreferenceParentScopeValue = (Get-Variable ConfirmPreference -Scope 1).Value ; 
            VerbosePSBoundParametersContains = $rPSBoundParameters.ContainsKey('Confirm') ; 
            VerbosePSBoundParameters = $rPSBoundParameters['Verbose'] ;
            VerbosePreferenceIsPresent = $VerbosePreference.IsPresent ; # -eq $true
            VerbosePreferenceValue = $VerbosePreference ;
            VerbosePreferenceParentScopeValue = (Get-Variable VerbosePreference -Scope 1).Value;
            VerboseMyInvContains = '-Verbose' -in $rPSBoundParameters.UnboundArguments ; 
            VerbosePSBoundParametersUnboundArgumentContains = '-Verbose' -in $rPSBoundParameters.UnboundArguments 
        } ;
        write-verbose "`n$(($script:prefVaris.GetEnumerator() | Sort-Object Key | Format-Table Key,Value -AutoSize|out-string).trim())`n" ; 
        #>
        #endregion PREF_VARI_DUMP ; #*------^ END PREF_VARI_DUMP ^------
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
        $mts = $pltRVEnv.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltRVEnv.remove($_.Name)} ; rv mts -ea 0 -whatif:$false -confirm:$false; 
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
        #endregion RV_ENVIRO ; #*------^ END RV_ENVIRO ^------
        #region NETWORK_INFO ; #*======v NETWORK_INFO v======
        if(get-command resolve-NetworkLocalTDO  -ea STOP){}ELSE{
            $smsg = "UNABLE TO gcm resolve-NetworkLocalTDO !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            BREAK ; 
        } ; 
        $netsettings = resolve-NetworkLocalTDO ; 
        if($env:Userdomain){ 
            switch($env:Userdomain){
                'CMW'{
                    #$logon_SID = $CMW_logon_SID 
                }
                'TORO'{
                    #$o365_SIDUpn = $o365_Toroco_SIDUpn ; 
                    #$logon_SID = $TOR_logon_SID ; 
                }
                $env:COMPUTERNAME{
                    $smsg = "%USERDOMAIN% -EQ %COMPUTERNAME%: $($env:computername) => non-domain-connected, likely edge role Ex server!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    if($netsettings.Workgroup){
                        $smsg = "WorkgroupName:$($netsettings.Workgroup)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;                    
                    } ; 
                } ; 
                default{
                    $smsg = "$($env:userdomain):UNRECOGIZED/UNCONFIGURED USER DOMAIN STRING!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    THROW $SMSG 
                    BREAK ; 
                }
            } ; 
        } ;  # $env:Userdomain-E
        #endregion NETWORK_INFO ; #*======^ END NETWORK_INFO ^======
        #region OS_INFO ; #*------v OS_INFO v------
        <# os detect, covers Server 2016, 2008 R2, Windows 10, 11
        if (get-command get-ciminstance -ea 0) {$OS = (Get-ciminstance -class Win32_OperatingSystem)} else {$Os = Get-WMIObject -class Win32_OperatingSystem } ;
        #$isWorkstationOS = $isServerOS = $isW2010 = $isW2011 = $isS2016 = $isS2008R2 = $false ;
        write-host "Detected:`$Os.Name:$($OS.name)`n`$Os.Version:$($Os.Version)" ;
        if ($OS.name -match 'Microsoft\sWindows\sServer') {
            $isServerOS = $true ;
            if ($os.name -match 'Microsoft\sWindows\sServer\s2016'){$isS2016 = $true ;} ;
            if ($os.name -match 'Microsoft\sWindows\sServer\s2008\sR2') { $isS2008R2 = $true ; } ;
        } else { 
            if ($os.name -match '^Microsoft\sWindows\s11') {
                $isWorkstationOS = $true ;
                if ($os.name -match 'Microsoft\sWindows\s11') { $isW2011 = $true ; } ;
            } elseif ($os.name -match '^Microsoft\sWindows\s10') {
                $isWorkstationOS = $true ; $isW2010 = $true
            } else {
                $isWorkstationOS = $true ;
            } ;         
        } ; 
        #>
        #endregion OS_INFO ; #*------^ END OS_INFO ^------
        #region TEST_EXOPLOCAL ; #*------v TEST_EXOPLOCAL v------
        if(get-command test-LocalExchangeInfoTDO -ea STOP){}ELSE{
            $smsg = "UNABLE TO gcm test-LocalExchangeInfoTDO !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            BREAK ; 
        } ; 
        $lclExOP = test-LocalExchangeInfoTDO ; 
        write-verbose "Expand returned NoteProperty properties into matching local variables" ; 
        if($host.version.major -gt 2){
            $lclExOP.PsObject.Properties | ?{$_.membertype -eq 'NoteProperty'} | foreach-object{set-variable -name $_.name -value $_.value -verbose -whatif:$false -Confirm:$false ;} ;
        }else{
            write-verbose "Psv2 lacks the above expansion capability; just create simpler variable set" ; 
            $ExVers = $lclExOP.ExVers ; $isLocalExchangeServer = $lclExOP.isLocalExchangeServer ; $IsEdgeTransport = $lclExOP.IsEdgeTransport ;
        } ;
        #
        #endregion TEST_EXOPLOCAL ; #*------^ END TEST_EXOPLOCAL ^------

        <#
        #region PsParams ; #*------v PSPARAMS v------
        $PSParameters = New-Object -TypeName PSObject -Property $rPSBoundParameters ;
        # DIFFERENCES $PSParameters vs $PSBoundParameters:
        # - $PSBoundParameters: System.Management.Automation.PSBoundParametersDictionary (native obj)
        # test/access: ($PSBoundParameters['Verbose'] -eq $true) ; $PSBoundParameters.ContainsKey('Referrer') #hash syntax
        # CAN use as a @PSBoundParameters splat to push through (make sure populated, can fail if wrong type of wrapping code)
        # - $PSParameters: System.Management.Automation.PSCustomObject (created obj)
        # test/access: ($PSParameters.verbose -eq $true) ; $PSParameters.psobject.Properties.name -contains 'SenderAddress' ; # cobj syntax
        # CANNOT use as a @splat to push through (it's a cobj)
        write-verbose "`$rPSBoundParameters:`n$(($rPSBoundParameters|out-string).trim())" ;
        # pre psv2, no $rPSBoundParameters autovari to check, so back them out:
        #>
        <# recycling $rPSBoundParameters into @splat calls: (can't use $psParams, it's a cobj, not a hash!)
        # rgx for filtering $rPSBoundParameters for params to pass on in recursive calls (excludes keys matching below)
        $rgxBoundParamsExcl = '^(Name|RawOutput|Server|Referrer)$' ; 
        if($rPSBoundParameters){
                $pltRvSPFRec = [ordered]@{} ;
                # add the specific Name for this call, and Server spec (which defaults, is generally not 
                $pltRvSPFRec.add('Name',"$RedirectRecord" ) ;
                $pltRvSPFRec.add('Referrer',$Name) ; 
                $pltRvSPFRec.add('Server',$Server ) ;
                $rPSBoundParameters.GetEnumerator() | ?{ $_.key -notmatch $rgxBoundParamsExcl} | foreach-object { $pltRvSPFRec.add($_.key,$_.value)  } ;
                write-host "Resolve-SPFRecord w`n$(($pltRvSPFRec|out-string).trim())" ;
                Resolve-SPFRecord @pltRvSPFRec  | write-output ;
        } else {
            $smsg = "unpopulated `$rPSBoundParameters!" ;
            write-warning $smsg ;
            throw $smsg ;
        };     
        #>
        #endregion PsParams ; #*------^ END PSPARAMS ^------    
        #endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------

        #region COMMON_CONSTANTS ; #*------v COMMON_CONSTANTS v------
    
        if(-not $DoRetries){$DoRetries = 4 } ;    # # times to repeat retry attempts
        if(-not $RetrySleep){$RetrySleep = 10 } ; # wait time between retries
        if(-not $RetrySleep){$DawdleWait = 30 } ; # wait time (secs) between dawdle checks
        if(-not $DirSyncInterval){$DirSyncInterval = 30 } ; # AADConnect dirsync interval
        if(-not $ThrottleMs){$ThrottleMs = 50 ;}
        if(-not $rgxDriveBanChars){$rgxDriveBanChars = '[;~/\\\.:]' ; } ; # ;~/\.:,
        if(-not $rgxCertThumbprint){$rgxCertThumbprint = '[0-9a-fA-F]{40}' } ; # if it's a 40char hex string -> cert thumbprint  
        if(-not $rgxSmtpAddr){$rgxSmtpAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ; } ; # email addr/UPN
        if(-not $rgxDomainLogon){$rgxDomainLogon = '^[a-zA-Z][a-zA-Z0-9\-\.]{0,61}[a-zA-Z]\\\w[\w\.\- ]+$' } ; # DOMAIN\samaccountname 
        if(-not $exoMbxGraceDays){$exoMbxGraceDays = 30} ; 
        if(-not $XOConnectionUri ){$XOConnectionUri = 'https://outlook.office365.com'} ; 
        if(-not $SCConnectionUri){$SCConnectionUri = 'https://ps.compliance.protection.outlook.com'} ; 
        if(-not $XODefaultPrefix){$XODefaultPrefix = 'xo' };
        if(-not $SCDefaultPrefix){$SCDefaultPrefix = 'sc' };
        #$rgxADDistNameGAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 1 ) -join ',')" 
        #$rgxADDistNameAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 2 ) -join ',')"

        write-verbose "Coerce configured but blank Resultsize to Unlimited" ; 
        if(get-variable -name resultsize -ea 0){
            if( ($null -eq $ResultSize) -OR ('' -eq $ResultSize) ){$ResultSize = 'unlimited' }
            elseif($Resultsize -is [int]){} else {throw "Resultsize must be an integer or the string 'unlimited' (or blank)"} ;
        } ; 
        #$ComputerName = $env:COMPUTERNAME ;
        #$NoProf = [bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
        # XXXMeta derived constants:
        # - MGU Licensing group checks
        # calc the rgxLicGrpName fr the existing $xxxmeta.rgxLicGrpDN: (get-variable tormeta).value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
        #$rgxLicGrpName = (get-variable -name "$($tenorg)meta").value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
        # use the dn vers LicGrouppDN = $null ; # | ?{$_ -match $tormeta.rgxLicGrpDN}
        #$rgxLicGrpDN = (get-variable -name "$($tenorg)meta").value.rgxLicGrpDN
        # email trigger vari, it will be semi-delimd list of mail-triggering events
        $script:PassStatus = $null ;
        # TenOrg or other looped-specific PassStatus (auto supported by 7pswlt)
        #New-Variable -Name PassStatus_$($tenorg) -scope Script -Value $null ;
        [array]$SmtpAttachment = $null ;
        #write-verbose "start-Timer:Master" ; 
        $swM = [Diagnostics.Stopwatch]::StartNew() ;
        # $ByPassLocalExchangeServerTest = $true # rough in, code exists below for exempting service/regkey testing on this variable status. Not yet implemented beyond the exemption code, ported in from orig source.
        #endregion COMMON_CONSTANTS ; #*------^ END COMMON_CONSTANTS ^------
    
        #region LOCAL_CONSTANTS ; #*------v LOCAL_CONSTANTS v------

        #  add password generator
        [Reflection.Assembly]::LoadWithPartialName("System.Web") | out-null ;

        # add retry support
        $Retries = 4 ; # number of re-attempts
        $RetrySleep = 5 ; # seconds to wait between retries
        # add CU5 validator 
        #$rgxCU5 = [moved to infra file] 

        #endregion LOCAL_CONSTANTS ; #*------^ END LOCAL_CONSTANTS ^------  
          
        #region ENCODED_CONTANTS ; #*------v ENCODED_CONTANTS v------
        # ENCODED CONsTANTS & SUPPORT FUNCTIONS:
        #region 2B4 ; #*------v 2B4 v------
        if(-not (get-command 2b4 -ea 0)){function 2b4{[CmdletBinding()][Alias('convertTo-Base64String')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ; PROCESS{$str|%{[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_))}  };} ; } ; 
        #endregion 2B4 ; #*------^ END 2B4 ^------
        #region 2B4C ; #*------v 2B4C v------
        # comma-quoted return
        if(-not (get-command 2b4c -ea 0)){function 2b4c{ [CmdletBinding()][Alias('convertto-Base64StringCommaQuoted')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ;BEGIN{$outs = @()} PROCESS{[array]$outs += $str | %{[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_))} ; } END {'"' + $(($outs) -join '","') + '"' | out-string | set-clipboard } ; } ; } ; 
        #endregion 2B4C ; #*------^ END 2B4C ^------
        #region FB4 ; #*------v FB4 v------
        # DEMO: $SitesNameList = 'THluZGFsZQ==','U3BlbGxicm9vaw==','QWRlbGFpZGU=' | fb4 ;
        if(-not (get-command fb4 -ea 0)){function fb4{[CmdletBinding()][Alias('convertFrom-Base64String')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ; PROCESS{$str | %{ [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($_)) }; } ; } ; }; 
        #endregion FB4 ; #*------^ END FB4 ^------
        # FOLLOWING CONSTANTS ARE USED FOR DEPENDANCY-LESS CONNECTIONS
        if(-not $o365_Toroco_SIDUpn){$o365_Toroco_SIDUpn = 'cy10b2RkLmthZHJpZUB0b3JvLmNvbQ==' | fb4 } ;
        $o365_SIDUpn = $o365_Toroco_SIDUpn ; 
        switch($env:Userdomain){
            'CMW'{
                if(-not $CMW_logon_SID){$CMW_logon_SID = 'Q01XXGQtdG9kZC5rYWRyaWU=' | fb4 } ; 
                $logon_SID = $CMW_logon_SID ; 
            }
            'TORO'{
                if(-not $TOR_logon_SID){$TOR_logon_SID = 'VE9ST1xrYWRyaXRzcw==' | fb4 } ; 
                $logon_SID = $TOR_logon_SID ; 
            }
            $env:COMPUTERNAME{
                $smsg = "%USERDOMAIN% -EQ %COMPUTERNAME%: $($env:computername) => non-domain-connected, likely edge role Ex server!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                if($WorkgroupName = (Get-WmiObject -Class Win32_ComputerSystem).Workgroup){
                    $smsg = "WorkgroupName:$($WorkgroupName)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                }
                if(($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup')) -or (
                        $isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup')) -or
                            $ByPassLocalExchangeServerTest){
                            $smsg = "We are on Exchange Server"
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            $IsEdgeTransport = $false
                            if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole') -or (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole')){
                                $smsg = "We are on Exchange Edge Transport Server"
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                $IsEdgeTransport = $true
                            } ; 
                } else {
                    $isLocalExchangeServer = $false 
                    $IsEdgeTransport = $false ;
                } ;
            } ; 
            default{
                $smsg = "$($env:userdomain):UNRECOGIZED/UNCONFIGURED USER DOMAIN STRING!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                THROW $SMSG 
                BREAK ; 
            }
        } ; 
        #endregion ENCODED_CONTANTS ; #*------^ END ENCODED_CONTANTS ^------

        #endregion CONSTANTS_AND_ENVIRO ; #*======^ CONSTANTS_AND_ENVIRO ^======
     
        $useSMTPCFG = $true ; 
        #region USE_SMTPCFG ; #*------v USE_SMTPCFG v------
        if($useSMTPCFG){
            # autoconfig smtp settings, keyed from cred or $env:userdomain
            $smtpToFailThru="dG9kZC5rYWRyaWVAdG9yby5jb20="| convertfrom-Base64String ;
            if(-not $rgxCertThumbprint){$rgxCertThumbprint = '[0-9a-fA-F]{40}'} ; 
            # pull the notifc smtpto from the xxxMeta.NotificationDlUs value
            if($Credential){
                if($credential.username.contains('\')){$credDom = ($Credential.username.split("\"))[0] }
                elseif($credential.username.contains('@')){$credDom = ($Credential.username.split("@"))[1] }
                elseif($credential.username -match $rgxCertThumbprint){
                    $credDom = (get-variable -name "$((get-childitem "Cert:\CurrentUser\My\$($credential.username)").subject.split('.')[0].split('-')[-1])meta").Value.o365_OPDomain ; 
                }
            }elseif($AdminAccount){
                if($AdminAccount.contains('\')){$credDom = ($AdminAccount.split("\"))[0] }
                elseif($AdminAccount.contains('@')){$credDom = ($AdminAccount.split("@"))[1] }
                elseif($AdminAccount -match $rgxCertThumbprint){
                    $credDom = (get-variable -name "$((get-childitem "Cert:\CurrentUser\My\$($AdminAccount)").subject.split('.')[0].split('-')[-1])meta").Value.o365_OPDomain ; 
                }
            }elseif($env:userdomain){$credDom = $env:userdomain}
            else { throw "Unrecognized or absent credential.username:$($AdminAccount)!. EXITING" ; } ;
            $Metas=(get-variable *meta|Where-Object{$_.name -match '^\w{3}Meta$'}) ;
            foreach ($Meta in $Metas){
                if( ($credDom -eq $Meta.value.legacyDomain) -OR ($credDom -eq $Meta.value.o365_TenantDomain) -OR ($credDom -eq $Meta.value.o365_OPDomain)){
                    if(-not $showdebug){
                        if($Meta.value.NotificationDlUs){ $smtpTo = $Meta.value.NotificationDlUs }
                        elseif($Meta.value.NotificationAddr1){$smtpTo = $Meta.value.NotificationAddr1 } 
                        else {$smtpTo=$smtpToFailThru} ;
                    } else {
                        write-verbose "debug pass: don't send to main dl, use NotificationAddr1"
                        if($Meta.value.NotificationAddr1){
                            $smtpTo = $Meta.value.NotificationAddr1 ;
                        } else {
                            $smtpTo=$smtpToFailThru ;
                        } ;
                    }
                    break ;
                } ;
            } ;
             #$smtpSubj = ("Daily Rpt: "+ (Split-Path $transcript -Leaf) + " " + [System.DateTime]::Now) ;
            $smtpSubj = "FAIL Rpt:"   ;
            #$smtpFrom = (($scriptBaseName.replace(".","-")) + "@$($Meta.value.o365_OPDomain)") ;
            if($rvEnv.isScript){
                $smtpFrom =  ($rvEnv.ScriptBaseName.replace(".","-") + "@$($Meta.value.o365_OPDomain)")  ; 
                $smtpSubj += "$($rvEnv.ScriptBaseName.replace(".","-")):$(get-date -format 'yyyyMMdd-HHmmtt')"   ;
            }elseif($rvEnv.isFunc){
                $smtpFrom =  ($rvEnv.FuncName.replace(".","-") + "@$($Meta.value.o365_OPDomain)") ; 
                $smtpSubj += "$($rvEnv.FuncName.replace(".","-")):$(get-date -format 'yyyyMMdd-HHmmtt')"   ;
            } ;        
            # if using [CmdletBinding(SupportsShouldProcess)] + -WhatIf:$($WhatIfPreference):
            if($WhatIfPreference.IsPresent -OR $whatif.IsPresent) {$smtpSubj+="WHATIF:"}        
            else {$smtpSubj+="EXEC:"} ;
            # prebuild the send-emailnotific splat w above defaults:
            $sdEmail = @{
                smtpFrom = $SMTPFrom ;
                SMTPTo = $SMTPTo ;
                SMTPSubj = $SMTPSubj ;
                #SMTPServer = $SMTPServer ;
                SmtpBody = $null ;
                SmtpAttachment = $null ;
                BodyAsHtml = $true ; # let the htmltag rgx in Send-EmailNotif flip on as needed
                verbose = $($VerbosePreference -eq "Continue") ;
            } ;
            <# send call:
            $sdEmail.SMTPSubj = "Proc Rpt:$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"   
            $sdEmail.SmtpBody = "`n===Processing Summary:" ;
            $sdEmail.SmtpBody += "`n" ;
            if($SmtpAttachment){
                $sdEmail.smtpBody +="`n(Logs Attached)" ; 
            };
            $sdEmail.SmtpBody += "Pass Completed $([System.DateTime]::Now)" ;
            $sdEmail.SmtpAttachment = $SmtpAttachment
            $smsg = "Send-EmailNotif w`n$(($sdEmail|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Send-EmailNotif @sdEmail ;
            #>
        } ; 
        #endregion USE_SMTPCFG ; #*------^ END USE_SMTPCFG ^------   

    } #  # BEG-E
    PROCESS{

        #region SUBMAIN ; #*======v SUB MAIN v======

        # 1:19 PM 2/13/2019 email trigger vari, it will be semi-delimd list of mail-triggering events
        # TenOrg or other looped-specific PassStatus (auto supported by 7pswlt)
        New-Variable -Name PassStatus_$($tenorg) -scope Script -Value $null ;
        # Clear error variable
        $Error.Clear() ;
    
        #region BANNER ; #*------v BANNER v------
        $sBnr="#*======v " ; 
        if($rvEnv.isScript){                
            if($rvEnv.PSCommandPathproxy){ $sBnr += $(split-path $rvEnv.PSCommandPathproxy -leaf) }
            elseif($script:PSCommandPath){$sBnr += $(split-path $script:PSCommandPath -leaf)}
            elseif($rPSCommandPath){$sBnr += $(split-path $rPSCommandPath -leaf)} ; 
        }elseif($rvEnv.isFunc){
            if($rvEnv.FuncDir -AND $rvEnv.FuncName){$sBnr += $rvEnv.FuncName } ; 
        } elseif($CmdletName){$sBnr += $rvEnv.CmdletName}; 
        $sBnr += ": v======" ;
        $smsg = $sBnr ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #endregion BANNER ; #*------^ END BANNER ^------
        
        <# prior specs: 
            $pltCco365Svcs=[ordered]@{
                # environment parameters:
                EnvSummary = $rvEnv ; 
                NetSummary = $netsettings ; 
                # service choices
                useEXO = $FALSE ;
                useSC = $false ; 
                UseMSOL = $false ;
                UseAAD = $false ; # M$ is actively blocking all AAD access now: Message: Access blocked to AAD Graph API for this application. https://aka.ms/AzureADGraphMigration.
                UseMG = $FALSE ;
                # Service Connection parameters
                TenOrg = $TenOrg ; # $global:o365_TenOrgDefault ; 
                Credential = $Credential ;
                AdminAccount = $AdminAccount ; 
                #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
                UserRole = $UserRole ; # @('SID','CSVC') ;
                # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
                silent = $silent ;
                MGPermissionsScope = $MGPermissionsScope ;
                MGCmdlets = $MGCmdlets ;
            } ;
            $pltCcOPSvcs=[ordered]@{
                # environment parameters:
                EnvSummary = $rvEnv ;
                NetSummary = $netsettings ;
                XoPSummary = $lclExOP ;
                # service choices
                UseExOP = $true ;
                useForestWide = $true ;
                useExopNoDep = $false ;
                ExopVers = 'Ex2010' ;
                UseOPAD = $true ;
                useExOPVers = $useExOPVers; # 'Ex2010' ;
                # Service Connection parameters
                TenOrg = $TenOrg ; # $global:o365_TenOrgDefault ;
                Credential = $Credential ;
                #[ValidateSet("SID","ESVC","LSVC")]
                #UserRole = $UserRole ; # @('SID','ESVC') ;
                # if inheriting same $userrole param/default, that was already used for cloud conn, filter out the op unsupported CBA roles
                # exclude csvc as well, go with filter on the supported ValidateSet from get-HybridOPCredentials: ESVC|LSVC|SID
                #UserRole = ($UserRole -match '(ESVC|LSVC|SID)' -notmatch 'CBA') ; # @('SID','ESVC') ;
                # coming through as match $true, not filtered
                UserRole = $UserRole |?{$_ -match '(ESVC|LSVC|SID)' -AND $_ -notmatch 'CBA'} ;  
                # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
                silent = $silent ;
            } ;
        #>

        #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======
    
        #region BROAD_SVC_CONTROL_VARIS ; #*======v BROAD_SVC_CONTROL_VARIS  v======   
        $useO365 = $true ; 
        $useOP = $true ;     
        # (config individual svcs in each block)
        #endregion BROAD_SVC_CONTROL_VARIS ; #*======^ END BROAD_SVC_CONTROL_VARIS ^======

        #region CALL_CONNECT_O365SERVICES ; #*======v CALL_CONNECT_O365SERVICES v======
        #$useO365 = $true ; 
        if($useO365){
            $pltCco365Svcs=[ordered]@{
                # environment parameters:
                EnvSummary = $rvEnv ; 
                NetSummary = $netsettings ; 
                # service choices
                useEXO = $true ;
                useSC = $false ; 
                UseMSOL = $false ;
                UseAAD = $false ; # M$ is actively blocking all AAD access now: Message: Access blocked to AAD Graph API for this application. https://aka.ms/AzureADGraphMigration.
                UseMG = $true ;
                # Service Connection parameters
                TenOrg = $TenOrg ; # $global:o365_TenOrgDefault ; 
                Credential = $Credential ;
                AdminAccount = $AdminAccount ; 
                #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
                UserRole = $UserRole ; # @('SID','CSVC') ;
                # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
                silent = $silent ;
                MGPermissionsScope = $MGPermissionsScope ;
                MGCmdlets = $MGCmdlets ;
            } ;
            write-verbose "(Purge no value keys from splat)" ; 
            $mts = $pltCco365Svcs.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltCco365Svcs.remove($_.Name)} ; rv mts -ea 0 ; 
            if((get-command connect-O365Services -EA STOP).parameters.ContainsKey('whatif')){
                $pltCco365SvcsnDSR.add('whatif',$($whatif))
            } ; 
            # add rertry on fail, up to $DoRetries
            $Exit = 0 ; # zero out $exit each new cmd try/retried
            # do loop until up to 4 retries...
            Do {
                $smsg = "connect-O365Services w`n$(($pltCco365Svcs|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $ret_ccSO365 = connect-O365Services @pltCco365Svcs ; 
                #region CONFIRM_CCEXORETURN ; #*------v CONFIRM_CCEXORETURN v------
                # matches each: $plt.useXXX:$true to matching returned $ret.hasXXX:$true 
                $vplt = $pltCco365Svcs ; $vret = 'ret_ccSO365' ; $ACtionCommand = 'connect-O365Services' ; $vtests = @() ; $vFailMsgs = @()  ; 
                $vplt.GetEnumerator() |?{$_.key -match '^use' -ANd $_.value -match $true} | foreach-object{
                    $pltkey = $_ ;
                    $smsg = "$(($pltkey | ft -HideTableHeaders name,value|out-string).trim())" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $tprop = $pltkey.name -replace '^use','has';
                    if($rProp = (gv $vret).Value.psobject.properties | ?{$_.name -match $tprop}){
                        $smsg = "$(($rprop | ft -HideTableHeaders name,value |out-string).trim())" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        if($rprop.Value -eq $pltkey.value){
                            $vtests += $true ; 
                            $smsg = "Validated: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } else {
                            $smsg = "NOT VALIDATED: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                            $vtests += $false ; 
                            $vFailMsgs += "`n$($smsg)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        };
                    } else{
                        $smsg = "Unable to locate: $($pltKey.name):$($pltKey.value) to any matching $($rprop.name)!)" ;
                        $smsg = "" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    } ; 
                } ; 
                if($vtests -notcontains $false){
                    $smsg = "==> $($ACtionCommand): confirmed specified connections *all* successful " ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    # populuate the $pltRXO.credential
                    if($ret_ccSO365.CredentialO365){
                        $pltRXO = [ordered]@{
                            Credential = $ret_ccSO365.CredentialO365 ;
                            verbose = $($VerbosePreference -eq "Continue")  ;
                        } ;
                    }else{
                        $smsg = "Unpopulated returnd:connect-O365Services.CredentialO365!" ; 
                        $smsg += "`nUNABLE TO POPULATE `$pltRXO!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    }
                    $Exit = $DoRetries ;
                } else {
                    $smsg = "==> $($ACtionCommand): FAILED SOME SPECIFIED CONNECTIONS" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    $smsg = "MISSING SOME KEY CONNECTIONS. DO YOU WANT TO IGNORE, AND CONTINUE WITH CONNECTED SERVICES?" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    $Exit ++ ;
                    $smsg = "Try #: $Exit" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    if($Exit -eq $DoRetries){
                        $smsg = "Unable to exec cmd!"; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        #-=-=-=-=-=-=-=-=
                        $sdEmail.SMTPSubj = "FAIL Rpt:$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"
                        $sdEmail.SmtpBody = "`n===Processing Summary:" ;
                        if($vFailMsgs){
                            $sdEmail.SmtpBody += "`n$(($vFailMsgs|out-string).trim())" ; 
                        } ; 
                        $sdEmail.SmtpBody += "`n" ;
                        if($SmtpAttachment){
                            $sdEmail.SmtpAttachment = $SmtpAttachment
                            $sdEmail.smtpBody +="`n(Logs Attached)" ;
                        };
                        $sdEmail.SmtpBody += "Pass Completed $([System.DateTime]::Now)" ;
                        $smsg = "Send-EmailNotif w`n$(($sdEmail|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Send-EmailNotif @sdEmail ;
                        $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ;
                        if ($bRet.ToUpper() -eq "YYY") {
                            $smsg = "(Moving on), WITH THE FOLLOW PARTIAL CONNECTION STATUS" ;
                            $smsg += "`n`n$(($ret_CcOPSvcs|out-string).trim())" ; 
                            write-host -foregroundcolor green $smsg  ;
                        } else {
                            throw $smsg ; 
                            break ; #exit 1
                        } ;  
                    } ;        
                } ; 
                #endregion CONFIRM_CCEXORETURN ; #*------^ END CONFIRM_CCEXORETURN ^------
            } Until ($Exit -eq $DoRetries) ; 
        } ; #  useO365-E
        #endregion CALL_CONNECT_O365SERVICES ; #*======^ END CALL_CONNECT_O365SERVICES ^======

        #region TEST_EXO_CONN ; #*------v TEST_EXO_CONN v------
        # ALT: simplified verify EXO conn: ALT to full CONNECT_O365SERVICES block - USE ONE OR THE OTHER!
        $useEXO = $FALSE ; 
        $useSC = $false ; 
        if(-not $XOConnectionUri ){$XOConnectionUri = 'https://outlook.office365.com'} ;
        if(-not $SCConnectionUri){$SCConnectionUri = 'https://ps.compliance.protection.outlook.com'} ;
        $EXOtestCmdlet = 'Get-xoOrganizationConfig' ; 
        if($useEXO){
            if(gcm $EXOtestCmdlet -ea 0){
                $conns = Get-ConnectionInformation -ea STOP  ; 
                $hasEXO = $hasSC = $false ; 
                #if($conns | %{$_ | ?{$_.ConnectionUri -eq 'https://outlook.office365.com' -AND $_.State -eq 'Connected' -AND $_.TokenStatus -eq 'Active'}}){
                $conns | %{
                    if($_ | ?{$_.ConnectionUri -eq $XOConnectionUri}){$hasEXO = $true } ; 
                    if($_ | ?{$_.ConnectionUri -eq $SCConnectionUri}){$hasSC = $true } ; 
                }
                if($useEXO -AND $hasEXO){
                    write-verbose "EXO ConnectionURI present" ; 
                }elseif(-not $useEXO){}else{
                    $smsg = "No Active EXO connection: Run - Connect-ExchangeOnline -Prefix xo -  before running this script!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    BREAK ; 
                } ; 
                if($useSC -AND $hasSC){
                    write-verbose "SCI ConnectionURI present" ; 
                }elseif(-not $useSC){}else{
                    $smsg = "No Active SC connection: Run - Connect-IPPSSession -Prefix SC -  before running this script!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    BREAK ; 
                } ; 
            }else {
                $smsg = "Missing gcm get-xoMailboxFolderStatistics: ExchangeOnlineManagement module *not* loaded!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                BREAK ; 
            } ;    
        }else{
            $smsg = "useEXO:$($useEXO): skipping EXO tests" ; 
            if($VerbosePreference -eq "Continue"){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } 
        #endregion TEST_EXO_CONN ; #*------^ END TEST_EXO_CONN ^------

        #region CALL_CONNECT_OPSERVICES ; #*======v CALL_CONNECT_OPSERVICES v======
        #$useOP = $false ; 
        if($useOP){
            $pltCcOPSvcs=[ordered]@{
                # environment parameters:
                EnvSummary = $rvEnv ;
                NetSummary = $netsettings ;
                XoPSummary = $lclExOP ;
                # service choices
                UseExOP = $true ;
                useForestWide = $true ;
                useExopNoDep = $false ;
                ExopVers = 'Ex2010' ;
                UseOPAD = $true ;
                useExOPVers = $useExOPVers; # 'Ex2010' ;
                # Service Connection parameters
                TenOrg = $TenOrg ; # $global:o365_TenOrgDefault ;
                Credential = $Credential ;
                #[ValidateSet("SID","ESVC","LSVC")]
                #UserRole = $UserRole ; # @('SID','ESVC') ;
                # if inheriting same $userrole param/default, that was already used for cloud conn, filter out the op unsupported CBA roles
                # exclude csvc as well, go with filter on the supported ValidateSet from get-HybridOPCredentials: ESVC|LSVC|SID
                #UserRole = ($UserRole -match '(ESVC|LSVC|SID)' -notmatch 'CBA') ; # @('SID','ESVC') ;
                # coming through as match $true, not filtered
                UserRole = $UserRole |?{$_ -match '(ESVC|LSVC|SID)' -AND $_ -notmatch 'CBA'} ; 
                # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
                silent = $silent ;
            } ;

            write-verbose "(Purge no value keys from splat)" ;
            $mts = $pltCcOPSvcs.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltCcOPSvcs.remove($_.Name)} ; rv mts -ea 0 ;
            if((get-command connect-OPServices -EA STOP).parameters.ContainsKey('whatif')){
                $pltCcOPSvcsnDSR.add('whatif',$($whatif))
            } ;
            $smsg = "connect-OPServices w`n$(($pltCcOPSvcs|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $ret_CcOPSvcs = connect-OPServices @pltCcOPSvcs ;

            #region CONFIRM_CCOPRETURN ; #*------v CONFIRM_CCOPRETURN v------
            # matches each: $plt.useXXX:$true to matching returned $ret.hasXXX:$true
            $vplt = $pltCcOPSvcs ; $vret = 'ret_CcOPSvcs' ;  ; $ACtionCommand = 'connect-OPServices' ; 
            $vplt.GetEnumerator() |?{$_.key -match '^use' -ANd $_.value -match $true} | foreach-object{
                $pltkey = $_ ;
                $smsg = "$(($pltkey | ft -HideTableHeaders name,value|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $vtests = @() ;  $vFailMsgs = @()  ; 
                $tprop = $pltkey.name -replace '^use','has';
                if($rProp = (gv $vret).Value.psobject.properties | ?{$_.name -match $tprop}){
                    $smsg = "$(($rprop | ft -HideTableHeaders name,value |out-string).trim())" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    if($rprop.Value -eq $pltkey.value){
                        $vtests += $true ; 
                        $smsg = "Validated: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } else {
                        $smsg = "NOT VALIDATED: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                        $vtests += $false ; 
                        $vFailMsgs += "`n$($smsg)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    };
                } else{
                    $smsg = "Unable to locate: $($pltKey.name):$($pltKey.value) to any matching $($rprop.name)!)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ; 
            } ; 
            if($useOP -AND $vtests -notcontains $false){
                $smsg = "==> $($ACtionCommand): confirmed specified connections *all* successful " ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                # 11:01 AM 1/27/2026 populate followon:            
                if($ret_CcOPSvcs.CredentialOP){
                    $pltRx10 = [ordered]@{
                        Credential = $ret_CcOPSvcs.CredentialOP ;
                        verbose = $($VerbosePreference -eq "Continue")  ;
                    } ;
                }else{
                    $smsg = "Unpopulated returned:connect-OPServices.CredentialOP!" ;
                    $smsg += "`nUNABLE TO POPULATE `$pltRx10!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                }
            }elseif($vtests -contains $false -AND (get-variable ret_CcOPSvcs) -AND (gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper() -ne $env:userdomain){
                $smsg = "==> $($ACtionCommand): FAILED SOME SPECIFIED CONNECTIONS" ; 
                $smsg += "`nCROSS-ORG ONPREM CONNECTION: ATTEMPTING TO CONNECT TO ONPREM '$((gv -name "$($tenorg)meta").value.o365_Prefix)' $((gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper()) domain, FROM $($env:userdomain)!" ;
                $smsg += "`nEXPECTED ERROR, SKIPPING ONPREM ACCESS STEPS (force `$useOP:$false)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $useOP = $false ; 
            }elseif(-not $useOP -AND -not (get-variable ret_CcOPSvcs)){
                $smsg = "-useOP: $($useOP), skipped connect-OPServices" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } else {
                $smsg = "==> $($ACtionCommand): FAILED SOME SPECIFIED CONNECTIONS" ; 
                $smsg += "`n`$ret_CcOPSvcs:`n$(($ret_CcOPSvcs|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $sdEmail.SMTPSubj = "FAIL Rpt:$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"
                $sdEmail.SmtpBody = "`n===Processing Summary:" ;
                if($vFailMsgs){
                    $sdEmail.SmtpBody += "`n$(($vFailMsgs|out-string).trim())" ; 
                } ; 
                $sdEmail.SmtpBody += "`n" ;
                if($SmtpAttachment){
                    $sdEmail.SmtpAttachment = $SmtpAttachment
                    $sdEmail.smtpBody +="`n(Logs Attached)" ;
                };
                $sdEmail.SmtpBody += "Pass Completed $([System.DateTime]::Now)" ;
                $smsg = "Send-EmailNotif w`n$(($sdEmail|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Send-EmailNotif @sdEmail ;
                throw $smsg ; 
                BREAK ; 
            } ; 
            #endregion CONFIRM_CCOPRETURN ; #*------^ END CONFIRM_CCOPRETURN ^------
            
            #region CONFIRM_OPFORESTWIDE ; #*------v CONFIRM_OPFORESTWIDE v------    
            if($useOP -AND $pltCcOPSvcs.useForestWide -AND $ret_CcOPSvcs.hasForestWide -AND $ret_CcOPSvcs.AdGcFwide){
                $smsg = "==> $($ACtionCommand): confirmed has BOTH .hasForestWide & .AdGcFwide ($($ret_CcOPSvcs.AdGcFwide))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success        
            }elseif($pltCcOPSvcs.useForestWide -AND (get-variable ret_CcOPSvcs) -AND (gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper() -ne $env:userdomain){
                $smsg = "`nCROSS-ORG ONPREM CONNECTION: ATTEMPTING TO CONNECT TO ONPREM '$((gv -name "$($tenorg)meta").value.o365_Prefix)' $((gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper()) domain, FROM $($env:userdomain)!" ;
                $smsg += "`nEXPECTED ERROR, SKIPPING ONPREM FORESTWIDE SPEC" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $useOP = $false ; 
            }elseif($useOP -AND $pltCcOPSvcs.useForestWide -AND -NOT $ret_CcOPSvcs.hasForestWide){
                $smsg = "==> $($ACtionCommand): MISSING CRITICAL FORESTWIDE SUPPORT COMPONENT:" ; 
                if(-not $ret_CcOPSvcs.hasForestWide){
                    $smsg += "`n----->$($ACtionCommand): MISSING .hasForestWide (Set-AdServerSettings -ViewEntireForest `$True) " ; 
                } ; 
                if(-not $ret_CcOPSvcs.AdGcFwide){
                    $smsg += "`n----->$($ACtionCommand): MISSING .AdGcFwide GC!:`n((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):326) " ; 
                } ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "MISSING SOME KEY CONNECTIONS. DO YOU WANT TO IGNORE, AND CONTINUE WITH CONNECTED SERVICES?" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ;
                if ($bRet.ToUpper() -eq "YYY") {
                    $smsg = "(Moving on), WITH THE FOLLOW PARTIAL CONNECTION STATUS" ;
                    $smsg += "`n`n$(($ret_CcOPSvcs|out-string).trim())" ; 
                    write-host -foregroundcolor green $smsg  ;
                } else {
                    throw $smsg ; 
                    break ; #exit 1
                } ;         
            }; 
            #endregion CONFIRM_OPFORESTWIDE ; #*------^ END CONFIRM_OPFORESTWIDE ^------
        } ; 
        #endregion CALL_CONNECT_OPSERVICES ; #*======^ END CALL_CONNECT_OPSERVICES ^======
    
        #endregion SERVICE_CONNECTIONS #*======^ END SERVICE_CONNECTIONS ^======
    
        <# Service Conditional if/thens: Tests above should BREAK on any fail, but these are for critical dependancy calls
        # o365 calls
        if($ret_ccSO365.hasAAD){ }else{
            $smsg = "`$ret_ccSO365.hasAAD:$($ret_ccSO365.hasAAD): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ;  ;
        if($ret_ccSO365.hasEXO){ }else{
            $smsg = "`$ret_ccSO365.hasEXO:$($ret_ccSO365.hasEXO): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        if($ret_ccO365S.hasSC){ }else{
            $smsg = "`$ret_ccO365S.hasSC:$($ret_ccO365S.hasSC): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ;         
        if($ret_ccSO365.hasMG){ }else{
            $smsg = "`$ret_ccSO365.hasMG:$($ret_ccSO365.hasMG): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        if($ret_ccSO365.hasMSOL){ }else{
            $smsg = "`$ret_ccSO365.hasMSOL:$($ret_ccSO365.hasMSOL): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        # XOP calls
        if($ret_ccOPSvcs.UseExOP){ }else{
            $smsg = "`$ret_ccOPSvcs.UseExOP:$($ret_ccOPSvcs.UseExOP): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        if($ret_ccOPSvcs.UseOPAD){ }else{
            $smsg = "`$ret_ccOPSvcs.UseOPAD:$($ret_ccOPSvcs.UseOPAD): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        #>

        # connect to ExOP X10
        <#
        if($pltRX10){
            ReConnect-Ex2010 @pltRX10 ;
        } else { Reconnect-Ex2010 ; } ;
        #>

    
        <# Service Conditional if/thens: Tests above should BREAK on any fail, but these are for critical dependancy calls
        # o365 calls
        if($ret_ccSO365.hasAAD){ }else{
            $smsg = "`$ret_ccSO365.hasAAD:$($ret_ccSO365.hasAAD): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ;  ;
        if($ret_ccSO365.hasEXO){ }else{
            $smsg = "`$ret_ccSO365.hasEXO:$($ret_ccSO365.hasEXO): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        if($ret_ccO365S.hasSC){ }else{
            $smsg = "`$ret_ccO365S.hasSC:$($ret_ccO365S.hasSC): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ;         
        if($ret_ccSO365.hasMG){ }else{
            $smsg = "`$ret_ccSO365.hasMG:$($ret_ccSO365.hasMG): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        if($ret_ccSO365.hasMSOL){ }else{
            $smsg = "`$ret_ccSO365.hasMSOL:$($ret_ccSO365.hasMSOL): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        # XOP calls
        if($ret_ccOPSvcs.UseExOP){ }else{
            $smsg = "`$ret_ccOPSvcs.UseExOP:$($ret_ccOPSvcs.UseExOP): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        if($ret_ccOPSvcs.UseOPAD){ }else{
            $smsg = "`$ret_ccOPSvcs.UseOPAD:$($ret_ccOPSvcs.UseOPAD): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        #>

        # connect to ExOP X10
        <#
        if($pltRX10){
            ReConnect-Ex2010 @pltRX10 ;
        } else { Reconnect-Ex2010 ; } ;
        #>

    
        #endregion SERVICE_CONNECTIONS #*======^ END SERVICE_CONNECTIONS ^======
    
        <# Service Conditional if/thens: Tests above should BREAK on any fail, but these are for critical dependancy calls
        # o365 calls
        if($ret_ccSO365.hasAAD){ }else{
            $smsg = "`$ret_ccSO365.hasAAD:$($ret_ccSO365.hasAAD): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ;  ;
        if($ret_ccSO365.hasEXO){ }else{
            $smsg = "`$ret_ccSO365.hasEXO:$($ret_ccSO365.hasEXO): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        if($ret_ccO365S.hasSC){ }else{
            $smsg = "`$ret_ccO365S.hasSC:$($ret_ccO365S.hasSC): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ;         
        if($ret_ccSO365.hasMG){ }else{
            $smsg = "`$ret_ccSO365.hasMG:$($ret_ccSO365.hasMG): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        if($ret_ccSO365.hasMSOL){ }else{
            $smsg = "`$ret_ccSO365.hasMSOL:$($ret_ccSO365.hasMSOL): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        # XOP calls
        if($ret_ccOPSvcs.UseExOP){ }else{
            $smsg = "`$ret_ccOPSvcs.UseExOP:$($ret_ccOPSvcs.UseExOP): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        if($ret_ccOPSvcs.UseOPAD){ }else{
            $smsg = "`$ret_ccOPSvcs.UseOPAD:$($ret_ccOPSvcs.UseOPAD): MISSING DEPENDANT CONNECTION!" ; 
            $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
    #>

        #$logging = $True ; # need to set in scope outside of functions
        $pltInput=[ordered]@{} ;

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
        if($verbose){$pltInput.add("verbose",$(($VerbosePreference -eq "Continue"))) } ;
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
            #$pltInput.Cu5=$Cu5;
            # looks like it's adding on assign (?.?)7
            if($Cu5){$pltInput.add("Cu5",$Cu5) } ;
        } else {
            $pltInput.add("Cu5",$null) ;
        }  ;
    
        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):new-MailboxShared w`n$(($pltInput|out-string).trim())" ;
        if(($psISE -ANd (get-date  -format 'M/d/yyyy') -eq '1/19/2026')){
            if((gcm new-mailboxshared).source -eq 'verb-Ex2010'){
                ipmo -fo -verb 'D:\scripts\new-MailboxShared_func.ps1'
            }
        } ;
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
