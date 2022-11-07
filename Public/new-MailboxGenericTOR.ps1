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

    ticket "355925";  DisplayName "XXX Confirms" ;  MInitial "" ;  Owner "LOGON";  BaseUser "AccountsReceivable";  IsContractor $false;  NonGeneric $true
    -ticket "355925" -DisplayName "XXX Confirms"  -MInitial ""  -Owner "LOGON" -BaseUser "AccountsReceivable" -NonGeneric

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
    $whatIf=$true ;
    $insplat=@{  ticket="TICKET" ;
      DisplayName="DNAME"  ;
      MInitial="" ;
      Owner="OWNER" ;
      showDebug=$true ;
      PermsDays=999 ;
      members="GRANTEE1,GRANTEE2";
      } ;
    if(!$dc){$dc=get-gcfast} ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):using common `$dc:$($dc)" ;
    $pltNmbx=[ordered]@{  ticket=$insplat.ticket ;
      DisplayName=$insplat.DisplayName  ;
      MInitial="" ;
      Owner=$insplat.Owner ;
      NonGeneric=$false  ;
      Vscan="YES" ;
      domaincontroller=$dc ;
      showDebug=$true ;
     whatIf=$($whatif) ;
    } ;
    if($insplat.Cu5){$pltNmbx.add("CU5",$insplat.CU5)} ;
    write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):`$insplat:`n$(($insplat|out-string).trim())" ;
    write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):new-MailboxGenericTOR.ps1 w`n$(($pltNmbx|out-string).trim())" ;
    new-MailboxGenericTOR.ps1 @pltNmbx;
    if(!$whatif){    Do {write-host "." -NoNewLine;
        Start-Sleep -m (1000 * 5)} Until (($tmbx = get-mailbox $insplat.DisplayName -domaincontroller $dc )) ;
        if($tmbx){               $pltGrant=[ordered]@{  ticket=$insplat.ticket  ;
      TargetID=$tmbx.samaccountname ;
      Owner=$insplat.Owner ;
      PermsDays=$insplat.PermsDays ;
      members=$insplat.members ;
      domaincontroller=$dc ;
      showDebug=$true ;
      whatIf=$whatif ;
    } ;
        write-host -foregroundcolor green "n$((get-date).ToString('HH:mm:ss')):add-MbxAccessGrant.ps1 w`n$(($pltGrant|out-string).trim())" ;
        add-MbxAccessGrant.ps1 @pltGrant ;
          cmsol ;
     write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):PREPARING DAWDLE LOOP!($($tmbx.PrimarySmtpAddress))`nMsolLastSync:`n$((get-MsolLastSync| ft -a TimeGMT,TimeLocal|out-string).trim())" ;
          Do {rxo ;
     write-host "." -NoNewLine;
    Start-Sleep -s 30} Until ((get-exorecipient $tmbx.PrimarySmtpAddress -EA 0)) ;
     write-host "`n*READY TO MOVE*!`a" ;
     sleep -s 1 ;
     write-host "*READY TO MOVE*!`a" ;
     sleep -s 1 ;
      write-host "*READY TO MOVE*!`a`n" ;
          write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):Running:`nmove-EXOmailboxNow.ps1 -TargetMailboxes $($tmbx.ALIAS) -showDebug -whatIf`n`n" ;
          . move-EXOmailboxNow.ps1 -TargetMailboxes $tmbx.ALIAS -showDebug -whatIf ;
          $strMoveCmd="move-EXOmailboxNow.ps1 -TargetMailboxes $($tmbx.ALIAS) -showDebug -NoTEST -whatIf:`$false" ;
          write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Move Command (copied to cb):`n`n$($strMoveCmd)`n" ;
          $strMoveCmd | out-clipboard ;
          $strCleanCmd="get-exomoverequest -BatchName ExoMoves-* | ?{`$_.status -eq 'Completed'} | Remove-exoMoveRequest -whatif" ;
          write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):Post-completion Cleanup Command :`n`n$($strCleanCmd)`n" ;
        } else { write-warning "No mbx found matching $($insplat.DisplayName). ABORTING"} ;
    } else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(WHATIF skipping AMPG & move)" } ;
    BP Scriptblock that pre-parses basic $insplat inputs and feeds new-MailboxGenericTOR.ps1, add-MbxAccessGrant.ps1 & move-EXOmailboxNow.ps1 (unwrap for use, normally stored in psb-PSnewMbxG.cbp)
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

    $verbose = ($VerbosePreference -eq "Continue") ;

    # Get the name of this function
    ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
    # Get parameters this function was invoked with
    $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;

    
    if ($Whatif){$bWhatif=$true ; write-host -foregroundcolor green "`$Whatif is $true (`$bWhatif:$bWhatif)" ; };
    # ISE also has not perfect but roughly equiv workingdir (unless script cd's):
    #$ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath('.\')
    
    write-verbose "`$PSScriptRoot:$($PSScriptRoot)`n`$PSCommandPath:$($PSCommandPath)" ;
    # most of this falls apart once you move a script to a module - all resolve to the module .ps(m|d)1 file, have to use ${CmdletName} for funcs, instead
    if ($psISE -AND (!($PSScriptRoot) -AND !($PSCommandPath))){
        $ScriptDir = Split-Path -Path $psISE.CurrentFile.FullPath ;
        $ScriptBaseName = split-path -leaf $psise.currentfile.fullpath ;
        $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($psise.currentfile.fullpath) ;
        $PSScriptRoot = $ScriptDir ;
        if($PSScriptRoot -ne $ScriptDir){ write-warning "UNABLE TO UPDATE BLANK `$PSScriptRoot TO CURRENT `$ScriptDir!"} ;
        $PSCommandPath = $psise.currentfile.fullpath ;
        if($PSCommandPath -ne $psise.currentfile.fullpath){ write-warning "UNABLE TO UPDATE BLANK `$PSCommandPath TO CURRENT `$psise.currentfile.fullpath!"} ;
    } else {
        if($host.version.major -lt 3){
            $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
            $PSCommandPath = $myInvocation.ScriptName ;
            $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
        } elseif($PSScriptRoot) {
            $ScriptDir = $PSScriptRoot ;
            if($PSCommandPath){
                $ScriptBaseName = split-path -leaf $PSCommandPath ;
                $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($PSCommandPath) ;
            } else {
                $PSCommandPath = $myInvocation.ScriptName ;
                $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ;
                $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
            } ;
        } else {
            if($MyInvocation.MyCommand.Path) {
                $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
                $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ;
                $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
            } else {
                throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$MyInvocation IS BLANK!" ;
            } ;
        } ;
    } ;
    write-verbose "`$ScriptDir:$($ScriptDir)`n`$ScriptBaseName:$($ScriptBaseName)`n`$ScriptNameNoExt:$($ScriptNameNoExt)`n`$PSScriptRoot:$($PSScriptRoot)`n`$PSCommandPath:$($PSCommandPath)" ;

    #  add password generator
    [Reflection.Assembly]::LoadWithPartialName("System.Web") | out-null ;

    # add retry support
    $Retries = 4 ; # number of re-attempts
    $RetrySleep = 5 ; # seconds to wait between retries
    # add CU5 validator 
    #$rgxCU5 = [moved to infra file] 

    # Clear error variable
    $Error.Clear() ;
    #endregion INIT; # ------

    #region FUNCTIONS ; # ------
    #*======v FUNCTIONS v======

    #-------v Function _cleanup v-------
    #make it defer to existing script-copy
    <#if(test-path function:Cleanup){
        "(deferring to `$script:cleanup())" ;
    } else {
        "(using default verb-transcript:cleanup())" ;
    #>
        function _cleanup {
            # clear all objects and exit
            # 8:55 AM 5/11/2021 ren: internal helper func Cleanup() -> _cleanup() 
            # 1:36 PM 11/16/2018 Cleanup:stop-transcriptlog left tscript running, test again and re-stop
            # 8:15 AM 10/2/2018 Cleanup:make it defer to $script:cleanup() (needs to be preloaded before verb-transcript call in script), added missing semis, replaced all $bDebug -> $showDebug
            # 2:02 PM 9/21/2018 missing $timestampnow, hardcode
            # 8:45 AM 10/13/2015 reset $DebugPreference to default SilentlyContinue, if on
            # # 8:46 AM 3/11/2015 at some time from then to 1:06 PM 3/26/2015 added ISE Transcript
            # 8:39 AM 12/10/2014 shifted to stop-transcriptLog function
            # 7:43 AM 1/24/2014 always stop the running transcript before exiting
            write-verbose "$((get-date).ToString('HH:mm:ss')):_cleanup" ; 
            #stop-transcript ;
            <#actually, with write-log in use, I don't even need cleanup /ISE logging, it's already covered in new-mailboxshared() etc)
            if($host.Name -eq "Windows PowerShell ISE Host"){
                # shift the logfilename gen out here, so that we can arch it
                #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -uformat "%Y%m%d-%H%M" ) + "-ISEtrans.log")) ;
                # shift to static timestamp $timeStampNow
                #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + $timeStampNow + "-ISEtrans.log")) ;
                # missing $timestampnow, hardcode
                #$Logname=(join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -format 'yyyyMMdd-HHmmtt') + "-ISEtrans.log")) ;
                #-=-=-=-=-=-=-=-=
                #$ParentPath = $MyInvocation.MyCommand.Definition ;
                $ParentPath = $PSCommandPath ; 
                if($ParentPath){
                    $rgxProfilePaths='(\\Documents\\WindowsPowerShell\\scripts|\\Program\sFiles\\windowspowershell\\scripts)' ;
                    if($ParentPath -match $rgxProfilePaths){
                        $ParentPath = "$(join-path -path 'c:\scripts\' -ChildPath (split-path $ParentPath -leaf))" ;
                    } ;
                    $logspec = start-Log -Path ($ParentPath) -showdebug:$($showdebug) -whatif:$($whatif) ;
                    if($logspec){
                        $logging=$logspec.logging ;
                        $logfile=$logspec.logfile ;
                        $Logname=$logspec.transcript ;
                    } else {$smsg = "Unable to configure logging!" ; write-warning "$((get-date).ToString('HH:mm:ss')):$($sMsg)" ; Exit ;} ;
                } else {$smsg = "No functional `$ParentPath found!" ; write-warning "$((get-date).ToString('HH:mm:ss')):$($sMsg)" ;  Exit ;} ;
                #-=-=-=-=-=-=-=-=
                $logname = $logname.replace(".log","-ISEtrans.log")
                write-host "`$Logname: $Logname";
                Start-iseTranscript -logname $Logname ;
                #Archive-Log $Logname ;
                # standardize processing file so that we can send a link to open the transcript for review
                $transcript = $Logname ;
            } else {
                write-verbose"$(Get-Date -Format 'HH:mm:ss'):Stop Transcript" ;
                Stop-TranscriptLog ;
                #write-verbose "$(Get-Date -Format 'HH:mm:ss'):Archive Transcript" ;
                #Archive-Log $transcript ;
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):`$transcript:$(($transcript|out-string).trim())" ;
            } # if-E
            # Cleanup:stop-transcriptlog left tscript running, test again and re-stop
            if (Test-Transcribing) {
                Stop-Transcript
                if ($showdebug) {write-host -foregroundcolor green "`$transcript:$transcript"} ;
            }  # if-E
            # add an exit comment
            #>
            write-host -foregroundcolor green "END $BARSD4 $scriptBaseName $BARSD4"  ;
            write-host -foregroundcolor green "$BARSD40" ;
            # finally restore the DebugPref if set
            if ($ShowDebug -OR ($DebugPreference = "Continue")) {
                write-host -foregroundcolor green "Resetting `$DebugPreference from 'Continue' back to default 'SilentlyContinue'" ;
                $showdebug=$false ;
                # also need to enable write-debug output (and turn this off at end of script, it's a global, normally SilentlyContinue)
                $DebugPreference = "SilentlyContinue" ;
            } # if-E
            exit ;
        #} ;
    } ; #*------^ END Function _cleanup ^------

    # moved new-MailboxShared => verb-Ex2010

    #*======^ END Functions ^======
    #endregion FUNCTIONS ; # ------

    #region SUBMAIN ; # ------
    #*======v SUB MAIN v======


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
    if($whatIf){$pltInput.add("whatIf",$whatIf) } ;

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
}

#*------^ new-MailboxGenericTOR.ps1 ^------
