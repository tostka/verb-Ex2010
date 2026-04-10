# add-MbxAccessGrant.ps1

# dbg: TestScriptMbx2247115 dbg: cls ; .\add-MbxAccessGrant-mod.ps1 -ticket 99999  -TargetID TestScriptMbx2247115 -Owner LOGON -PermsDays 999 -members "LOGON" -NoPrompt -showDebug -domaincontroller SERVER -whatIf ;

function add-MbxAccessGrant {
    <#
    .SYNOPSIS
    add-MbxAccessGrant.ps1 - Wrapper/pre-processor function to Add Mbx Access to a specified mailbox (leverages generic add-MailboxAccessGrant())
    .NOTES
    Version     : 1.0.1
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2020-
    FileName    : 
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell,Exchange,Permissions,Mailbox,Exchange2010
    AddedCredit : 
    AddedWebsite: 
    AddedTwitter: 
    REVISIONS
    * 2:05 PM 4/27/2020 debugged, fully ported to published/installed use
    * 3:57 PM 4/9/2020 genericized for pub, moved material into infra, updated hybrid mod loads, cleaned up comments/remmed material ; updated to use start-log, debugged to funciton on jumpbox, w divided modules
    * 2:30 PM 10/1/2019 fixed 2405 errant duped catch block post merge
    * 1:54 PM 10/1/2019 manually merged branch master with the long open/outstanding branch tostka/update-new-MailboxGenericTOR.ps1
    # 9:54 AM 9/27/2019 added `a beep to all "YYY" prompts, to draw attn while multitasking
    # 8:34 AM 6/27/2019 fixed add-adgroupmember autoresolution param break - had typo -member which worked until they added a new conflicting -membertimetolive param, should have been -members (tho' add-dgmember uses -member)
    # 2:32 PM 6/13/2019 converted to function: add-MbxAccessGrant(), passed
    debugging, saving out to add-MbxAccessGrant-function.ps1 ->
    add-MbxAccessGrant.ps1 and backing up the original
    # 11:05 AM 6/13/2019 updated get-admininitials(), # 12:50 PM 6/13/2019 repl get-timestamp() -> Get-Date -Format 'HH:mm:ss' throughout
    # 2:19 PM 4/29/2019 add TOL to the domain param validateset on get-gcfast copy (sync'd in from verb-ex2010.ps1 vers)
    # 11:43 AM 2/15/2019 debugged update through a prod revision
    # 11:15 AM 2/15/2019 copied in bug-fixed write-log() with fixed debug support
    # 10:41 AM 2/15/2019 updated write-log to latest deferring version
    # 10:39 AM 2/15/2019 added full write-log logging support
    # 3:24 PM 2/6/2019 #1416:needs -prop * to pull msExchRecipientDisplayType,showInAddressBook,mail etc
    # 8:36 AM 9/6/2018 switched secgrp to Global->Universal scope, mail-enabled as DG, and hiddenfromaddressbook, debugged out issues, used in prod creation
    # 10:28 AM 6/27/2018 add $domaincontroller param option - skips dc discovery process and uses the spec, also updated $findOU code to work with TOL dom
    # 11:05 AM 3/29/2018 #1116: added trycatch, UST lacked the secgrp ou and was failing ou lookup
    # 10:31 AM 11/22/2017 shifted a block of "User mbx grant:" confirmation into review block, also tightened up the formatted whitespace to make the material pasted into cw reflect all that you need to know on the grant status. also added distrib code
    # 1:17 PM 11/15/2017 949: no, this needs to be the obj (was extracting samaccountname)
    # 12:35 PM 11/15/2017 debugged EXO-hosted Owner code to function. worked granting GRANTEE2 (exo) access to shared 'SharedTestEXOOwner' OP
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
    # 12:11 PM 2/27/2017 fixed compat in SITE (prolly SITE too) - resolved any Owner entered as email, to the samaccountname ; #1081: drop the pipe!
     $oSG | Get-ADGroupMember -server $Domain | select distinguishedName ; #1263 threw up trying to do the get-aduser on the members, # 1283 replace user lookup with this (skip the getxxx member, pull members right out of properties)
    # 1:04 PM 2/24/2017 tweak below
    #12:56 PM 2/24/2017 doesn't run worth a damn SITE-> SITE/SITE, force it to abort (avoid half-built remote objects that take too long to replicate back to SITE)
    # 12:24 PM 2/24/2017 fixed updated membership report bug - pulled pipe, probably dehydrated object issue sin remote ps
    # 12:11 PM 2/24/2017 fix vscode/code.exe char set damage: It replaced dashes (-) with "?"
    # fix -join typo/damage
    # 12:44 PM 10/18/2016 update rgx for ticket to accommodate 5-digit (or 6) CW numbers "^\d{6}$"=>^\d{5,6}$
    # 9:11 AM 9/30/2016 added pretest if(get-command -name set-AdServerSettings -ea 0)
    # # 12:22 PM 6/21/2016 secgrp membership seldom comes through clean, add a refresh loop
    # 10:52 AM 6/20/2016 fixed typo $InputSplatSiteOverride => $InputSplat.SiteOverride (broke -SiteOverride function)
    # 11:02 AM 6/7/2016 updated get-aduser review cmds to use the same dc, not the -domain global.ad.toro.com etc
    # 1:34 PM 5/26/2016 confirmed/verified works fine with SITE-hosted mbx under 376336 
    # 11:45 AM 5/19/2016 corrected $tmbx ref's to use $tmbx.identity v. $tmbx.samaccountname, now working. Retry code in place for SITE, but it didn't trigger during testing
    # 2:37 PM 5/18/2016 implmented Secgrp OU and Secgrp stnd name
    # 2:28 PM 5/18/2016 support dmg's latest unilateral changes:With the recent AD changes, all email access groups should be named     XXX-SEC-Email-firstname lastname-G and stored in XXX\Managed Groups\SEC Groups\Email Access. The generics were also renamed to XXX\Generic Email Accounts.

    # 2:17 PM 5/10/2016 used successfully to set a SITE manager perm's on an SERVERMail02-hosted user. didn't time out, Set-MailboxPermission command completed after ~3 secs
    #     fixed bad param example, remmed out non-functional Owner in the SGSplat (nosuch param), and re-enabled the ManagedBy on the SG - it's not a mbx,
    #     so why not set ManagedBy, doesn't get used by the org chart in SP
    # 2:38 PM 3/17/2016 stop populating anything into any managed-by; it's an OrgChart political value now. Rename ManagedBy param and object names in here to 'Owner'
    # 1:12 PM 2/11/2016 fixed new bug in get-GCFast, wasn't detecting blank $site
    # 12:20 PM 2/11/2016 updated to standard EMS/AD Call block & Add-EMSRemote()
    # 9:36 AM 2/11/2016 just shifting to a single copy, with no # Requires at all, losing the -psv2.ps1 version
    # 2:23 PM 2/10/2016 debugged mismatched {}, working from SITE now
    # 1:54 PM 2/10/2016 recoded to work on SITE and SITE, this version just needs the #Requires -Version 3 for psv2 enabled to be a psv3 version
    #         added fundemental upgrade to the AD Site detection, to work from SITE/SITENAME and SITE SITENAME
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


    #-=-=DISTRIB SCRIPT (ALL EX SERVERS)-=-=-=-=-=-=
    [array]$files = (gci -path "\\$env:COMPUTERNAME\c$\usr\work\exch\scripts\add-MbxAccessGrant.ps1" | ?{$_.Name -match "(?i:(.*\.(PS1|CMD)))$" })  ;[array]$srvrs = get-exchangeserver | ?{(($_.IsMailboxServer) -OR ($_.IsHubTransportServer))} | select  @{Name='COMPUTER';Expression={$_.Name }} ;$srvrs = $srvrs|?{$_.computer -ne $($env:COMPUTERNAME) } ; $srvrs | % { write-host "$($_.computer)" ; copy $files -Destination \\$($_.computer)\c$\scripts\ -whatif ; } ; get-date ;
    #-=-=-=-=-=-=-=-=


    .DESCRIPTION 
    add-MbxAccessGrant.ps1 - Add Mbx Access to a specified mailbox

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
    .PARAMETER TenOrg 
	TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']
	.PARAMETER Credential
	Credential to use for this connection [-credential [credential obj variable]")][System.Management.Automation.PSCredential]
	.PARAMETER UserRole
	Role of account (SID|CSID|UID|B2BI|CSVC|ESvc|LSvc)[-UserRole SID]
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
    .PARAMETER NoOutput
    Switch to enable output (success/fail), defaults false, but adding to support tested function execution.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    .\add-MbxAccessGrant.ps1 -ticket 123456 -SiteOverride SITE -TargetID lynctest13 -Owner LOGON -PermsDays 999 -members "GRANTEE1,GRANTEE2" -showDebug -whatIf ;
    Parameter Whatif test with Debug messages displayed
    .LINK
    #>


    <# SecGrp Name spec: =
    2:29 PM 5/18/2016updated per dawn:
            XXX-SEC-Email-firstname lastname-G
    # orig spec:
    ($sSite + "-Data-Email-" + $Tmbx.DisplayName + "-G") ;
    Create the grp :Scope: Global, Type:Secureity
    Add members.
    Add mbx permission to the grp:
    add-mailboxpermission "bossvisiplex" -User "IRO-Data-Email-Boss Visiplex" -AccessRights FullAccess -whatif ;
    #>

    Param(
        [Parameter(HelpMessage="Target Mailbox for Access Grant[name,emailaddr,alias]")]
        [string]$TargetID,
        [Parameter(HelpMessage="Custom override default generated name for Perm-hosting Security Group[[SIT]-SEC-Email-[DisplayName]-G]")]
        [string]$SecGrpName,
        [Parameter(HelpMessage="Specify the userid to be responsible for access-grant-approvals[name,emailaddr,alias]")]
        [string]$Owner,
        [Parameter(HelpMessage="Specify the number of day's the access-grant should be in place. (60 default. 999=permanent)[30-60,999]")]
        [ValidateRange(7,999)]
        [int]$PermsDays,
        [Parameter(HelpMessage="Specify a 3-letter Site Code. Used to force DL name/placement to vary from TargetID's current site[3-letter Site code]")]
        [string]$SiteOverride,
        [Parameter(HelpMessage="Comma-delimited string of potential users to be granted access[name,emailaddr,alias]")]
        [string]$Members,
        [Parameter(HelpMessage="Incident number for the change request[[int]nnnnnn]")]
        [int]$Ticket,
        [Parameter(HelpMessage="Suppress YYY confirmation prompts [-NoPrompt]")]
        [switch] $NoPrompt,
        [Parameter(HelpMessage="Option to hardcode a specific DC [-domaincontroller xxxx]")]
        [string]$domaincontroller,
        $TenOrg = 'TOR',
	    [Parameter(HelpMessage="Credential to use for cloud actions [-credential [credential obj variable]")][System.Management.Automation.PSCredential]
	    $Credential,
	    [ValidateSet('SID','CSID','UID','B2BI','CSVC')]
	    [string]$UserRole='SID',
        [Parameter(HelpMessage='Debugging Flag [$switch]')]
        [switch] $showDebug,
        [Parameter(HelpMessage='Whatif Flag [$switch]')]
    [switch] $whatIf) ;

    # NoPrompt Suppress YYY confirmation prompts

    # Get the name of this function
    ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
    # Get parameters this function was invoked with
    $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters 

    # 2:50 PM 5/18/2016 add SITE retry code
    $Retries = 4 ; # number of re-attempts
    $RetrySleep = 5 ; # seconds to wait between retries

    # 12:26 PM 5/31/2017 maintain-offboards.ps1 regex constants
    #$rgxBannedOUs=[infra file]
    #$rgxUserOUs==[infra file]


    # 12:49 PM 5/10/2016 updated the BP INIT block to stnd
    #region INIT; # ------
    #*======v SCRIPT/DOMAIN/MACHINE/INITIALIZATION-DECLARE-BOILERPLATE v======
    # * vers: 8:31 AM 5/3/2016: updated tightened up, reflects most common matl want in every script
    # * 2:10 PM 2/4/2015 shifted to here to accommodate include locations
    # pick up the bDebug from the $ShowDebug switch parameter
    # SCRIPT-CONFIG MATERIAL TO SET THE UNDERLYING $DBGPREF:
    if ($Whatif){$bWhatif=$true ; Write-Verbose -Verbose:$true "`$Whatif is $true (`$bWhatif:$bWhatif)" ; };
    if($bdebug){$ErrorActionPreference = 'Stop' ; write-debug "(Setting `$ErrorActionPreference:$ErrorActionPreference;"};

    if($showDebug){
        write-host -foregroundcolor green "`SHOWDEBUG: `$PSScriptRoot:$($PSScriptRoot)`n`$PSCommandPath:$($PSCommandPath)" ;
    } ;
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
    # ISE also has not perfect but roughly equiv workingdir (unless script cd's):
    #$ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath('.\')
    write-verbose "`$ScriptDir:$($ScriptDir)`n`$ScriptBaseName:$($ScriptBaseName)`n`$ScriptNameNoExt:$($ScriptNameNoExt)`n`$PSScriptRoot:$($PSScriptRoot)`n`$PSCommandPath:$($PSCommandPath)" ;
    

    # 11:19 AM 2/24/2017 add password generator
    [Reflection.Assembly]::LoadWithPartialName("System.Web") | out-null ;

    # Clear error variable
    $Error.Clear() ;

    #endregion INIT; # ------

    #region FUNCTIONS ; # ------
    #*======v FUNCTIONS v======


    function _cleanup  {
        # clear all objects and exit
        # 11:15 AM 5/11/2021 renamed helper func Cleanup -> _cleanup
        # 1:36 PM 11/16/2018 Cleanup:stop-transcriptlog left tscript running, test again and re-stop
        # 8:15 AM 10/2/2018 Cleanup:make it defer to $script:cleanup() (needs to be preloaded before verb-transcript call in script), added missing semis, replaced all $bDebug -> $showDebug
        # 2:02 PM 9/21/2018 missing $timestampnow, hardcode
        # 8:45 AM 10/13/2015 reset $DebugPreference to default SilentlyContinue, if on
        # # 8:46 AM 3/11/2015 at some time from then to 1:06 PM 3/26/2015 added ISE Transcript
        # 8:39 AM 12/10/2014 shifted to stop-transcriptLog function
        # 7:43 AM 1/24/2014 always stop the running transcript before exiting
        if ($showdebug) {"_cleanup  "} ;
        #stop-transcript ;
        <#actually, with write-log in use, I don't even need cleanup /ISE logging, it's already covered in new-mailboxshared() etc)
        if($host.Name -eq "Windows PowerShell ISE Host"){
            # 8:46 AM 3/11/2015 shift the logfilename gen out here, so that we can arch it
            #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -uformat "%Y%m%d-%H%M" ) + "-ISEtrans.log")) ;
            # 2:16 PM 4/27/2015 shift to static timestamp $timeStampNow
            #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + $timeStampNow + "-ISEtrans.log")) ;
            # 2:02 PM 9/21/2018 missing $timestampnow, hardcode
            $Logname=(join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -format 'yyyyMMdd-HHmmtt') + "-ISEtrans.log")) ;
            write-host "`$Logname: $Logname";
            Start-iseTranscript -logname $Logname ;
            #Archive-Log $Logname ;
            # 1:23 PM 4/23/2015 standardize processing file so that we can send a link to open the transcript for review
            $transcript = $Logname ;
        } else {
            if($showdebug){ write-debug "$(Get-Date -Format 'HH:mm:ss'):Stop Transcript" };
            Stop-TranscriptLog ;
            #if($showdebug){ write-debug "$(Get-Date -Format 'HH:mm:ss'):Archive Transcript" };
            #Archive-Log $transcript ;
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):`$transcript:$(($transcript|out-string).trim())" ;
        } # if-E
        # 1:36 PM 11/16/2018 _cleanup  :stop-transcriptlog left tscript running, test again and re-stop
        if (Test-Transcribing) {
            Stop-Transcript
            if ($showdebug) {write-host -foregroundcolor green "`$transcript:$transcript"} ;
        }  # if-E
        #>
        #11:10 AM 4/2/2015 add an exit comment
        Write-Verbose -Verbose:$verbose "END $BARSD4 $scriptBaseName $BARSD4"  ;
        Write-Verbose -Verbose:$verbose "$BARSD40" ;
        # finally restore the DebugPref if set
        if ($ShowDebug -OR ($DebugPreference = "Continue")) {
            Write-Verbose -Verbose:$true "Resetting `$DebugPreference from 'Continue' back to default 'SilentlyContinue'" ;
            $showdebug=$false ;
            # 8:41 AM 10/13/2015 also need to enable write-debug output (and turn this off at end of script, it's a global, normally SilentlyContinue)
            $DebugPreference = "SilentlyContinue" ;
        } # if-E
        exit ;
    #} ;
    } ; #*------^ END Function _cleanup   ^------




    #*======^ END Functions ^======
    #endregion FUNCTIONS ; # ------

    #region SUBMAIN ; # ------
    #*======v SUB MAIN v======


    $pltInput=[ordered]@{} ;

    if ($PSCommandPath) { $pltInput.add("ParentPath", $PSCommandPath) } ;
    if($TargetID){$pltInput.add("TargetID",$TargetID) } ;
    if($SecGrpName){$pltInput.add("SecGrpName",$SecGrpName) } ;
    if($Owner){$pltInput.add("Owner",$Owner) } ;
    if($PermsDays){$pltInput.add("PermsDays",$PermsDays) } ;
    if($SiteOverride){$pltInput.add("SiteOverride",$SiteOverride) } ;
    if($Members){$pltInput.add("Members",$Members) } ;
    if($Ticket){$pltInput.add("Ticket",$Ticket) } ;
    if($NoPrompt){$pltInput.add("NoPrompt",$NoPrompt) } ;
    if($domaincontroller){$pltInput.add("domaincontroller",$domaincontroller) } ;
    if($showDebug){$pltInput.add("showDebug",$showDebug) } ;
    if($whatIf){$pltInput.add("whatIf",$whatIf) } ;

    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):add-MbxAccessGrant w`n$(($pltInput|out-string).trim())" ;
    if(-not($NoOutput)){
        $bRet = add-MailboxAccessGrant @pltInput ;  
        $bRet | write-output ;
    } else { 
        add-MailboxAccessGrant @pltInput
    } ; 
    _cleanup   ;
    #Exit ;

    #*======^ END SUB MAIN ^======
    #endregion SUBMAIN ; # ------
} #*------^ END Function add-MbxAccessGrant ^------