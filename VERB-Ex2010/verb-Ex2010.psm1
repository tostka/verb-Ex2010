﻿# verb-ex2010.psm1


<#
.SYNOPSIS
VERB-Ex2010 - Exchange 2010 PS Module-related generic functions
.NOTES
Version     : 6.2.3
Author      : Todd Kadrie
Website     :	https://www.toddomation.com
Twitter     :	@tostka
CreatedDate : 1/16.2.30
FileName    : VERB-Ex2010.psm1
License     : MIT
Copyright   : (c) 1/16.2.30 Todd Kadrie
Github      : https://github.com/tostka
REVISIONS
* 11:22 AM 3/13/2020 Get-ExchangeServerInSite added a ping-test, to only return matches that are pingable, added -NoPing param, to permit (faster) untested bypass
* 6:25 PM 1/21/2020 - 1.0.0.1, rebuild, see if I can get a functional module out
* 1/16.2.30 - 1.0.0.0
# 7:31 PM 1/15/2020 major revise - subbed out all identifying constants, rplcd regex hardcodes with builds sourced in tor-incl-infrastrings.ps1. Tests functional.
# 11:34 AM 12/30/2019 ran vsc alias-expansion
# 7:51 AM 12/5/2019 Connect-Ex2010:retooled $ExAdmin variant webpool support - now has detect in the server-pick logic, and on failure, it retries to the stock pool.
# 10:19 AM 11/1/2019 trimmed some whitespace
# 10:05 AM 10/31/2019 added sample load/call info
# 12:02 PM 5/6.2.39 added cx10,rx10,dx10 aliases
# 11:29 AM 5/6.2.39 load-EMSLatest: spliced in from tsksid-incl-ServerApp.ps1, purging ; alias Add-EMSRemote-> Connect-Ex2010 ; toggle-ForestView():moved from tsksid-incl-ServerApp.ps1
# * 1:02 PM 11/7/2018 updated Disconnect-PssBroken
# 4:15 PM 3/24/2018 updated pshhelp
# 1:24 PM 11/2/2017 fixed connect-Ex2010 example code to include $Ex2010SnapinName vari for the snapin name (regex no worky for that)
# 1:33 PM 11/1/2017 add load-EMSSnapin (for use on server desktops)
# 11:37 AM 11/1/2017 shifted get-GcFast into here
# 9:29 AM 11/1/2017 spliced in Get-ExchangeServerInSite with updated auto-switch for ADL|SPB|LYN runs
# 8:02 AM 11/1/2017 updated connect-ex2010 & disconnect-ex2010 (add/remove-PSTitlebar), added disconnect-PssBroken
# 1:28 PM 12/9/2016: Reconnect-Ex2010, put in some logic to suppress errors
# 1:05 PM 12/9/2016 updated the docs & comments on new connect-Ex2010;Reconnect-Ex2010;Disconnect-Ex2010;Get-ExchangeServerInSite;Disconnect-PssBroken funcs and bp CALL code into function
# 11:03 AM 12/9/2016 debugged the new connect-Ex2010;Reconnect-Ex2010;Disconnect-Ex2010;Get-ExchangeServerInSite;Disconnect-PssBroken funcs and bp CALL code into function
.DESCRIPTION
VERB-Ex2010 - Exchange 2010 PS Module-related generic functions
.INPUTS
None
.OUTPUTS
None
.EXAMPLE
.EXAMPLE
.LINK
https://github.com/tostka/verb-Ex2010

#>


    $script:ModuleRoot = $PSScriptRoot ;
    $script:ModuleVersion = (Import-PowerShellDataFile -Path (get-childitem $script:moduleroot\*.psd1).fullname).moduleversion ;
    $runningInVsCode = $env:TERM_PROGRAM -eq 'vscode' ;

#*======v FUNCTIONS v======




#*------v add-MailboxAccessGrant.ps1 v------
function add-MailboxAccessGrant {
    <#
    .SYNOPSIS
    add-MailboxAccessGrant.ps1 - Add Mbx Access to a specified mailbox
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
    Tags        : Powershell,Exchange,Permissions,Exchange2010
    REVISIONS
    # 5:12 PM 10/13/2021 fixed long standing random add-adgroupmember bug (failed to see target sg/dg), by swapping in ADGM ex cmd;  pulled [int] from $ticket , to permit non-numeric & multi-tix
    # 11:36 AM 9/16.2.31 string
    # 10:27 AM 9/14/2021 beefed up echos w 7pswhsplat's pre
    # 3:21 PM 8/17/2021 recoded grabbing outputs, on object creations in EMS, tho' AD object creations generate no proper output. functoinal on current ticket.
    # 4:48 PM 8/16.2.31 still wrestling grant fails, switched the *permission -user targets to the dg object.primarysmtpaddr (was adg.samaccountname), if the adg.sama didn't match the alias, that woulda caused failures. Seemd to work better in debugging
    # 1:51 PM 6/30/201:51 PM 6/30/2021 trying to work around sporadic $oSG add-mailboxperm fails, played with -user $osg designator - couldn't use DN, went back to samacctname, but added explicit RETRY echos on failretries (was visible evid at least one retry was in mix) ; hardened up the report gathers - stuck thge get-s in a try/catch ahead of echos, (vs inlines) ; we'll see if the above improve the issue - another option is to build something that can parse the splt echo back into a functional splat, to at least make remediation easier (copy, convert, rerun on fly).
    # 4:21 PM 5/19/2021 added -ea STOP to splats, to force retry's to trigger
    # 5:03 PM 5/18/2021 fixed the fundementally borked start-log I created below
    # 11:10 AM 5/11/2021 swapped parentpath code for dyn module-support code (moving the new-mailboxgenerictor & add-mbxaccessgrant preproc .ps1's to ex2010 mod functions)
    # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods
    # 1:27 PM 4/23/2020 updated loadmod & dynamic logging/exec code
    # 4:28 PM 4/22/2020 updated logging code, to accomodate dynamic locations and $ParentPath
    # 3:37 PM 4/9/2020 works fully on jumpbox, but ignores whatif, renamed $bwhatif -> $whatif (as the b variant was prev set in the same-script, now separate scopes); swapped out CU5 switch, moved settings into infra file, genericized
    # 1:38 PM 4/9/2020 modularized updated to reflect debugging on jumpbox
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
    # 2:28 PM 5/18/2016 With the recent AD changes, all email access groups should be named         XXX-SEC-Email-firstname lastname-G     and stored in XXX\Managed Groups\SEC Groups\Email Access.     The generics were also renamed to XXX\Generic Email Accounts.
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
    .PARAMETER ParentPath
    Calling script path (used for log construction)[-ParentPath c:\pathto\script.ps1]
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
    .\add-MailboxAccessGrant -ticket 123456 -SiteOverride LYN -TargetID lynctest13 -Owner SOMERECIP -PermsDays 999 -members "lynctest16,lynctest18" -showDebug -whatIf ;
    Parameter Whatif test with Debug messages displayed
    .EXAMPLE
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
    add-MbxAccessGrant @pltInput
    Splatted version
    .LINK
    https://github.com/tostka/verb-Ex2010/
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
        # [int] # 10:30 AM 10/13/2021 pulled, to permit non-numeric & multi-tix
        $Ticket,
        [Parameter(HelpMessage = "Suppress YYY confirmation prompts [-NoPrompt]")]
        [switch] $NoPrompt,
        [Parameter(HelpMessage = "Option to hardcode a specific DC [-domaincontroller xxxx]")]
        [string]$domaincontroller,
        [Parameter(HelpMessage = "Calling script path (used for log construction)[-ParentPath c:\pathto\script.ps1]")]
        [string]$ParentPath,
        [Parameter(HelpMessage = 'Debugging Flag [$switch]')]
        [switch] $showDebug,
        [Parameter(HelpMessage = 'Whatif Flag [$switch]')]
        [switch] $whatIf
    ) ;

    # NoPrompt Suppress YYY confirmation prompts

    BEGIN {

        # don't use the LoadModFile(), it has scoping issues returning the mods, they aren't accessible outside the function itself

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
            if($lVers){                 $lVers=($lVers | Sort-Object version)[-1];                 try {                     import-module -name $tModName -RequiredVersion $lVers.Version.tostring() -force -DisableNameChecking -verbose:$($false)                 }   catch {                      write-warning "*BROKEN INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;import-module -name $tModDFile -force -DisableNameChecking -verbose:$($false)                } ;
            } elseif (test-path $tModFile) {                 write-warning "*NO* INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;                 try {import-module -name $tModDFile -force -DisableNameChecking -verbose:$($false)}                 catch {                     write-error "*FAILED* TO LOAD MODULE*:$($tModName) VIA $(tModFile) !" ;                     $tModFile = "$($tModName).ps1" ;                     $sLoad = (join-path -path $LocalInclDir -childpath $tModFile) ;                     if (Test-Path $sLoad) {                         Write-Verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                         . $sLoad ;                         if ($showdebug) { Write-Verbose "Post $sLoad" };                     } else {                         $sLoad = (join-path -path $backInclDir -childpath $tModFile) ;                         if (Test-Path $sLoad) {                             write-verbose  ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                             . $sLoad ;                             if ($showdebug) { write-verbose  "Post $sLoad" };                         } else {                             Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;                             exit;                         } ;                     } ;                 } ;             } ;
            if(!(test-path function:$tModCmdlet)){                 write-warning "UNABLE TO VALIDATE PRESENCE OF $tModCmdlet`nfailing through to `$backInclDir .ps1 version" ;                 $sLoad = (join-path -path $backInclDir -childpath "$($tModName).ps1") ;                 if (Test-Path $sLoad) {                     write-verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                     . $sLoad ;                     if ($showdebug) { Write-Verbose "Post $sLoad" };                     if(!(test-path function:$tModCmdlet)){                         write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO CONFIRM `$tModCmdlet:$($tModCmdlet) FOR $($tModName)" ;                     } else {                         write-verbose  "(confirmed $tModName loaded: $tModCmdlet present)"                     }                 } else {                     Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;                     exit;                 } ;
            } else {                 write-verbose  "(confirmed $tModName loaded: $tModCmdlet present)"             } ;
        } ;  # loop-E
        #*------^ END MOD LOADS ^------

        <#
        if($ParentPath){
            $rgxProfilePaths='(\\Documents\\WindowsPowerShell\\scripts|\\Program\sFiles\\windowspowershell\\scripts)' ;
            if($ParentPath -match $rgxProfilePaths){
                $ParentPath = "$(join-path -path 'c:\scripts\' -ChildPath (split-path $ParentPath -leaf))" ;
            } ;
            $logspec = start-Log -Path ($ParentPath) -showdebug:$($showdebug) -whatif:$($whatif) -tag $TargetID;
            if($logspec){
                $logging=$logspec.logging ;
                $logfile=$logspec.logfile ;
                $transcript=$logspec.transcript ;
            } else {$smsg = "Unable to configure logging!" ; write-warning "$((get-date).ToString('HH:mm:ss')):$($sMsg)" ; Exit ;} ;
        } else {$smsg = "No functional `$ParentPath found!" ; write-warning "$((get-date).ToString('HH:mm:ss')):$($sMsg)" ;  Exit ;} ;
        #>
        # with shift of add-mbxaccessgrant preprocessor to mod func, the above needs to be recoded, as $ParentPath would wind up a module file
        # detect profile installs (installed mod or script), and redir to stock location
        $dPref = 'd','c' ; foreach($budrv in $dpref){ if(test-path -path "$($budrv):\scripts" -ea 0 ){ break ;  } ;  } ;
        [regex]$rgxScriptsModsAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)" ;
        [regex]$rgxScriptsModsCurrUserScope="^$([regex]::escape([environment]::getfolderpath('Mydocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)" ;
        # -Tag "($TenOrg)-LASTPASS" 
        $pltSLog = [ordered]@{ NoTimeStamp=$false ; Tag=$lTag  ; showdebug=$($showdebug) ;whatif=$($whatif) ;} ;
        if($PSCommandPath){
            if(($PSCommandPath -match $rgxScriptsModsAllUsersScope) -OR ($PSCommandPath -match $rgxScriptsModsCurrUserScope) ){
                # AllUsers or CU installed script, divert into [$budrv]:\scripts (don't write logs into allusers context folder)
                if($PSCommandPath -match '\.ps(d|m)1$'){
                    # module function: use the ${CmdletName} for childpath
                    $pltSLog.Path= (join-path -Path "$($budrv):\scripts" -ChildPath "$(${CmdletName}).ps1" )  ;
                } else { 
                    $pltSLog.Path=(join-path -Path "$($budrv):\scripts" -ChildPath (split-path $PSCommandPath -leaf)) ;
                } ; 
            }else {
                $pltSLog.Path=$PSCommandPath ;
            } ;
        } else {
            if( ($MyInvocation.MyCommand.Definition -match $rgxScriptsModsAllUsersScope) -OR ($MyInvocation.MyCommand.Definition -match $rgxScriptsModsCurrUserScope) ){
                $pltSLog.Path=(join-path -Path "$($budrv):\scripts" -ChildPath (split-path $PSCommandPath -leaf)) ;
            } else {
                $pltSLog.Path=$MyInvocation.MyCommand.Definition ;
            } ;
        } ;
        $smsg = "start-Log w`n$(($pltSLog|out-string).trim())" ;
        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        $logspec = start-Log @pltSLog ;
        
        if($logspec){
            $logging=$logspec.logging ;
            $logfile=$logspec.logfile ;
            $transcript=$logspec.transcript ;
            
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
            
            if(Test-TranscriptionSupported){
                $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                start-transcript -Path $transcript ;
            } ;
        } else {throw "Unable to configure logging!" } ;
        

        <#
        $sBnr="#*======v START PASS:$($ScriptBaseName) v======" ;
        $smsg= "$($sBnr)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
        #>

        $xxx = "====VERB====";
        $xxx = $xxx.replace("VERB", "NewMbxAccess") ;
        $BARS = ("=" * 10);
        #write-host -fore green ((get-date).ToString('HH:mm:ss') + ":===PASS STARTED=== ")


        $reqMods += "Add-PSTitleBar;Remove-PSTitleBar".split(";") ;
        #Disconnect-EMSR (variant name in some ps1's for Disconnect-Ex2010)
        #$reqMods+="Reconnect-CCMS;Connect-CCMS;Disconnect-CCMS".split(";") ;
        #$reqMods+="Reconnect-SOL;Connect-SOL;Disconnect-SOL".split(";") ;
        $reqMods += "Test-TranscriptionSupported;Test-Transcribing;Stop-TranscriptLog;Start-IseTranscript;Start-TranscriptLog;get-ArchivePath;Archive-Log;Start-TranscriptLog".split(";") ;
        # 12:15 PM 9/12/2018 remove dupes
        $reqMods = $reqMods | Select-Object -Unique ;

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
            # switch to EXO-compatible group type: Univ, mail-enable
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
            # pulling id, pipeline it in
            $ADMbxGrantSplat = @{
                User           = "" ;
                ExtendedRights = "Send As" ;
            };
        } else {
            # psv3 code
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
                ErrorAction     = 'STOP' # need this to trigger retries
            };
            $SGUpdtSplat = [ordered]@{
                Identity = "";
                Server   = ""
                ErrorAction     = 'STOP' # need this to trigger retries
            };
            $DGEnableSplat = [ordered]@{
                Identity         = "";
                DomainController = ""
                ErrorAction     = 'STOP' # need this to trigger retries
            };
            $DGUpdtSplat = [ordered]@{
                Identity                      = "";
                HiddenFromAddressListsEnabled = $true ;
                DomainController              = "" ;
                ErrorAction     = 'STOP' # need this to trigger retries
            } ;
            $GrantSplat = [ordered]@{
                Identity        = "" ;
                User            = "" ;
                AccessRights    = "FullAccess";
                InheritanceType = "All";
                ErrorAction     = 'STOP' # need this to trigger retries
            };
            # add for AD SendAs perms grant
            #pulling id, pipeline it in
            $ADMbxGrantSplat = [ordered]@{
                User           = "" ;
                ExtendedRights = "Send As" ;
                ErrorAction     = 'STOP' # need this to trigger retries
            };
        }

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
        if(get-pssession |Where-Object{($_.configurationname -eq 'Microsoft.Exchange') -AND ($_.ComputerName -match $rgxEx10HostName) -AND ($_.IdleTimeout -ne -1)} ){
            write-verbose  "$((get-date).ToString('HH:mm:ss')):LOCAL EMS detected" ;
            $Global:E10IsDehydrated=$false ;
        # REMS detect dleTimeout -eq -1
        } elseif(get-pssession |Where-Object{$_.configurationname -eq 'Microsoft.Exchange' -AND $_.ComputerName -match $rgxEx10HostName -AND ($_.IdleTimeout -eq -1)} ){
            write-verbose  "$((get-date).ToString('HH:mm:ss')):REMOTE EMS detected" ;
            $reqMods+="get-GCFast;Get-ExchangeServerInSite;connect-Ex2010;Reconnect-Ex2010;Disconnect-Ex2010;Disconnect-PssBroken".split(";") ;
            if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
            reconnect-ex2010 ;
            $Global:E10IsDehydrated=$true ;
        } else {
            write-verbose  "$((get-date).ToString('HH:mm:ss')):No existing Ex2010 Connection detected" ;
            # Server snapin defer
            if(($host.version.major -lt 3) -AND (get-service MSExchangeADTopology -ea SilentlyContinue)){
                write-verbose "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Loading Local Server EMS10 Snapin" ;
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
        write-verbose  "$((get-date).ToString('HH:mm:ss')):(loading ADMS...)" ;
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

                if ( ($InputSplat.OwnerMbx = (get-mailbox -identity $($InputSplat.Owner) -ea stop)) ) {
                    if ($showdebug) { $smsg = "UserMailbox detected" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; } ;
                } else {
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
        if ($InputSplat.SiteOverride) {
            $SiteCode = $($InputSplat.SiteOverride);
        } else {
            # we need to use the OwnerMbx - Owner currently is the alias, we want the object with it's dn
            $SiteCode = $InputSplat.OwnerMbx.identity.tostring().split("/")[1]  ;
        } ;
        if ($env:USERDOMAIN -eq $TORMeta['legacyDomain']) {
            $FindOU = "^OU=Email\sAccess,OU=SEC\sGroups,OU=Managed\sGroups,";
        } ELSEif ($env:USERDOMAIN -eq $TOLMeta['legacyDomain']) {
            # CN=Lab-SEC-Email-Thomas Jefferson,OU=Email Access,OU=SEC Groups,OU=Managed Groups,OU=LYN,DC=SUBDOM,DC=DOMAIN,DC=DOMAIN,DC=com
            $FindOU = "^OU=Email\sAccess,OU=SEC\sGroups,OU=Managed\sGroups,"; ;
        } else {
            throw "UNRECOGNIZED USERDOMAIN:$($env:USERDOMAIN)" ;
        } ;

        $SGSplat.DisplayName = "$($SiteCode)-SEC-Email-$($Tmbx.DisplayName)-G";

        TRY {
            $OU = (Get-ADObject -filter { ObjectClass -eq 'organizationalunit' } -server $($DomainController) | Where-Object { $_.distinguishedname -match "^$($FindOU).*OU=$($SiteCode),.*,DC=ad,DC=toro((lab)*),DC=com$" } | Select-Object distinguishedname).distinguishedname.tostring() ;
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
        # 3:50 PM 10/13/2021 flip ea stop to continue, we want it to get through, even if it throws error, and continue will complain
        $SGMembers = ($InputSplat.members.split(",") | ForEach-Object { get-recipient $_ -ea continue | select -expand primarysmtpaddress  | select -unique})
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

            # 4:16 PM 8/16.2.31 nope, it's not dg-enabled yet, it's a secgrp, can't pull it.
            # try flipping the $osg adg, to a resolved equiv dg obj, and use it's primarysmtpaddress rather than the adg.samaccountname (which may not match the alias, as an id). The dG should be a native rcp obj better fit to the add-mailboxperm cmd
            $oDG = get-distributiongroup -DomainController $InputSplat.DomainController -Identity $osg.DistinguishedName -ErrorAction 'STOP' ;
            #$SGUpdtSplat.Identity = $DGEnableSplat.Identity = $DGUpdtSplat.Identity = $GrantSplat.User = $ADMbxGrantSplat.User = $oSG.samaccountname ;
            $GrantSplat.User = $ADMbxGrantSplat.User = $oDG.primarysmtpaddress ;

            # _append_ the $InfoStr into any existing Info for the object
            # can't use [ordered] on psv2 if we must have these in order use a psv2 OrderedDictionary
            if (($host.version.major) -lt 3) {
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

            if ($oSG.info) {
                # existing info tag
                # update the splat
                # just loop each line split on `n: (Get-ADUser lynctest9 -Properties info).info.split("`n")| foreach{"Ln:$_"}
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
            # build the Notes/Info field as a hashcode: OtherAttributes=@{    info="TargetMbx:SOMERECIP`r`nPermsExpire:6/19/2015"  } ;
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
            # unlike most other modules, ADMS and it's new-ADGroup does *not* return the created object. Have to qry the object back, cold. 
            New-AdGroup @SGSplat -whatif ;
            $DGEnableSplat.identity = $SGSplat.SamAccountName ;
            $DGUpdtSplat.identity = $SGSplat.SamAccountName ;

            $smsg = "`$DGEnableSplat:`n---`n$(($DGEnableSplat|out-string).trim())`n---`n" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

            write-host -foregroundcolor yellow "$((get-date).ToString("HH:mm:ss")):Continue with $($SGSplat.Name) creation?...";
            if ($NoPrompt) { $bRet = "YYY" } else { $bRet = Read-Host "Enter YYY to continue. Anything else will exit`a" ; } ;
            if ($bRet.ToUpper() -eq "YYY") {


                if ($whatif) {
                    $smsg = "-Whatif pass, skipping exec." ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                } else {
                    $smsg = "Executing...";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                    New-AdGroup @SGSplat  ;
                    Do { write-host "." -NoNewLine; Start-Sleep -s 1 } Until (Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -server $($InputSplat.DomainController)) ;
                    #$oSG= (get-adgroup "$($SGSplat.DisplayName)" -server $($InputSplat.Domain) -ea stop );
                    $oSG = Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -server $($InputSplat.DomainController) -ErrorAction stop;
                    if ($bDebug) {
                        $smsg = "`$oSG:$($oSG.SamAccountname)`n`$oSG.DN:$($oSG.DistinguishedName)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; #Error|Warn
                    } ;
                    $smsg = "Enable-DistributionGroup w`n$(($DGEnableSplat|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    # capture the enabl - EMS ps returns the intact $osg DG object
                    $oDG = Enable-DistributionGroup @DGEnableSplat ;
                    $smsg = "Set HiddenFromAddressListsEnabled:Set-DistributionGroup w`n$(($DGUpdtSplat|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    # but set-dg does *not* return the updated object, it has to be re-queried for current status. 
                    Set-DistributionGroup @DGUpdtSplat ;
                    $oSG = Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -prop * -server $($InputSplat.DomainController) -ErrorAction stop;
                    $smsg = "Final SecGrp Config:$($oSG.SamAccountname)`n:$(($oSG | fl Name,GroupCategory,GroupScope,msExchRecipientDisplayType,showInAddressBook,mail|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    # qry back updated status
                    $oDG = get-distributiongroup -DomainController $InputSplat.DomainController -Identity $oDG.DistinguishedName -ErrorAction 'STOP' ;
                    #$SGUpdtSplat.Identity = $DGEnableSplat.Identity = $DGUpdtSplat.Identity = $GrantSplat.User = $ADMbxGrantSplat.User = $oSG.samaccountname ;
                    $GrantSplat.User = $ADMbxGrantSplat.User = $oDG.primarysmtpaddress ;
                } ;
            } else { $smsg = "INVALID KEY ABORTING NO CHANGE!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; Exit ; } ;
        } # if-E $osg

        $smsg = "`nTesting SecGrp Members Add `nto group: $($oSG.Name)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
        if ($oSG -AND $oDG) {
            
            $DGEnableSplat.Identity = $DGUpdtSplat.Identity = $SGUpdtSplat.Identity = $oSG.samaccountname ;
            
            # we're using the samaccountname for -user spec, 
            # can't use the DN either (won't resolve)
            $SGUpdtSplat.Server = $($InputSplat.DomainController) ;
            $DGEnableSplat.DomainController = $($InputSplat.DomainController) ;
            $DGUpdtSplat.DomainController = $($InputSplat.DomainController) ;
            # 12:47 PM 10/6/2015 add dc
            $GrantSplat.Add("DomainController", $($InputSplat.domaincontroller)) ;
            #8:41 AM 10/14/2015 add adp
            $ADMbxGrantSplat.Add("DomainController", $($InputSplat.domaincontroller)) ;

            $ExistMbrs = get-distributiongroupmember -Identity $oSG.samaccountname -DomainController $domaincontroller -ErrorAction 'Stop' | select -expand primarysmtpaddress ; 
            $pltAddDGM=[ordered]@{
                identity=$oDG.alias ;
                #Member= $mbr  ; 
                BypassSecurityGroupManagerCheck=$true 
                ErrorAction = 'Stop' ; 
                whatif=$($whatif) ; 
                DomainController= $domaincontroller
            } ;
            <# with AddDGM, if you're not the explicit owner, you get:
            You don't have sufficient permissions. This operation can only be performed by a manager of the group.
            use the -BypassSecurityGroupManagerCheck param to quash the check
            #>
            if ($whatif) {
                $smsg = "-Whatif pass, skipping exec." ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
            } else {
                foreach ($Mbr in $SGMembers) { 
                    if ($ExistMbrs -notcontains $Mbr) {
                        $smsg = "ADD:$($mbr)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                        #Add-ADGroupMember @SGUpdtSplat -members $($mbr.samaccountname)  -whatif ;
                        <# AAGM keeps throwing
                            Couldn't resolve the user or group "GROUP DNAME." If the user or group is a foreign forest principal, you must have either a two-way trust or an outgoing trust.
                            + CategoryInfo          : InvalidOperation: (:) [], LocalizedException
                            + FullyQualifiedErrorId : 9A7F344F
                            + PSComputerName        : DC.DOMAIN.COM
                        #> 
                        # flip the adds to adgm
                        Add-DistributionGroupMember @pltAddDGM -member $mbr ; 
                    } else {
                        $smsg = "SKIPPING:$($mbr.samaccountname) is already a member of $($oSG.samaccountname)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    } ;
                }  # loop-E ;
                <# toss out the whole prompted thing, just do it above
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
                                Add-ADGroupMember @SGUpdtSplat -members $($mbr.samaccountname)  ;
                            } ;
                        } else {
                            "SKIPPING:$($mbr.samaccountname) is already a member of $($oSG.samaccountname)"
                        } ;
                    } #  # loop-E;
                } ;
                #>
            } # if-E whatif ;
            #$mbxp = $Tmbx | get-mailboxpermission -user ($oSG).Name -domaincontroller $InputSplat.domaincontroller -ea silentlycontinue | 
            $mbxp = $Tmbx | get-mailboxpermission -user $oSG.samaccountname -domaincontroller $InputSplat.domaincontroller  | 
                Where-Object { $_.user -match ".*-(SEC|Data)-Email-.*$" }
            $smsg = "`nChecking Mailbox Permission on $($Tmbx.samaccountname) mailbox to accessing user:`n $($oSG.Name)...`n(blank if none)`n---`n$(($mbxp | select user,AccessRights,IsInhertied,Deny | format-list|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug


            # AD SendAs too

            $mbxadp = $Tmbx | Get-ADPermission -domaincontroller $($InputSplat.domaincontroller) -ea Silentlycontinue |
                 Where-Object { ($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and ($_.user -match ".*-(SEC|Data)-Email-.*$") };

            $smsg = "`nChecking AD SendAs Permission on $($Tmbx.samaccountname) mailbox to accessing user:`n $($oSG.Name)...`n(blank if none)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

            $smsg = "`n$(($mbxadp | select identity,User,ExtendedRights,Deny,Inherited | format-list|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
            # format-table -wrap ;

            $smsg = "`n---`nExisting $($oSG.Name) Membership...`n(blank if none)`n$((Get-ADGroupMember -identity $oSG.samaccountname -server $($DomainController) | select distinguishedName|out-string).trim())`n" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

            $smsg = "Testing Permissions Grant Update...`nAdd-MailboxPermission -whatif w`n$(($GrantSplat|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

            # add retry :
            $Exit = 0 ;
            # do loop until up to 4 retries...
            Do {
                if($Exit -gt 0){
                    $smsg = "RETRY#:$($exit)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
                Try {
                     # capture returned added perms (not full perms on mbx)
                    $addedmbxp = add-mailboxpermission @GrantSplat -whatif ;
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
                    Continue ;
                } # try-E
            } Until ($Exit -eq $Retries) # loop-E

            #$smsg = "Add-ADPermission -whatif... w`n$(($ADMbxGrantSplat|out-string).trim())" ;
            #if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
            $smsg = "add-adpermission w`n-identity $($TMbx.Identity)`n$(($ADMbxGrantSplat|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            $Exit = 0 ;
            Do {
                if($Exit -gt 0){
                    $smsg = "RETRY#:$($exit)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
                Try {
                    # capture returned added perms (not full perms on mbx)
                    $addedadmbxp = add-adpermission -identity $($TMbx.Identity) @ADMbxGrantSplat -whatif ;
                    $Exit = $Retries ;
                } Catch {
                    $ErrTrapd = $Error[0] ;
                    Start-Sleep -Seconds $RetrySleep ;
                    $Exit ++ ;
                    $smsg = "Failed to exec add-mailboxpermission -whatif cmd because: $($ErrTrpd)`nTry #: $($Exit)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ;
                    If ($Exit -eq $Retries) { $smsg = "Unable to exec cmd!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } ; } ;
                    Continue ;
                } # try-E
            } Until ($Exit -eq $Retries) # loop-E

            write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):Exec Permissions Grant Update";
            if ($whatif) {
                # 11:17 AM 6/22/2015 whatif-only pass
                write-verbose "SKIPPING EXEC: Whatif-only pass";
            } else {
                $smsg = "add-mailboxpermission w`n$(($GrantSplat|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $Exit = 0 ;
                # do loop until up to 4 retries...
                Do {
                    if($Exit -gt 0){
                        $smsg = "RETRY#:$($exit)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 
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
                        Continue ;
                    } # try-E
                } Until ($Exit -eq $Retries) # loop-E

                $smsg = "Add-ADPermission -whatif:identity $($TMbx.Identity) w`n$(($ADMbxGrantSplat|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                $Exit = 0 ;
                Do {
                    if($Exit -gt 0){
                        $smsg = "RETRY#:$($exit)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 
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
                        Continue ;
                    } # try-E
                } Until ($Exit -eq $Retries) # loop-E

                # generics don't need this, test the OU path and only add folks below users
                # we're only hiding folks matching:
                #$rgxBannedOUs=[xxx]
                # and unhiding folks matching
                if ($Tmbx.distinguishedname -match $rgxUserOUs) {
                    # block that adds the $tmbx to the maintain-offboards.ps1 target AccGrant group for the region
                    $smsg = "Add TMBX $($tMbx.samaccountname) to AccGrant Group`n$(($TMbx | select -expand distinguishedname|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                    $grpN = "LYN-DL-Exch-AGUnHide" ;
                    $smsg = "==TGroup:$($grpN)";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                    if ($tdl = get-DistributionGroup -identity $grpN -domaincontroller $($InputSplat.domaincontroller) ) {
                        
                        $pltAddDGM=@{
                            identity=$tdl.alias ;Member=$TMbx.distinguishedname; domaincontroller=$($InputSplat.domaincontroller) ;whatif=$($whatif);ErrorAction='STOP';
                        } ; 

                        $smsg = "==Add $($TMbx.name) to $($tdl.alias):" ;
                        $smsg += "`nadd-DistributionGroupMember w`n$(($pltAddDGM|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;                        
                        $Exit = 0 ;
                        # do loop until up to 4 retries...
                        Do {
                            Try {
                                add-DistributionGroupMember @pltAddDGM ;
                                #-identity $tdl.alias -Member $TMbx.distinguishedname -domaincontroller $($InputSplat.domaincontroller) -whatif:$($whatif) ;

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
            write-verbose "$(Get-Date -Format 'HH:mm:ss'):Waiting 5secs to refresh";
            Start-Sleep -s 5 ;

            # secgrp membership seldom comes through clean, add a refresh loop
            do {
                # 12:53 PM 6/30/2021 idsolate the get's & add try/catch
                TRY{
                    $propsMbxP = 'user','AccessRights','IsInhertied','Deny' ; 
                    $propsAMbxP = 'User','ExtendedRights','Inherited','Deny' ; 
                    $rMbxP = get-mailboxpermission -identity $($TMbx.Identity) -user $oSG.samaccountname -domaincontroller $($InputSplat.domaincontroller) |
                        ?{$_.user -match ".*-(SEC|Data)-Email-.*$"} ; 
                    $rAMbxP = Get-ADPermission -identity $($TMbx.Identity) -domaincontroller $($InputSplat.domaincontroller) -user $oSG.distinguishedName ; 
                    $mbrs = Get-ADGroupMember -identity $oSG.distinguishedName -server $($DomainController) | 
                        Select-Object distinguishedName ;
                 <# orig, revised to modern standard below
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
                #>
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    
                    Continue ;#Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                } ; 
                
                $smsg = "===REVIEW SETTINGS:===`n----Updated Permissions:`n`nChecking Mailbox/AD Permission on $($Tmbx.samaccountname) mailbox `n to accessing user:`n $($oSG.SamAccountName)`n---" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                #$smsg = "`n$((get-mailboxpermission -identity $($TMbx.Identity) -user $(($oSG).Name) -domaincontroller $($InputSplat.domaincontroller) | ?{$_.user -match ".*-(SEC|Data)-Email-.*$"} | format-list user,AccessRights,IsInhertied,Deny |out-string).trim())" ;
                # 12:52 PM 6/30/2021 fix typo:
                #$smsg = "`n$((get-mailboxpermission -identity $($TMbx.Identity) -user $oSG.distinguishedName -domaincontroller $($InputSplat.domaincontroller) | ?{$_.user -match ".*-(SEC|Data)-Email-.*$"} | format-list user,AccessRights,IsInhertied,Deny |out-string).trim())" ;
                $smsg = "`n$(($rMbxP | format-list $propsMbxP |out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                #$smsg = "`n$((get-mailboxpermission -identity $($TMbx.Identity) -user $(($oSG).Name) -domaincontroller $($InputSplat.domaincontroller) | ?{$_.user -match ".*-(SEC|Data)-Email-.*$"} | format-list user,AccessRights,IsInhertied,Deny|out-string).trim())" ;
                #if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                $smsg = "`n==User mbx grant: Confirming $($TMbx.name) member of $($grpN):" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                # 10:04 AM 11/22/2017 put the accgrant confirmation into the output:
                if ($Tmbx.distinguishedname -match $rgxUserOUs) {
                    #$smsg = "$((Get-ADPermission -identity $($TMbx.Identity) -domaincontroller $($InputSplat.domaincontroller) -user "$($oSG.SamAccountName)"|  format-list User,ExtendedRights,Inherited,Deny | out-string).trim())" ;
                    $smsg = "`n$(($rAMbxP|out-string | format-list $propsAMbxP ).trim())" ; 
                    # $rAMbxP 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                } else {
                    $smsg = "TMBX $($tMbx.samaccountname) is in a non-User OU: Term Hide/Unhide groups do not apply...";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                }  ;

                $smsg = "`nUpdated $($oSG.Name) Membership...`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):---";
                #if ($mbrs = Get-ADGroupMember -identity $oSG.samaccountname -server $($DomainController) | Select-Object distinguishedName ) {
                if ($mbrs) {
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


#*------v add-MbxAccessGrant.ps1 v------
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
    # 10:30 AM 10/13/2021 pulled [int] from $ticket , to permit non-numeric & multi-tix
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
        # [int] # 10:30 AM 10/13/2021 pulled, to permit non-numeric & multi-tix
        $Ticket,
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
    #_cleanup   ;
    #Exit ;

    #*======^ END SUB MAIN ^======
    #endregion SUBMAIN ; # ------
}

#*------^ add-MbxAccessGrant.ps1 ^------


#*------v Connect-Ex2010.ps1 v------
Function Connect-Ex2010 {
  <#
    .SYNOPSIS
    Connect-Ex2010 - Setup Remote ExchOnPrem Mgmt Shell connection (validated functional Exch2010 - Exch2016)
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    AddedCredit : Inspired by concept code by ExactMike Perficient, Global Knowl... (Partner)
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Version     : 1.1.0
    CreatedDate : 2020-02-24
    FileName    : Connect-Ex2010()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell
    REVISIONS   :
    * 1:30 PM 9/5/2024 added  update-SecurityProtocolTDO() SB to begin
    * 3:11 PM 7/15/2024 needed to change CHKPREREQ to check for presence of prop, not that it had a value (which fails as $false); hadn't cleared $MetaProps = ...,'DOESNTEXIST' ; confirmed cxo working non-based
    * 10:47 AM 7/11/2024 cleared debugging NoSuch etc meta tests
    * 1:34 PM 6/21/2024 ren $Global:E10Sess -> $Global:EXOPSess ; add: prereq checks, and $isBased support, to devert into most connect-exchangeServerTDO, get-ADExchangeServerTDO 100% generic fall back support (including buffering in the pair of funcs)
    # 9:43 AM 7/27/2021 revised -PSTitleBar to support suffix EMS[ctl]
    # 1:31 PM 7/21/2021 revised Add-PSTitleBar $sTitleBarTag with TenOrg spec (for prompt designators)
    # 3:18 PM 5/18/2021 somehow lost $credOpTORSID, so flipped lost default $credOPTor -> $credTORSID
    * 11:40 AM 5/14/2021 added -ea 0 to the gv tests (suppresses not-found error when called without logging config)
    * 11:22 AM 4/21/2021 coded around recent 'verbose the heck out of everything', yanked 99% of the verbose support - this seldom fails in a way that you need verbose, and when it's on, every cmdlet in the modules get echo'd, spams the heck out of console & logging. One key change (not sure if source) was to switch from inline import-pss & import-mod, into 2 steps with varis.
    * 10:02 AM 4/12/2021 add alias connect-ExOP (eventually rename verb-ex2010 to verb-exOnPrem)
    * 12:06 PM 4/2/2021 added alias cxOP ; added explicit echo on import-session|module, removed redundant catch block; added trycatch around import-sess|mod ; added recStatus support
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods ; renamed-standardized splat names (EMSSplat ->pltNSess ; ) ; flipped prefix into splat add ;
    * 2:36 PM 3/23/2021 getting away from dyn, random from array in $XXXMeta.Ex10Server, doesn't rely on AD lookups for referrals
    * 10:14 AM 3/23/2021 flipped default $Cred spec, pointed at an OP cred (matching reconnect-ex2010())
    * 11:36 AM 3/5/2021 updated colorcode, subed wv -verbose with just write-verbose, added cred.uname echo
    * 1:15 PM 3/1/2021 added org-level color-coded console
    * 3:28 PM 2/17/2021 updated to support cross-org, leverages new $XXXMeta.ExRevision, ExViewForest
    * 5:16 PM 10/22/2020 switched to no-loop meta lookup; debugged, fixed
    * 7:13 AM 7/22/2020 replaced codeblock w get-TenantTag(), flipped ExAdmin fr switch to un-typed
    * 5:11 PM 7/21/2020 added VEN support
    * 12:20 PM 5/27/2020 moved aliases: Add-EMSRemote,cx10 win func
    * 10:13 AM 5/15/2020 with vpn AD Ex lookup issue, patched in backup pass of get-ExchangeServerFromExGroup, in case of fail ; added failthrough to updated get-ExchangeServerFromExGroup, and finally to profile $smtpserver
    * 10:19 AM 2/24/2020 Connect-Ex2010/-OBS v1.1.0: updated cx10 to reflect infra file cred name change: cred####SID -> cred###SID, debugged, working, updated output banner to draw from global session, rather than imported module (was blank output). Ren'ing this one to the primary vers, and the prior to -OBS. Changed attribution, other than function names & concept, none of the code really sources back to Mike's original any more.
    * 6:59 PM 1/15/2020 cleanup
    * 7:51 AM 12/5/2019 Connect-Ex2010:retooled $ExAdmin variant webpool support - now has detect in the server-pick logic, and on failure, it retries to the stock pool.
    * 8:55 AM 11/27/2019 expanded $Credential support to switch to torolab & - potentiall/uncfg'd - CMW mail infra. Fw seems to block torolab access (wtf)
    * # 7:54 AM 11/1/2017 add titlebar tag & updated example to test for pres of Add-PSTitleBar
    * 12:09 PM 12/9/2016 implented and debugged as part of verb-Ex2010 set
    * 2:37 PM 12/6/2016 ported to local EMSRemote
    * 2/10/14 posted version
    $Credential can leverage a global: $Credential = $global:SIDcred
    .DESCRIPTION
    Connect-Ex2010 - Setup Remote Exch2010 Mgmt Shell connection
    This supports Non-Restricted IIS custom pools, which are created via create-EMSOpenRemotePool.ps1
    .PARAMETER  ExchangeServer
    Exch server to Remote to
    .PARAMETER  ExAdmin
    Use exadmin IIS WebPool for remote EMS[-ExAdmin]
    .PARAMETER  Credential
    Credential object
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    # -----------
    try{
        $reqMods="Connect-Ex2010;Reconnect-Ex2010;Disconnect-Ex2010;Get-ExchangeServerInSite;Disconnect-PssBroken;Add-PSTitleBar".split(";") ;
        $reqMods | % {if( !(test-path function:$_ ) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing $($_) function. EXITING." } } ;
        Reconnect-Ex2010 ;
    } CATCH {
        Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
        Exit #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ;
    } ;

    # -----------
    .EXAMPLE
    # -----------
    $rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" ;
    $rgxRemsPssName="^(Exchange2010|Session\sfor\simplicit\sremoting\smodule\sat\s.*)" ;
    $rgxSnapPssname="^Session\d{1}$" ;
    $rgxEx2010SnapinName="^Microsoft\.Exchange\.Management\.PowerShell\.E2010$";
    $Ex2010SnapinName="Microsoft.Exchange.Management.PowerShell.E2010" ;
    $Error.Clear() ;
    TRY {
    if(($host.version.major -lt 3) -AND (get-service MSExchangeADTopology -ea SilentlyContinue)){
        if (!(Get-PSSnapin | where {$_.Name -match $rgxEx2010SnapinName})) {Add-PSSnapin $Ex2010SnapinName -ea Stop} ;
            write-verbose -verbose:$bshowVerbose  "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Using Local Server EMS10 Snapin" ;
            $Global:E10IsDehydrated=$false ;
        } else {
            $reqMods="Connect-Ex2010;Reconnect-Ex2010;Disconnect-Ex2010;Get-ExchangeServerInSite;Disconnect-PssBroken;Cleanup;Add-PSTitleBar;Remove-PSTitleBar".split(";") ;
            $reqMods | % {if( !(test-path function:$_ ) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing $($_) function. EXITING." } } ;
            if(!(Get-PSSession |?{$_.ComputerName -match "^(adl|spb|lyn|bcc)ms\d{3}\.global\.ad\.toro\.com$" -AND $_.ConfigurationName -eq "Microsoft.Exchange" -AND $_.Name -eq "Exchange2010" -AND $_.State -eq "Opened" -AND $_.Availability -eq "Available"})){
    reconnect-Ex2010 ;
            $Global:E10IsDehydrated=$true ;
        } else {
          write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Functional REMS connection found. " ;
        } ;
    } ;
    get-exchangeserver | out-null ;
    # -----------
    More detailed REMS & server-EMS snapin coexistince version.
    .EXAMPLE
    # -----------
    if(!(Get-PSSnapin | where {$_.Name -match $rgxEx2010SnapinName})){
        Do {
            write-host "." -NoNewLine;Start-Sleep -m (1000 * 5)
            if( !(Get-PSSession|?{$_.Name -match $rgxRemsPssName -AND $_.ComputerName -match $rgxProdEx2010ServersFqdn -AND $_.State -eq 'Opened' -AND $_.Availability -eq 'Available'}) ){
                    Reconnect-Ex2010 ;
            } ;
        } Until ((Get-PSSession|?{($_.Name -match $rgxRemsPssName -AND $_.ComputerName -match $rgxProdEx2010ServersFqdn) -AND ($_.State -eq 'Opened') -AND ($_.Availability -eq 'Available')}))
    } ;
    # -----------
    Looping reconnect test example ; defers to existing Snapin (which should be self-maintaining)
    .LINK
    https://github.com/tostka/verb-Ex2010/
    #>
    #[CmdletBinding()] # 10:03 AM 4/21/2021 disable, see if it kills verbose
    [Alias('Add-EMSRemote','cx10','cxOP','connect-ExOP')]
    Param(
        [Parameter(Position = 0, HelpMessage = "Exch server to Remote to")]
            [string]$ExchangeServer,
        [Parameter(HelpMessage = 'Use exadmin IIS WebPool for remote EMS[-ExAdmin]')]
            $ExAdmin,
        [Parameter(HelpMessage = 'Credential object')][System.Management.Automation.PSCredential]
            $Credential = $credTORSID
    )  ;
    BEGIN{
        #$verbose = ($VerbosePreference -eq "Continue") ;
		$CurrentVersionTlsLabel = [Net.ServicePointManager]::SecurityProtocol ; # Tls, Tls11, Tls12 ('Tls' == TLS1.0)  ;
        write-verbose "PRE: `$CurrentVersionTlsLabel : $($CurrentVersionTlsLabel )" ;
        # psv6+ already covers, test via the SslProtocol parameter presense
        if ('SslProtocol' -notin (Get-Command Invoke-RestMethod).Parameters.Keys) {
            $currentMaxTlsValue = [Math]::Max([Net.ServicePointManager]::SecurityProtocol.value__,[Net.SecurityProtocolType]::Tls.value__) ;
            write-verbose "`$currentMaxTlsValue : $($currentMaxTlsValue )" ;
            $newerTlsTypeEnums = [enum]::GetValues('Net.SecurityProtocolType') | Where-Object { $_ -gt $currentMaxTlsValue }
            if($newerTlsTypeEnums){
                write-verbose "Appending upgraded/missing TLS `$enums:`n$(($newerTlsTypeEnums -join ','|out-string).trim())" ;
            } else {
                write-verbose "Current TLS `$enums are up to date with max rev available on this machine" ;
            };
            $newerTlsTypeEnums | ForEach-Object {
                [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $_
            } ;
        } ;
        
        #region CHKPREREQ ; #*------v CHKPREREQ v------
        # critical dependancy Meta variables
        $MetaNames = ,'TOR','CMW','TOL' #,'NOSUCH' ; 
        # critical dependancy Meta variable properties
        $MetaProps = 'Ex10Server','Ex10WebPoolVariant','ExRevision','ExViewForest','ExOPAccessFromToro','legacyDomain' #,'DOESNTEXIST' ; 
        # critical dependancy parameters
        $gvNames = 'Credential' 
        $isBased = $true ; $gvMiss = @() ; $ppMiss = @() ; 
        foreach($met in $metanames){
            write-verbose "chk:`$$($met)Meta" ; 
            if(-not (gv -name "$($met)Meta" -ea 0)){$isBased = $false; $gvMiss += "$($met)Meta" } ; 
            if($MetaProps){
                foreach($mp in $MetaProps){
                    write-verbose "chk:`$$($met)Meta.$($mp)" ; 
                    #if(-not (gv -name "$($met)Meta" -ea 0).value[$mp]){
                    if(-not (gv -name "$($met)Meta" -ea 0).value.keys -contains $mp){
                        $isBased = $false; $ppMiss += "$($met)Meta.$($mp)" 
                    } ; 
                } ; 
            } ; 
        } ; 
        if($gvNames){
            foreach($gvN in $gvNames){
                write-verbose "chk:`$$($gvN)" ; 
                if(-not (gv -name "$($gvN)" -ea 0)){$isBased = $false; $gvMiss += "$($gvN)" } ; 
            } ; 
        } ; 
        if($gvmiss){write-warning "Missing Dependant Meta variables:`n$(($gvMiss |%{"`$$($_)" }) -join ',')" } ; 
        if($ppMiss){write-warning "Missing Dependant Meta vari properties:`n$(($ppMiss |%{"`$$($_)" }) -join ',')" } ; 
        if(-not $isBased){ write-warning  "missing critical dependancy profile config!" } ; 
        #endregion CHKPREREQ ; #*------^ END CHKPREREQ ^------


        $sWebPoolVariant = "exadmin" ;
        $CommandPrefix = $null ;
        # use credential domain to determine target org
        $rgxLegacyLogon = '\w*\\\w*' ;

        #region CONNEXOPTDO ; #*------v  v------
        #*------v Function Connect-ExchangeServerTDO v------
        #if(-not(get-command Connect-ExchangeServerTDO -ea SilentlyContinue)){
            Function Connect-ExchangeServerTDO {
                <#
                .SYNOPSIS
                Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
                will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellREmote (REMS) connect to each server, 
                stopping at the first successful connection.
                .NOTES
                Version     : 3.0.3
                Author      : Todd Kadrie
                Website     : http://www.toddomation.com
                Twitter     : @tostka / http://twitter.com/tostka
                CreatedDate : 2024-05-30
                FileName    : Connect-ExchangeServerTDO.ps1
                License     : (none-asserted)
                Copyright   : (none-asserted)
                Github      : https://github.com/tostka/verb-Ex2010
                Tags        : Powershell, ActiveDirectory, Exchange, Discovery
                AddedCredit : Brian Farnsworth
                AddedWebsite: https://codeandkeep.com/
                AddedTwitter: URL
                AddedCredit : David Paulson
                AddedWebsite: https://techcommunity.microsoft.com/t5/exchange-team-blog/exchange-health-checker-has-a-new-home/ba-p/2306671
                AddedTwitter: URL
                REVISIONS
                * 12:49 PM 6/21/2024 flipped PSS Name to Exchange$($ExchVers[dd])
                * 12:57 PM 6/11/2024 Validated, Ex2010 & Ex2019, hub, mail & edge roles: tested ☑️ on CMW mail role (Curly); and Jumpbox; 
                    copied in CBH from repo copy, which has been updated/debugged compat on CMW Edge 
                    includes local snapin detect & load for edge role (simplest EMS load option for Edge role, from David Paulson's original code; no longer published with Ex2010 compat)
                * 11:28 AM 5/30/2024 fixed failure to recognize existing functional PSSession; Made substantial update in logic, validate works fine with other orgs, and in our local orgs.
                * 4:02 PM 8/28/2023 debuged, updated CBH, renamed connect-ExchangeSErver -> Connect-ExchangeServerTDO (avoid name clashes, pretty common verb-noun combo).
                * 12:36 PM 8/24/2023 init

                .DESCRIPTION
                Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
                will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellRemote (REMS) connect to each server, 
                stopping at the first successful connection.

                Relies upon/requires get-ADExchangeServerTDO(), to return a descriptive summary of the Exchange server(s) revision etc, for connectivity logic.
                Supports Exchange 2010 through 2019, as implemented.
            
                Intent, as contrasted with verb-EXOP/Ex2010 is to have no local module dependancies, when running EXOP into other connected orgs, where syncing profile & supporting modules code can be problematic. 
                This uses native ADSI calls, which are supported by Windows itself, without need for external ActiveDirectory module etc.

                The particular approach inspired by BF's demo func that accompanied his take on get-adExchangeServer(), which I hybrided with my own existing code for cred-less connectivity. 
                I added get-OrganizationConfig testing, for connection pre/post confirmation, along with Exchange Server revision code for continutional handling of new-pssession remote powershell EMS connections.
                Also shifted connection code into _connect-EXOP() internal func.
                As this doesn't rely on local module presnece, it doesn't have to do the usual local remote/local invocation detection you'd do for non-dehydrated on-server EMS (more consistent this way, anyway; 
                there are only a few cmdlet outputs I'm aware of, that have fundementally broken returns dehydrated, and require local non-remote EMS use to function.

                My core usage would be to paste the function into the BEGIN{} block for a given remote org process, to function as a stricly local ad-hoc function.
                .PARAMETER name
                FQDN of a specific Exchange server[-Name EXSERVER.DOMAIN.COM]
                .PARAMETER discover
                Boolean paraameter that drives auto-discovery of target Exchange servers for connection (defaults `$true)[-discover:`$false]
                .PARAMETER credential
                Use specific Credentials[-Credentials [credential object]
                    .PARAMETER Site
                Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-Site 'SITENAME']
                .PARAMETER RoleNames
                Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']
                .PARAMETER TenOrg
                Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
                .INPUTS
                None. Does not accepted piped input.(.NET types, can add description)
                .OUTPUTS
                [system.object] Returns a system object containing a successful PSSession
                System.Boolean
                [| get-member the output to see what .NET obj TypeName is returned, to use here]
                System.Array of System.Object's
                .EXAMPLE
                PS> $PSSession = Connect-ExchangeServerTDO -siteName SITENAME -RoleNames @('HUB','CAS') -verbose 
                Demo's connecting to a functional Hub or CAS server in the SITENAME site with verbose outputs, the `PSSession variable will contain information about the successful connection. Makes automatic Exchangeserver discovery calls into AD (using ADSI) leveraging the separate get-ADExchangeServerTDO()
                .EXAMPLE
                PS> TRY{$Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name}CATCH{$Site=$env:COMPUTERNAME} ;
                PS> $PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
                Demo including support for EdgeRole, which is detected on it's lack of AD Site specification (which gets fed through to call, by setting the Site to the machine itself).
                .LINK
                https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/
                .LINK
                https://github.com/Lucifer1993/PLtools/blob/main/HealthChecker.ps1
                .LINK
                https://microsoft.github.io/CSS-Exchange/Diagnostics/HealthChecker/
                .LINK
                https://bitbucket.org/tostka/powershell/
                .LINK
                https://github.com/tostka/verb-Ex2010
                #>        
                [CmdletBinding(DefaultParameterSetName='discover')]
                PARAM(
                    [Parameter(Position=0,Mandatory=$true,ParameterSetName='name',HelpMessage="FQDN of a specific Exchange server[-Name EXSERVER.DOMAIN.COM]")]
                        [String]$name,
                    [Parameter(Position=0,ParameterSetName='discover',HelpMessage="Boolean paraameter that drives auto-discovery of target Exchange servers for connection (defaults `$true)[-discover:`$false]")]
                        [bool]$discover=$true,
                    [Parameter(Position=1,HelpMessage = "Use specific Credentials[-Credentials [credential object]")]
                        [Management.Automation.PSCredential]$credential,
                    [Parameter(Position=1,HelpMessage="Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-Site 'SITENAME']")]
                        [Alias('Site')]
                        [string]$SiteName = (gv -name "$($TenOrg)_ADSiteDefaultName" -ea 0).Value,
                    [Parameter(Position=2,HelpMessage="Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']")]
                        [ValidateSet('MBX','CAS','HUB','UM','MBX','EDGE')]
                        [string[]]$RoleNames = @('HUB','CAS'),
                    [Parameter(Mandatory=$FALSE,HelpMessage="Tenant Tag (3-letter abbrebiation - defaults to global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']")]
                        [ValidateNotNullOrEmpty()]
                        [string]$TenOrg = $global:o365_TenOrgDefault
                ) ;
                BEGIN{
                    $Verbose = ($VerbosePreference -eq 'Continue') ;
                    write-verbose "#*------v Function _connect-ExOP v------" ;
                    function _connect-ExOP{
                        [CmdletBinding()]
                        PARAM(
                            [Parameter(Position=0,Mandatory=$true,HelpMessage="Exchange server AD Summary system object[-Server EXSERVER.DOMAIN.COM]")]
                                [system.object]$Server,
                            [Parameter(Position=1,HelpMessage = "Use specific Credentials[-Credentials [credential object]")]
                                [Management.Automation.PSCredential]$credential
                        );
                        $verbose = $($VerbosePreference -eq "Continue") ;
                        if([double]$ExVersNum = [regex]::match($Server.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
                            switch -regex ([string]$ExVersNum) {
                                '15.2' { $isEx2019 = $true ; $ExVers = 'Ex2019' }
                                '15.1' { $isEx2016 = $true ; $ExVers = 'Ex2016'}
                                '15.0' { $isEx2013 = $true ; $ExVers = 'Ex2013'}
                                '14.*' { $isEx2010 = $true ; $ExVers = 'Ex2010'}
                                '8.*' { $isEx2007 = $true ; $ExVers = 'Ex2007'}
                                '6.5' { $isEx2003 = $true ; $ExVers = 'Ex2003'}
                                '6' {$isEx2000 = $true ; $ExVers = 'Ex2000'} ;
                                default {
                                    $smsg = "UNRECOGNIZED ExVersNum.Major.Minor string:$($ExVersNum)! ABORTING!" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    THROW $SMSG ;
                                    BREAK ;
                                }
                            } ;
                        }else {
                            $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$Server.version:$($Server.version)!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            throw $smsg ;
                            break ;
                        } ;
                        if($Server.RoleNames -eq 'EDGE'){
                            if(($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup')) -or
                                ($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup')) -or
                                $ByPassLocalExchangeServerTest)
                            {
                                if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole') -or
                                     (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'))
                                {
                                    write-verbose ("We are on Exchange Edge Transport Server")
                                    $IsEdgeTransport = $true
                                }
                                TRY {
                                    Get-ExchangeServer -ErrorAction Stop | Out-Null
                                    write-verbose "Exchange PowerShell Module already loaded."
                                    $passed = $true 
                                }CATCH {
                                    write-verbose ("Failed to run Get-ExchangeServer")
                                    if($isLocalExchangeServer){
                                        write-host  "Loading Exchange PowerShell Module..."
                                        TRY{
                                            if($IsEdgeTransport){
                                                # implement local snapins access on edge role: Only way to get access to EMS commands.
                                                [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exshell.psc1" -ErrorAction Stop
                                                ForEach($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn){
                                                    write-verbose ("Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name)
                                                    Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop
                                                } ; 
                                                Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop ; 
                                                $passed = $true #We are just going to assume this passed.
                                            }else{
                                                Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                                                Connect-ExchangeServer -Auto -ClientApplication:ManagementShell
                                                $passed = $true #We are just going to assume this passed.
                                            } 
                                        }CATCH {
                                            write-host ("Failed to Load Exchange PowerShell Module...")
                                        }                               
                                    } ;
                                } FINALLY {
                                    if($LoadExchangeVariables -and $passed -and $isLocalExchangeServer){
                                        if($ExInstall -eq $null -or $ExBin -eq $null){
                                            if(Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup'){
                                                $Global:ExInstall = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath
                                            }else{
                                                $Global:ExInstall = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath
                                            }
            
                                            $Global:ExBin = $Global:ExInstall + "\Bin"
            
                                            write-verbose ("Set ExInstall: {0}" -f $Global:ExInstall)
                                            write-verbose ("Set ExBin: {0}" -f $Global:ExBin)
                                        }
                                    }
                                }
                            } else  {
                                write-verbose ("Does not appear to be an Exchange 2010 or newer server.")
                            }
                            if(get-command -Name Get-OrganizationConfig -ea 0){
                                $smsg = "Running in connected/Native EMS" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                Return $true ; 
                            } else { 
                                TRY{
                                    $smsg = "Initiating Edge EMS local session (exshell.psc1 & exchange.ps1)" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                    # 5;36 PM 5/30/2024 didn't work, went off to nowhere for a long time, and exited the script
                                    #& (gcm powershell.exe).path -PSConsoleFile "$($env:ExchangeInstallPath)bin\exshell.psc1" -noexit -command ". '$($env:ExchangeInstallPath)bin\Exchange.ps1'"
                                    <# [Adding the Transport Server to Exchange - Mark Lewis Blog](https://marklewis.blog/2020/11/19/adding-the-transport-server-to-exchange/)
                                    To access the management console on the transport server, I opened PowerShell then ran
                                    exshell.psc1
                                    Followed by
                                    exchange.ps1
                                    At this point, I was able to create a new subscription using he following PowerShel
                                    #>
                                    invoke-command exshell.psc1 ; 
                                    invoke-command exchange.ps1
                                    if(get-command -Name Get-OrganizationConfig -ea 0){
                                        $smsg = "Running in connected/Native EMS" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                        Return $true ;
                                    } else { return $false };  
                                }CATCH{
                                    Write-Error $_ ;
                                } ;
                            } ; 
                        } else {
                            $pltNPSS=@{ConnectionURI="http://$($Server.FQDN)/powershell"; ConfigurationName='Microsoft.Exchange' ; name="Exchange$($ExVersNum.tostring())"} ;
                            # use ExVersUnm dd instead of hardcoded (Exchange2010)
                            if($ExVersNum -ge 15){
                                write-verbose "EXOP.15+:Adding -Authentication Kerberos" ;
                                $pltNPSS.add('Authentication',"Kerberos") ;
                                $pltNPSS.name = $ExVers ;
                            } ;
                            $smsg = "Adding EMS (connecting to $($Server.FQDN))..." ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $smsg = "New-PSSession w`n$(($pltNPSS|out-string).trim())" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $ExPSS = New-PSSession @pltNPSS  ;
                            $ExIPSS = Import-PSSession $ExPSS -allowclobber ;
                            $ExPSS | write-output ;
                            $ExPSS= $ExIPSS = $null ;
                        } ; 
                    } ;
                    write-verbose "#*------^ END Function _connect-ExOP ^------" ;
                    $pltGADX=@{
                        ErrorAction='Stop';
                    } ;
                } ;
                PROCESS{
                    if($PSBoundParameters.ContainsKey('credential')){
                        $pltGADX.Add('credential',$credential) ;
                    }
                    if($SiteName){
                        $pltGADX.Add('siteName',$siteName) ;
                    } ;
                    if($RoleNames){
                        $pltGADX.Add('RoleNames',$RoleNames) ;
                    } ;
                    TRY{
                        if($discover){
                            $smsg = "Getting list of Exchange Servers" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $exchServers=get-ADExchangeServerTDO @pltGADX | sort responsetime ;
                        }else{
                            $exchServers=get-ADExchangeServerTDO @pltGADX | sort responsetime ;
                        } ;
                        $pltTW=@{
                            'ErrorAction'='Stop';
                        } ;
                        $pltCXOP = @{
                            verbose = $($VerbosePreference -eq "Continue") ;
                        } ;
                        if($pltGADX.credential){
                            $pltCXOP.Add('Credential',$pltCXOP.Credential) ;
                        } ;
                        $prpPSS = 'Id','Name','ComputerName','ComputerType','State','ConfigurationName','Availability' ; 
                        foreach($exServer in $exchServers){
                            write-verbose "testing conn to:$($exServer.name.tostring())..." ; 
                            if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig' -ea SilentlyContinue){
                                if($pssEXOP = Get-PSSession |  where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ( $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND ( $_.runspace.ConnectionInfo.Port -eq '80') }){
                                    if($pssEXOP.State -ne "Opened" -OR $pssEXOP.Availability -ne "Available"){
                                        $pssEXOP | remove-pssession ; $pssEXOP = $null ;
                                    } ;
                                } ; 
                            } else {
                                $smsg = "(mangled ExOP conn: disconnect/reconnect...)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                if($pssEXOP = Get-PSSession |  where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ( $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND ( $_.runspace.ConnectionInfo.Port -eq '80') }){
                                    if($pssEXOP.State -ne "Opened" -OR $pssEXOP.Availability -ne "Available"){
                                        $pssEXOP | remove-pssession ; $pssEXOP = $null ;
                                    } ;
                                } ; 
                            } ;
                            if(-not $pssEXOP){
                                $smsg = "Connecting to: $($exServer.FQDN)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                if($NoTest){
                                    $ExPSS =$ExPSS = _connect-ExOP @pltCXOP -Server $exServer
                               } else {
                                    TRY{
                                        $smsg = "Testing Connection: $($exServer.FQDN)" ;
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                        If(test-connection $exServer.FQDN -count 1 -ea 0) {
                                            $smsg = "confirmed pingable..." ;
                                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                        } else {
                                            $smsg = "Unable to Ping $($exServer.FQDN)" ; ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        } ;
                                        $smsg = "Testing WinRm: $($exServer.FQDN)" ;
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                        $winrm=Test-WSMan @pltTW -ComputerName $exServer.FQDN ;
                                        if($winrm){
                                            $ExPSS = _connect-ExOP @pltCXOP -Server $exServer;
                                        } else {
                                            $smsg = "Unable to Test-WSMan $($exServer.FQDN) (skipping)" ; ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        } ;
                                    }CATCH{
                                        $errMsg="Server: $($exServer.FQDN)] $($_.Exception.Message)" ;
                                        Write-Error -Message $errMsg ;
                                        continue ;
                                    } ;
                                };
                            } else {
                                $smsg = "$((get-date).ToString('HH:mm:ss')):Accepting first valid connection w`n$(($pssEXOP | ft -a $prpPSS|out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $ExPSS = $pssEXOP ; 
                                break ; 
                            }  ;
                        } ;
                    }CATCH{
                        Write-Error $_ ;
                    } ;
                } ;
                END{
                    if(-not $ExPSS){
                        $smsg = "NO SUCCESSFUL CONNECTION WAS MADE, WITH THE SPECIFIED INPUTS!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $smsg = "(returning `$false to the pipeline...)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        return $false
                    } else{
                        if($ExPSS.State -eq "Opened" -AND $ExPSS.Availability -eq "Available"){
                            if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                                $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ;
                                throw $smsg ;
                                $smsg | write-warning  ;
                            } else {
                                $smsg = "(connected to EXOP.Org:$($orgName))" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                            return $ExPSS
                        } ;
                    } ; 
                } ;
            } ;
        #} ; 
        #*------^ END Function Connect-ExchangeServerTDO ^------
        #endregion CONNEXOPTDO ; #*------^ END CONNEXOPTDO ^------
    
        #region GADEXSERVERTDO ; #*------v  v------
        #*------v Function get-ADExchangeServerTDO v------
        #if(-not(get-command get-ADExchangeServerTDO -ea SilentlyContinue)){
            Function get-ADExchangeServerTDO {
                <#
                .SYNOPSIS
                get-ADExchangeServerTDO.ps1 - Returns Exchangeserver summary(s) from AD records
                .NOTES
                Version     : 3.0.1
                Author      : Todd Kadrie
                Website     : http://www.toddomation.com
                Twitter     : @tostka / http://twitter.com/tostka
                CreatedDate : 2015-09-03
                FileName    : get-ADExchangeServerTDO.ps1
                License     : (none-asserted)
                Copyright   : (none-asserted)
                Github      : https://github.com/tostka/verb-Ex2010
                Tags        : Powershell, ActiveDirectory, Exchange, Discovery
                AddedCredit : Mike Pfeiffer
                AddedWebsite: mikepfeiffer.net
                AddedTwitter: URL
                AddedCredit : Sammy Krosoft 
                AddedWebsite: http://aka.ms/sammy
                AddedTwitter: URL
                AddedCredit : Brian Farnsworth
                AddedWebsite: https://codeandkeep.com/
                AddedTwitter: URL
                REVISIONS
                * 12:57 PM 6/11/2024 Validated, Ex2010 & Ex2019, hub, mail & edge roles: tested ☑️ on CMW mail role (Curly); and Jumpbox; copied in CBH from repo copy, which has been updated/debugged compat on CMW Edge 
                * 2:05 PM 8/28/2023 REN -> Get-ExchangeServerInSite -> get-ADExchangeServerTDO (aliased orig); to better steer profile-level options - including in cmw org, added -TenOrg, and default Site to constructed vari, targeting new profile $XXX_ADSiteDefault vari; Defaulted -Roles to HUB,CAS as well.
                * 3:42 PM 8/24/2023 spliced together combo of my long-standing, and some of the interesting ideas BF's version had. Functional prod:
                    - completely removed ActiveDirectory module dependancies from BF's code, and reimplemented in raw ADSI calls. Makes it fully portable, even into areas like Edge DMZ roles, where ADMS would never be installed.

                * 3:17 PM 8/23/2023 post Edge testing: some logic fixes; add: -Names param to filter on server names; -Site & supporting code, to permit lookup against sites *not* local to the local machine (and bypass lookup on the local machine) ; 
                    ren $Ex10siteDN -> $ExOPsiteDN; ren $Ex10configNC -> $ExopconfigNC
                * 1:03 PM 8/22/2023 minor cleanup
                * 10:31 AM 4/7/2023 added CBH expl of postfilter/sorting to draw predictable pattern 
                * 4:36 PM 4/6.2.33 validated Psv51 & Psv20 and Ex10 & 16; added -Roles & -RoleNames params, to perform role filtering within the function (rather than as an external post-filter step). 
                For backward-compat retain historical output field 'Roles' as the msexchcurrentserverroles summary integer; 
                use RoleNames as the text role array; 
                    updated for psv2 compat: flipped hash key lookups into properties, found capizliation differences, (psv2 2was all lower case, wouldn't match); 
                flipped the [pscustomobject] with new... psobj, still psv2 doesn't index the hash keys ; updated for Ex13+: Added  16  "UM"; 20  "CAS, UM"; 54  "MBX" Ex13+ ; 16385 "CAS" Ex13+ ; 16439 "CAS, HUB, MBX" Ex13+
                Also hybrided in some good ideas from SammyKrosoft's Get-SKExchangeServers.psm1 
                (emits Version, Site, low lvl Roles # array, and an array of Roles, for post-filtering); 
                # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
                * 12:08 PM 5/15/2020 fixed vpn issue: Try/Catch'd around recently failing $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName qry
                * 11:22 AM 3/13/2020 Get-ExchangeServerInSite added a ping-test, to only return matches that are pingable, added -NoPing param, to permit (faster) untested bypass
                * 6:59 PM 1/15/2020 cleanup
                # 10:03 AM 11/16/2018 Get-ExchangeServerInSite:can't do AD-related functions when not AD authentictaed (home, pre-vpn connect). Added if/then test on status and abort balance when false.
                * 11/18/18 BF's posted rev
                # 12:10 PM 8/1/2017 updated example code at bottom, to accommodate variant sites
                # 11:28 AM 3/31/2016 validated that latest round of updates are still functional
                #1:58 PM 9/3/2015 - added pshelp and some docs
                #April 12, 2010 - web version
                .DESCRIPTION
                get-ADExchangeServerTDO.ps1 - Returns Exchangeserver summary(s) from AD records

                Hybrided together ideas from Brian Farnsworth's blog post
                [PowerShell - ActiveDirectory and Exchange Servers – CodeAndKeep.Com – Code and keep calm...](https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/)
                ... with much older concepts from  Sammy Krosoft, and much earlier Mike Pfeiffer. 

                - Subbed in MP's use of ADSI for ActiveDirectory Ps mod cmds - it's much more dependancy-free; doesn't require explicit install of the AD ps module
                ADSI support is built into windows.
                - spliced over my addition of Roles, RoleNames, Name & NoTest params, for prefiltering and suppressing testing.


                [briansworth · GitHub](https://github.com/briansworth)

                Uses an ADSI DirectorySearcher to search the current Active Directory site for Exchange on-prem servers.
                        Intent is to discover connection points for Powershell, wo the need to preload/pre-connect to Exchange.

                        But, as a non-Exchange-Management-Shell-dependant info source on Exchange Server configs, it can be used before connection, with solely AD-available data, to check configuration spes on the subject server(s). 

                        For example, this query will return sufficient data under Version to indicate which revision of Exchange is in use:


                        Returned object (in array):
                        Site      : {ADSITENAME}
                        Roles     : {64}
                        Version   : {Version 15.1 (Build 32375.7)}
                        Name      : SERVERNAME
                        RoleNames : EDGE
                        FQDN      : SERVERNAME.DOMAIN.TLD

                        ... includes the post-filterable Role property ($_.Role -contains 'CAS') which reflects the following
                        installed-roles ('msExchCurrentServerRoles') on the discovered servers
                            2   {"MBX"} # Ex10
                            4   {"CAS"}
                            16  {"UM"}
                            20  {"CAS, UM" -split ","} # 
                            32  {"HUB"}
                            36  {"CAS, HUB" -split ","}
                            38  {"CAS, HUB, MBX" -split ","}
                            54  {"MBX"} # Ex13+
                            64  {"EDGE"}
                            16385   {"CAS"} # Ex13+
                            16439   {"CAS, HUB, MBX"  -split ","} # Ex13+

                .PARAMETER Roles
                Array of msExchCurrentServerRoles 'role' integers to be filtered against (2|4|16|20|32|36|38|54|64|16385|16439)[-Roles @(38,16385)]
                .PARAMETER RoleNames
                Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']
                .PARAMETER Server
                Array of Server name strings to be filtered against[-Server 'SERVER1','SERVER2']
                .PARAMETER SiteName
                Name of specific AD SiteName to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-SiteName 'SITENAME']
                .PARAMETER TenOrg
                Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
                .PARAMETER NoPing
                Switch to suppress default 'pingable' test (e.g. returns all matches, no testing)[-NoPing]
                .INPUTS
                None. Does not accepted piped input.(.NET types, can add description)
                .OUTPUTS
                None. Returns no objects or output (.NET types)
                System.Boolean
                [| get-member the output to see what .NET obj TypeName is returned, to use here]
                System.Array of System.Object's
                .EXAMPLE
                PS> If(!($ExchangeServer)){$ExchangeServer = (get-ADExchangeServerTDO| ?{$_.RoleNames -contains 'CAS' -OR $_.RoleNames -contains 'HUB' -AND ($_.FQDN -match "^SITECODE") } | Get-Random ).FQDN
                Return a random Hub Cas Role server in the local Site with a fqdn beginning SITECODE
                .EXAMPLE
                PS> $localADExchserver = get-ADExchangeServerTDO -Names $env:computername -SiteName ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().name)
                Demo, if run from an Exchange server, return summary details about the local server (-SiteName isn't required, is default imputed from local server's Site, but demos explicit spec for remote sites)
                .EXAMPLE
                PS> $regex = '(' + [regex]($ADSiteCodeUK,$ADSiteCodeAU -join "|") + ')'
                PS> switch -regex ($($env:computername).substring(0,3)){
                PS>    "$($ADSiteCodeUS)" {$tExRole=36 } ;
                PS>    "$($regex)" {$tExRole= 38 }  default { write-error "$((get-date).ToString('HH:mm:ss')):UNRECOGNIZED SERVER PREFIX!."; } ;
                PS> } ;
                PS> $exhubcas = (get-ADExchangeServerTDO |?{($_.roles -eq $tExRole) -AND ($_.FQDN -match "$($env:computername.substring(0,3)).*")} | Get-Random ).FQDN ;
                Use a switch block to select different role combo targets for a given server fqdn prefix string.
                .EXAMPLE
                PS> $ExchangeServer = get-ADExchangeServerTDO | ?{$_.Roles -match '(4|20|32|36|38|16385|16439)'} | select -expand fqdn | get-random ; 
                Another/Older approach filtering on the Roles integer (targeting combos with Hub or CAS in the mix)
                .EXAMPLE
                PS> $ret = get-ADExchangeServerTDO -Roles @(4,20,32,36,38,16385,16439) -verbose 
                Demo use of the -Roles param, feeding it an array of Role integer values to be filtered against. In this case, the Role integers that include a CAS or HUB role.
                .EXAMPLE
                PS> $ret = get-ADExchangeServerTDO -RoleNames 'HUB','CAS' -verbose ;
                Demo use of the -RoleNames param, feeding it the array 'HUB','CAS' Role name strings to be filtered against
                PS> $ret = get-ADExchangeServerTDO -Names 'SERVERName' -verbose ;
                Demo use of the -RoleNames param, feeding it the array 'HUB','CAS' Role name strings to be filtered against
                .EXAMPLE
                PS> $ExchangeServer = get-ADExchangeServerTDO | sort version,roles,name | ?{$_.rolenames -contains 'CAS'}  | select -last 1 | select -expand fqdn ;
                Demo post sorting & filtering, to deliver a rule-based predictable pattern for server selection: 
                Above will always pick the highest Version, 'CAS' RoleName containing, alphabetically last server name (that is pingable). 
                And should stick to that pattern, until the servers installed change, when it will shift to the next predictable box.
                .EXAMPLE
                PS> $ExOPServer = get-ADExchangeServerTDO -Name LYNMS650 -SiteName Lyndale
                PS> if([double]$ExVersNum = [regex]::match($ExOPServer.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
                PS>     switch -regex ([string]$ExVersNum) {
                PS>         '15\.2' { $isEx2019 = $true ; $ExVers = 'Ex2019' }
                PS>         '15\.1' { $isEx2016 = $true ; $ExVers = 'Ex2016'}
                PS>         '15\.0' { $isEx2013 = $true ; $ExVers = 'Ex2013'}
                PS>         '14\..*' { $isEx2010 = $true ; $ExVers = 'Ex2010'}
                PS>         '8\..*' { $isEx2007 = $true ; $ExVers = 'Ex2007'}
                PS>         '6\.5' { $isEx2003 = $true ; $ExVers = 'Ex2003'}
                PS>         '6|6\.0' {$isEx2000 = $true ; $ExVers = 'Ex2000'} ;
                PS>         default {
                PS>             $smsg = "UNRECOGNIZED ExchangeServer.AdminDisplayVersion.Major.Minor string:$($ExOPServer.version)! ABORTING!" ;
                PS>             write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                PS>         }
                PS>     } ; 
                PS> }else {
                PS>     $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$ExOPServer.version:$($ExOPServer.version)!" ; 
                PS>     write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ; 
                PS>     throw $smsg ; 
                PS>     break ; 
                PS> } ; 
                Demo of parsing the returned Version property, into the proper Exchange Server revision.      
                .LINK
                https://github.com/tostka/verb-XXX
                .LINK
                https://bitbucket.org/tostka/powershell/
                .LINK
                http://mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
                .LINK
                https://github.com/SammyKrosoft/Search-AD-Using-Plain-PowerShell/blob/master/Get-SKExchangeServers.psm1
                .LINK
                https://github.com/tostka/verb-Ex2010
                .LINK
                https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/
                #>
                [CmdletBinding()]
                [Alias('Get-ExchangeServerInSite')]
                PARAM(
                    [Parameter(Position=0,HelpMessage="Array of Server name strings to be filtered against[-Server 'SERVER1','SERVER2']")]
                        [string[]]$Server,
                    [Parameter(Position=1,HelpMessage="Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-SiteName 'SITENAME']")]
                        [Alias('Site')]
                        [string]$SiteName = (gv -name "$($TenOrg)_ADSiteDefaultName" -ea 0).Value,
                    [Parameter(Position=2,HelpMessage="Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']")]
                        [ValidateSet('MBX','CAS','HUB','UM','MBX','EDGE')]
                        [string[]]$RoleNames = @('HUB','CAS'),
                    [Parameter(HelpMessage="Array of msExchCurrentServerRoles 'role' integers to be filtered against (2|4|16|20|32|36|38|54|64|16385|16439)[-Roles @(38,16385)]")]
                        [ValidateSet(2,4,16,20,32,36,38,54,64,16385,16439)]
                        [int[]]$Roles,
                    [Parameter(HelpMessage="Switch to suppress default 'pingable' test (e.g. returns all matches, no testing)[-NoTest]")]
                        [Alias('NoPing')]
                        [switch]$NoTest,
                    [Parameter(HelpMessage="Milliseconds of max timeout to wait during port 80 test (defaults 100)[-SpeedThreshold 500]")]
                        [int]$SpeedThreshold=100,
                    [Parameter(Mandatory=$FALSE,HelpMessage="Tenant Tag (3-letter abbrebiation - defaults to global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']")]
                        [ValidateNotNullOrEmpty()]
                        [string]$TenOrg = $global:o365_TenOrgDefault,
                    [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials[-Credentials [credential object]]")]
                        [System.Management.Automation.PSCredential]$Credential
                ) ;
                BEGIN{
                    ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
                    $Verbose = ($VerbosePreference -eq 'Continue') ;
                    $_sBnr="#*======v $(${CmdletName}): v======" ;
                    $smsg = $_sBnr ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
                PROCESS{
                    TRY{
                        $configNC = ([ADSI]"LDAP://RootDse").configurationNamingContext ;
                        $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC") ;
                        $bLocalEdge = $false ; 
                        if($Sitename -eq $env:COMPUTERNAME){
                            $smsg = "`$SiteName -eq `$env:COMPUTERNAME:$($SiteName):$($env:COMPUTERNAME)" ; 
                            $smsg += "`nThis computer appears to be an EdgeRole system (non-ADConnected)" ; 
                            $smsg += "`n(Blanking `$sitename and continuing discovery)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            #$bLocalEdge = $true ; 
                            $SiteName = $null ; 
                        
                        } ; 
                        If($siteName){
                            $smsg = "WVGetting Site: $siteName" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $objectClass = "objectClass=site" ;
                            $objectName = "name=$siteName" ;
                            $search.Filter = "(&($objectClass)($objectName))" ;
                            $site = ($search.Findall()) ;
                            $siteDN = ($site | select -expand properties).distinguishedname  ;
                        } else {
                            $smsg = "(No -Site specified, resolving site from local machine domain-connection...)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                            else{ write-host -foregroundcolor green "$($smsg)" } ;
                            TRY{$siteDN = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().GetDirectoryEntry().distinguishedName}
                            CATCH [System.Management.Automation.MethodInvocationException]{
                                $ErrTrapd=$Error[0] ;
                                if(($ErrTrapd.Exception -match 'The computer is not in a site.') -AND $env:ExchangeInstallPath){
                                    $smsg = "$($env:computername) is non-ADdomain-connected" ;
                                    $smsg += "`nand has `$env:ExchangeInstalled populated: Likely Edge Server" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                                    else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    $vers = (get-item "$($env:ExchangeInstallPath)\Bin\Setup.exe").VersionInfo.FileVersionRaw ; 
                                    $props = @{
                                        Name=$env:computername;
                                        FQDN = ([System.Net.Dns]::gethostentry($env:computername)).hostname;
                                        Version = "Version $($vers.major).$($vers.minor) (Build $($vers.Build).$($vers.Revision))" ; 
                                        #"$($vers.major).$($vers.minor)" ; 
                                        #$exServer.serialNumber[0];
                                        Roles = [System.Object[]]64 ;
                                        RoleNames = @('EDGE');
                                        DistinguishedName =  "CN=$($env:computername),CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=First Organization,CN=Microsoft Exchange,CN=Services,CN=Configuration,CN={nnnnnnnn-FAKE-GUID-nnnn-nnnnnnnnnnnn}" ;
                                        Site = [System.Object[]]'NOSITE'
                                        ResponseTime = if($rsp){$rsp.ResponseTime} else { 0} ;
                                        NOTE = "This summary object, returned for a non-AD-connected EDGE server, *approximates* what would be returned on an AD-connected server" ;
                                    } ;
                                
                                    $smsg = "(-NoTest:Defaulting Fast:`$true)" ;
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                    $props.add('Fast',$true) ;
                                
                                    return (New-Object -TypeName PsObject -Property $props) ;
                                }elseif(-not $env:ExchangeInstallPath){
                                    $smsg = "Non-Domain Joined machine, with NO ExchangeInstallPath e-vari: `nExchange is not installed locally: local computer resolution fails:`nPlease specify an explicit -Server, or -SiteName" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    $false | write-output ;
                                } else {
                                    $smsg = "$($env:computername) is both NON-Domain-joined -AND lacks an Exchange install (NO ExchangeInstallPath e-vari)`nPlease specify an explicit -Server, or -SiteName" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    $false | write-output ;
                                };
                            } CATCH {
                                $siteDN =$ExOPsiteDN ;
                                write-warning "`$siteDN lookup FAILED, deferring to hardcoded `$ExOPsiteDN string in infra file!" ;
                            } ;
                        } ;
                        $smsg = "Getting Exservers in Site:$($siteDN)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC") ;
                        $objectClass = "objectClass=msExchExchangeServer" ;
                        $version = "versionNumber>=1937801568" ;
                        $site = "msExchServerSite=$siteDN" ;
                        $search.Filter = "(&($objectClass)($version)($site))" ;
                        $search.PageSize = 1000 ;
                        [void] $search.PropertiesToLoad.Add("name") ;
                        [void] $search.PropertiesToLoad.Add("msexchcurrentserverroles") ;
                        [void] $search.PropertiesToLoad.Add("networkaddress") ;
                        [void] $search.PropertiesToLoad.Add("msExchServerSite") ;
                        [void] $search.PropertiesToLoad.Add("serialNumber") ;
                        [void] $search.PropertiesToLoad.Add("DistinguishedName") ;
                        $exchServers = $search.FindAll() ;
                        $Aggr = @() ;
                        foreach($exServer in $exchServers){
                            $fqdn = ($exServer.Properties.networkaddress |
                                Where-Object{$_ -match '^ncacn_ip_tcp:'}).split(':')[1] ;
                            if($NoTest){} else {
                                $rsp = test-connection $fqdn -count 1 -ea 0 ;
                            } ;
                            $props = @{
                                Name = $exServer.Properties.name[0]
                                FQDN=$fqdn;
                                Version = $exServer.Properties.serialnumber
                                Roles = $exserver.Properties.msexchcurrentserverroles
                                RoleNames = $null ;
                                DistinguishedName = $exserver.Properties.distinguishedname;
                                Site = @("$($exserver.Properties.msexchserversite -Replace '^CN=|,.*$')") ;
                                ResponseTime = if($rsp){$rsp.ResponseTime} else { 0} ;
                            } ;
                            $props.RoleNames = switch ($exserver.Properties.msexchcurrentserverroles){
                                2       {"MBX"}
                                4       {"CAS"}
                                16      {"UM"}
                                20      {"CAS;UM".split(';')}
                                32      {"HUB"}
                                36      {"CAS;HUB".split(';')}
                                38      {"CAS;HUB;MBX".split(';')}
                                54      {"MBX"}
                                64      {"EDGE"}
                                16385   {"CAS"}
                                16439   {"CAS;HUB;MBX".split(';')}
                            }
                            if($NoTest){
                                $smsg = "(-NoTest:Defaulting Fast:`$true)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                $props.add('Fast',$true) ;
                            }else {
                                $props.add('Fast',[boolean]($rsp.ResponseTime -le $SpeedThreshold)) ;
                            };
                            $Aggr += New-Object -TypeName PsObject -Property $props ;
                        } ;
                        $httmp = @{} ;
                        if($Roles){
                            [regex]$rgxRoles = ('(' + (($roles |%{[regex]::escape($_)}) -join '|') + ')') ;
                            $matched =  @( $aggr | ?{$_.Roles -match $rgxRoles}) ;
                            foreach($m in $matched){
                                if($httmp[$m.name]){} else {
                                    $httmp[$m.name] = $m ;
                                } ;
                            } ;
                        } ;
                        if($RoleNames){
                            foreach ($RoleName in $RoleNames){
                                $matched = @($Aggr | ?{$_.RoleNames -contains $RoleName} ) ;
                                foreach($m in $matched){
                                    if($httmp[$m.name]){} else {
                                        $httmp[$m.name] = $m ;
                                    } ;
                                } ;
                            } ;
                        } ;
                        if($Server){
                            foreach ($Name in $Server){
                                $matched = @($Aggr | ?{$_.Name -eq $Name} ) ;
                                foreach($m in $matched){
                                    if($httmp[$m.name]){} else {
                                        $httmp[$m.name] = $m ;
                                    } ;
                                } ;
                            } ;
                        } ;
                        if(($httmp.Values| measure).count -gt 0){
                            $Aggr  = $httmp.Values ;
                        } ;
                        $smsg = "Returning $((($Aggr|measure).count|out-string).trim()) match summaries to pipeline..." ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $Aggr | write-output ;
                    }CATCH{
                        Write-Error $_ ;
                    } ;
                } ;
                END{
                    $smsg = "$($_sBnr.replace('=v','=^').replace('v=','^='))" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            } ;
        #} ; 
        #*------^ END Function get-ADExchangeServerTDO ^------ ;
        #endregion GADEXSERVERTDO ; #*------^ END GADEXSERVERTDO ^------

        if($isBased){
            $TenOrg = get-TenantTag -Credential $Credential ;
        } else { 

        } ; 
        <#
        if($Credential.username -match $rgxLegacyLogon){
            $credDom =$Credential.username.split('\')[0] ;
        } elseif ($Credential.username.contains('@')){
            $credDom = ($Credential.username.split("@"))[1] ;
        } else {
            write-warning "$((get-date).ToString('HH:mm:ss')):UNRECOGNIZED CREDENTIAL!:$($Credential.Username)`nUNABLE TO RESOLVE DEFAULT EX10SERVER FOR CONNECTION!" ;
        } ;
        #>
    } ;  # BEG-E
    PROCESS{
        if($isBased){
            $ExchangeServer=$null ;
            # flip from dyn lookup to array in Ex10Server, and always use get-random to pick between. Returns a value, even when only a single value
            $ExchangeServer = (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server|get-random ;
            $ExAdmin = (Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant ;
            $ExVers = (Get-Variable  -name "$($TenOrg)Meta").value.ExRevision ;
            $ExVwForest = (Get-Variable  -name "$($TenOrg)Meta").value.ExViewForest ;
            $ExOPAccessFromToro = (Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro
            # force unresolved to dyn
            if(!$ExchangeServer){
                #$ExchangeServer = 'dynamic' ;
                # getting away from dyn, random from array in Ex10Server
                throw "Undefined `$ExchangeServer for $($TenOrg) org, and `$$($TenOrg)Meta.Ex10Server property" ;
                Exit ;
            } ;
            if($ExchangeServer -eq 'dynamic'){
                if( $ExchangeServer = (Get-ExchangeServerInSite | Where-Object { ($_.roles -eq 36) } | Get-Random ).FQDN){}
                else {
                    write-warning "$((get-date).ToString('HH:mm:ss')):Get-ExchangeServerInSite *FAILED*,`ndeferring to Get-ExchServerFromExServersGroup" ;
                    if(!($ExchangeServer = Get-ExchServerFromExServersGroup)){
                        write-warning "$((get-date).ToString('HH:mm:ss')):Get-ExchServerFromExServersGroup *FAILED*,`n deferring to profile `$smtpserver:$($smtpserver))"  ;
                        $ExchangeServer = $smtpserver ;
                    };
                } ;
            } ;

            write-host -foregroundcolor darkgray "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Adding EMS (connecting to $($ExchangeServer))..." ;
            # splat to open a session - # stock 'PSLanguageMode=Restricted' powershell IIS Webpool
            $pltNSess = @{ConnectionURI = "http://$ExchangeServer/powershell"; ConfigurationName = 'Microsoft.Exchange' ; name = "Exchange$($ExVers)" } ;
            if($env:USERDOMAIN -ne (Get-Variable  -name "$($TenOrg)Meta").value.legacyDomain){
                # if not in the $TenOrg legacy domain - running cross-org -  add auth:Kerberos
                <#suppresses: The WinRM client cannot process the request. It cannot determine the content type of the HTTP response f rom the destination computer. The content type is absent or invalid
                #>
                $pltNSess.add('Authentication','Kerberos') ;
            } ;
            if ($ExAdmin) {
              # use variant IIS Webpool
              $pltNSess.ConnectionURI = $pltNSess.ConnectionURI.replace("/powershell", "/$($sWebPoolVariant)") ;
            }
            if ($Credential) {
                 $pltNSess.Add("Credential", $Credential)
                 write-verbose "(using cred:$($credential.username))" ;
            } ;

            # -Authentication Basic only if specif needed: for Ex configured to connect via IP vs hostname)
            # try catch against and retry into stock if fails
            $error.clear() ;
            TRY {
                $Global:EXOPSess = New-PSSession @pltNSess -ea STOP  ;
            } CATCH {
                $ErrTrapd = $_ ;
                write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                #-=-record a STATUSWARN=-=-=-=-=-=-=
                $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                if(Get-Variable passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                if(Get-Variable -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                #-=-=-=-=-=-=-=-=
                if ($ExAdmin) {
                    # switch to stock pool and retry
                    $pltNSess.ConnectionURI = $pltNSess.ConnectionURI.replace("/$($sWebPoolVariant)", "/powershell") ;
                    write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TARGETING EXADMIN POOL`nRETRY W STOCK POOL: New-PSSession w`n$(($pltNSess|out-string).trim())" ;
                    $Global:EXOPSess = New-PSSession @pltNSess -ea STOP  ;
                } else {
                    BREAK ;
                } ;
            } ;

            write-verbose "$((get-date).ToString('HH:mm:ss')):Importing Exchange 2010 Module" ;
            #$pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ; verbose=$false} ; # force verbose off, suppress spam in console
            # tear verbose out
            $pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ;} ;
            #$pltISess = [ordered]@{Session = $Global:EXOPSess ; DisableNameChecking = $true  ; AllowClobber = $true ; ErrorAction = 'Stop' ; Verbose = $false ;} ;
            $pltISess = [ordered]@{Session = $Global:EXOPSess ; DisableNameChecking = $true  ; AllowClobber = $true ; ErrorAction = 'Stop' ; } ;
            if($CommandPrefix){
                $pltIMod.add('Prefix',$CommandPrefix) ;
                $pltISess.add('Prefix',$CommandPrefix) ;
            } ;
            $smsg = "$((get-date).ToString('HH:mm:ss')):Import-PSSession  w`n$(($pltISess|out-string).trim())`nImport-Module w`n$(($pltIMod|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $error.clear() ;
            TRY {
                # 9:57 AM 4/21/2021 coming through full verbose, suppress the pref
                if($VerbosePreference -eq "Continue"){
                    $VerbosePrefPrior = $VerbosePreference ;
                    $VerbosePreference = "SilentlyContinue" ;
                    $verbose = ($VerbosePreference -eq "Continue") ;
                } ;
                #$Global:E10Mod = Import-Module (Import-PSSession @pltISess) @pltIMod   ;
                # try 2-stopping (suppress verbose)
                $xIPS = Import-PSSession @pltISess ;
                $Global:E10Mod = Import-Module $xIPS @pltIMod ;
                if($ExVwForest){
                    write-host "Setting EMS Session: Set-AdServerSettings -ViewEntireForest `$True" ;
                    Set-AdServerSettings -ViewEntireForest $True ;
                } ;
                # reenable VerbosePreference:Continue, if set, during mod loads
                if($VerbosePrefPrior -eq "Continue"){
                    $VerbosePreference = $VerbosePrefPrior ;
                    $verbose = ($VerbosePreference -eq "Continue") ;
                } ;
            } CATCH {
                $ErrTrapd = $_ ;
                write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                #-=-record a STATUSERROR=-=-=-=-=-=-=
                $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                if(Get-Variable passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                if(Get-Variable -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                #-=-=-=-=-=-=-=-=
            } ;
            # 7:54 AM 11/1/2017 add titlebar tag
            #Add-PSTitleBar 'EMS' ;
            # 1:31 PM 7/21/2021 build with TenOrg spec
            # 9:00 AM 7/27/2021 revise to support EMS[tlc] single-letter onprem conne designator (already in infra file OrgSvcs list).
            if($TenOrg){
                <# can't just use last char, lab varies from others
                if($TenOrg -ne 'TOL' ){
                    $sTitleBarTag = @("EMS$($TenOrg.substring(0,1).tolower())") ; # 1st char
                }else{
                    $sTitleBarTag = @("EMS$($TenOrg.substring(2,1).tolower())") ; # last char
                } ; 
                #>
                switch -regex ($TenOrg){
                    '^(CMW|TOR)$'{
                        $sTitleBarTag = @("EMS$($TenOrg.substring(0,1).tolower())") ; # 1st char
                    }
                    '^TOL$'{
                        $sTitleBarTag = @("EMS$($TenOrg.substring(2,1).tolower())") ; # last char
                    } ; 
                    default{
                        throw "$($TenOrg):unsupported `$TenOrg!" ; 
                        break ; 
                    }
                } ; 
            } else { 
                $sTitleBarTag = @("EMS") ;
            } ; 
            write-verbose "`$sTitleBarTag:$($sTitleBarTag)" ; 
        
            #$sTitleBarTag += $TenOrg ;
            Add-PSTitleBar $sTitleBarTag -verbose:$($VerbosePreference -eq "Continue")  ;
            # tag E10IsDehydrated
            $Global:E10IsDehydrated = $true ;
            write-host -foregroundcolor darkgray "`n$(($Global:EXOPSess | select ComputerName,Availability,State,ConfigurationName | format-table -auto |out-string).trim())" ;

        } else { 
            
            #region SERVICE_CONNECTIONS_DEPENDANCYLESS #*======v SERVICE_CONNECTIONS_DEPENDANCYLESS v======
            # DON'T RUN THIS AND THE USEXOP= BLOCK TOGETHER!
            # SIMPLE DEP-LESS VARIANT FOR EXOP-ONLY, NO DEPS ON VERB-*, OTHER THAN REQS: load-ADMs() & get-ADExchangeSErverTDO() (both should be local function includes)
            # PRETUNE STEERING separately *before* pasting in balance of region
            #*------v STEERING VARIS v------
            # CAN USE THIS BLOCK TO FORCE UPPER SERVICE_CONNECTIONS DRIVERS FALSE BEFORE USE HERE
            #$UseOP=$false ; 
            #$UseExOP=$false ;
            #$useForestWide = $false ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
            #$UseOPAD = $false ; 
            # ---
            $UseOPDYN=$true ; 
            $UseExOPDYN=$true ;
            $UseEXOPInSite=$true ;
            $useForestWideDYN = $false ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
            $UseOPADDYN = $true ; 
            $UseOPDYN = [boolean]($UseOPDYN -OR $UseExOPDYN -OR $UseOPADDYN) ;
            if($UseOPDYN -AND $UseOP){ write-warning "BOTH `$UseOPDYN -AND `$UseOP ARE TRUE!`nUSE ONE OR THE OTHER!" } ;
            if($UseExOPDYN -AND $UseExOP){ write-warning "BOTH `$UseExOPDYN -AND `$UseExOP ARE TRUE!`nUSE ONE OR THE OTHER!" } ;
            if($UseOPADDYN -AND $UseOPAD){ write-warning "BOTH `$UseOPADDYN -AND `$UseOPAD ARE TRUE!`nUSE ONE OR THE OTHER!" } ;
            if($useForestWideDYN -AND $useForestWide){ write-warning "BOTH `$useForestWideDYN -AND `$useForestWide ARE TRUE!`nUSE ONE OR THE OTHER!" } ;
            #*------^ END STEERING VARIS ^------
            #region GENERIC_EXOP_SRVR_CONN #*------v GENERIC_EXOP_SRVR_CONN BP v------
            if($UseOPDYN){
                #*------v GENERIC EXOP SRVR CONN BP v------
                # connect to ExOP 
                if($UseExOPDYN){
                    'get-ADExchangeSErverTDO','Connect-ExchangeServerTDO'  | foreach-object{
                        if(get-command $_ -ErrorAction STOP){} else { 
                            $smsg = "MISSING DEP FUNC:$($_)! must be either local include, or pre-loaded for EXOP connectivity to work!" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            throw $smsg ; 
                            break ; 
                        } ; 
                    } ; 
                    if($UseEXOPInSite){
                        TRY{
                            $Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name 
                        }CATCH{$Site=$env:COMPUTERNAME} ;
                        #$HubServers = get-ADExchangeSErverTDO -RoleNames 'HUB' -verbose
                        $PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
                    }else{
                        $PSSession = Connect-ExchangeServerTDO -RoleNames @('HUB','CAS') -verbose ; 
                    } ; 
                    # from get-ADExchangeSErverTDO return 
                    #if([double]$ExVersNum = [regex]::match($ExOPServer.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
                    # from Connect-ExchangeServerTDO pssession return
                    if([double]$ExVersNum = [regex]::match($PsSession.applicationprivatedata.supportedversions,'(\d+\.\d+)\.(\d+\.\d+)').groups[1].value){
                        switch -regex ([string]$ExVersNum) {
                            '15.2' { $isEx2019 = $true ; $ExVers = 'Ex2019' }
                            '15.1' { $isEx2016 = $true ; $ExVers = 'Ex2016'}
                            '15.0' { $isEx2013 = $true ; $ExVers = 'Ex2013'}
                            '14.*' { $isEx2010 = $true ; $ExVers = 'Ex2010'}
                            '8.*' { $isEx2007 = $true ; $ExVers = 'Ex2007'}
                            '6.5' { $isEx2003 = $true ; $ExVers = 'Ex2003'}
                            '6' {$isEx2000 = $true ; $ExVers = 'Ex2000'} ;
                            default {
                                $smsg = "UNRECOGNIZED ExVersNum.Major.Minor string:$($ExVersNum)! ABORTING!" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                THROW $SMSG ;
                                BREAK ;
                            }
                        } ;
                        $smsg = "`$ExVersNum: $($PsSession.applicationprivatedata.supportedversions)`n$((gv isex*| %{"`n`$$($_.name): `$$($_.value)"}|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    }else {
                        #$smsg = "UNABLE TO RESOLVE `$ExVersNum from `$ExOPServer.version:$($ExOPServer.version)!" ;
                        $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$PsSession.applicationprivatedata.supportedversions:$($PsSession.applicationprivatedata.supportedversions)!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        throw $smsg ;
                        break ;
                    } ; 
                    TRY{
                        if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig' -ea SilentlyContinue){
                            if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                                $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ; 
                                throw $smsg ; 
                                $smsg | write-warning  ; 
                            }else{
                                $smsg = "Connected to Orgname: $($OrgName)" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            } ; 
                        }else{
                            $smsg = "Missing 'tmp_*' module with 'get-OrganizationConfig'!" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ; 
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = $ErrTrapd ;
                        $smsg += "`n";
                        $smsg += $ErrTrapd.Exception.Message ;
                        if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        CONTINUE ;
                    } ;
                } ; 
                if($useForestWideDYN){
                    #region  ; #*------v USEFORESTWIDEDYN OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT v------
                    $smsg = "(`$useForestWideDYN:$($useForestWideDYN)):Enabling EXoP Forestwide)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Set-AdServerSettings -ViewEntireForest $True ;
                    #endregion  ; #*------^ END USEFORESTWIDEDYN OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT ^------
                } ;
            } else {
                $smsg = "(`$UseOPDYN:$($UseOPDYN))" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            }  ;  # if-E $UseOPDYN
            #endregion GENERIC_EXOP_SRVR_CONN #*------^ GENERIC_EXOP_SRVR_CONN BP ^------
            #region UseOPDYN #*------v UseOPDYN v------
            if($UseOPDYN -OR $UseOPADDYN){
                #region GENERIC_ADMS_CONN_&_XO #*------v GENERIC ADMS CONN & XO  v------
                if($UseOPADDYN){
                    'load-ADMs' | foreach-object{
                        if(get-command $_ -ErrorAction STOP){} else { 
                            $smsg = "MISSING DEP FUNC:$($_)! must be either local include, or pre-loaded for EXOP connectivity to work!" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            throw $smsg ; 
                            break ; 
                        } ; 
                    } ; 
                } ; 
                $smsg = "(loading ADMS...)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                # always capture load-adms return, it outputs a $true to pipeline on success
                $ADMTLoaded = load-ADMS -Verbose:$FALSE ;
                # 9:32 AM 4/20/2023 trimmed disabled/fw-borked cross-org code
                TRY {
                    if(-not(Get-ADDomain  -ea STOP).DNSRoot){
                        $smsg = "Missing AD Connection! (no (Get-ADDomain).DNSRoot returned)" ; 
                        throw $smsg ; 
                        $smsg | write-warning  ; 
                    } ; 
                    $objforest = get-adforest -ea STOP ; 
                    # Default new UPNSuffix to the UPNSuffix that matches last 2 elements of the forestname.
                    if($forestdom = $UPNSuffixDefault = $objforest.UPNSuffixes | ?{$_ -eq (($objforest.name.split('.'))[-2..-1] -join '.')}){

                    } elseif( $objforest.RootDomain -eq 'cmw.internal'){
                        # cmw doesn't use cmw.internal (forestname), as a UPN suffix, to try to match
                        # they have no default, tho' if we build up in the route, they'd have charlesmachineworks.com
                        $forestdom = $objforest.RootDomain; 
                        $UPNSuffixDefault = 'charlesmachine.works'
                    } else {
                         $smsg = "Unsupported `$objforest.RootDomain ($objforest.RootDomain), aborting!" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        throw $smsg ; 
                        break ; 
                    }
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = $ErrTrapd ;
                    $smsg += "`n";
                    $smsg += $ErrTrapd.Exception.Message ;
                    if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    CONTINUE ;
                } ;        
                #endregion GENERIC_ADMS_CONN_&_XO #*------^ END GENERIC ADMS CONN & XO ^------
            } else {
                $smsg = "(`$UseOPDYN:$($UseOPDYN))" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            }  ;
            #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
            #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller = get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
            # use new get-GCFastXO cross-org dc finde
            # default to Op_ExADRoot forest from $TenOrg Meta
            #if($UseOPDYN -AND -not $domaincontroller){
            if($UseOPDYN -AND -not (get-variable domaincontroller -ea 0)){
                #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((get-variable -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};
                # need to debug the above, credential issue?
                # just get it done
                #$domaincontroller = get-GCFast
                # AD – Return one GC in local site: (uses ADMS) Completely dynamic, no installed verb-xxx dependancies, 
                $domaincontroller = (Get-ADDomainController -Filter  {isGlobalCatalog -eq $true -AND Site -eq "$((get-adreplicationsite).name)"}).name| Get-Random ; 
            }  else { 
                # have to defer to get-azuread, or use EXO's native cmds to poll grp members
                # TODO 1/15/2021
                $useEXOforGroups = $true ; 
                $smsg = "$($TenOrg):HAS NO ON-PREM ACTIVEDIRECTORY, DEFERRING ALL GROUP ACCESS & MGMT TO NATIVE EXO CMDS!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
            if($useForestWideDYN -AND -not $GcFwide){
                #region GCFWIDE ; #*------v GCFWIDE OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT: v------
                $smsg = "`$GcFwide = Get-ADDomainController -Discover -Service GlobalCatalog" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;
                $smsg = "Discovered `$GcFwide:$($GcFwide)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #endregion GCFWIDE ; #*------^ END GCFWIDE OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT ^------
            } ;
            #endregion UseOPDYN #*------^ END UseOPDYN ^------
            #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
            #endregion SERVICE_CONNECTIONS_DEPENDANCYLESS #*======^ END SERVICE_CONNECTIONS_DEPENDANCYLESS ^======

        } ;  # if-E $isBased
    } ;  # PROC-E
    END {
        <# borked by psreadline v1/v2 breaking changes
        if(($PSFgColor = (Get-Variable  -name "$($TenOrg)Meta").value.PSFgColor) -AND ($PSBgColor = (Get-Variable  -name "$($TenOrg)Meta").value.PSBgColor)){
            write-verbose "(setting console colors:$($TenOrg)Meta.PSFgColor:$($PSFgColor),PSBgColor:$($PSBgColor))" ;
            $Host.UI.RawUI.BackgroundColor = $PSBgColor
            $Host.UI.RawUI.ForegroundColor = $PSFgColor ;
        } ;
        #>
    }
}

#*------^ Connect-Ex2010.ps1 ^------


#*------v Connect-Ex2010XO.ps1 v------
Function Connect-Ex2010XO {
    <#
    .SYNOPSIS
    Connect-Ex2010XO - Establish PSS to Ex2010, with multi-org support & validation
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-10-15
    FileName    :
    License     :
    Copyright   :
    Github      : https://github.com/tostka
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    AddedCredit : ExactMike Perficient
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    AddedTwitter:
    AddedCredit2 : Jeremy Bradshaw
    AddedWebsite2:	https://github.com/JeremyTBradshaw
    AddedTwitter2:
    REVISIONS   :
    * 1:30 PM 9/5/2024 added  update-SecurityProtocolTDO() SB to begin
    # 3:18 PM 5/18/2021 somehow lost $credOpTORSID, so flipped lost default $credOPTor -> $credTORSID
    # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods; flipped import-psess & import-mod to splats (cleaner) ; line-wrapped longer post-filters for legib
    * 8:30 AM 10/22/2020 ren'd $TentantTag -> $TenOrg, swapped looping meta resolve with 1-liner approach ; added AcceptedDom caching to the middle status test (suppress one more get-exoaccepteddomain call if possible), replaced all $Meta.value with the $TenOrg version
    * 12:56 PM 10/15/2020 converted connect-exo to Ex2010, adding onprem validation
    .DESCRIPTION
    Connect-Ex2010XO - Establish PSS to Ex2010, with multi-org support & validation
    .PARAMETER  ExchangeServer
    On Prem Exch server to Remote to
    .PARAMETER  ExAdmin
    Use exadmin IIS WebPool for remote EMS[-ExAdmin]
    .PARAMETER  Credential
    Credential object
    .PARAMETER  showDebug
    Debugging Flag [-showDebug]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Connect-Ex2010XO
    Connect using defaults, and leverage any pre-set $global:credo365TORSID variable
    .EXAMPLE
    connect-exo -CommandPrefix exo -credential (Get-Credential -credential s-todd.kadrie@torolab.com)  ;
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .EXAMPLE
    $cred = get-credential -credential $o365_Torolab_SIDUpn ;
    connect-exo -credential $cred ;
    Pass in a prefab credential object (useful for auto-shifting to MFA - the function will autoresolve MFA reqs based on the cred domain)
    .LINK
    .LINK
    https://github.com/verb-Exch2010
    #>
    [CmdletBinding()]
    [Alias('cxoxo')]
    Param(
        [Parameter(Position = 0, HelpMessage = "Exch server to Remote to")]
        [string]$ExchangeServer,
        [Parameter(HelpMessage = 'Use variant IIS WebPool for remote EMS[-ExAdmin]')]
        $ExAdmin,
        [Parameter(HelpMessage = 'Credential object')]
        [System.Management.Automation.PSCredential]$Credential = $credTORSID,
        [Parameter(HelpMessage = "Debugging Flag [-showDebug]")]
        [switch] $showDebug
    ) ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ;
        $CurrentVersionTlsLabel = [Net.ServicePointManager]::SecurityProtocol ; # Tls, Tls11, Tls12 ('Tls' == TLS1.0)  ;
        write-verbose "PRE: `$CurrentVersionTlsLabel : $($CurrentVersionTlsLabel )" ;
        # psv6+ already covers, test via the SslProtocol parameter presense
        if ('SslProtocol' -notin (Get-Command Invoke-RestMethod).Parameters.Keys) {
            $currentMaxTlsValue = [Math]::Max([Net.ServicePointManager]::SecurityProtocol.value__,[Net.SecurityProtocolType]::Tls.value__) ;
            write-verbose "`$currentMaxTlsValue : $($currentMaxTlsValue )" ;
            $newerTlsTypeEnums = [enum]::GetValues('Net.SecurityProtocolType') | Where-Object { $_ -gt $currentMaxTlsValue }
            if($newerTlsTypeEnums){
                write-verbose "Appending upgraded/missing TLS `$enums:`n$(($newerTlsTypeEnums -join ','|out-string).trim())" ;
            } else {
                write-verbose "Current TLS `$enums are up to date with max rev available on this machine" ;
            };
            $newerTlsTypeEnums | ForEach-Object {
                [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $_
            } ;
        } ;
        #if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        # $rgxEx10HostName : ^(lyn|bcc|adl|spb)ms6[4,5][0,1].global.ad.toro.com$
        # we'd need to define all possible hostnames to cover potential span. Should probably build dynamically from $XXXMeta vari
        # can build from $TorMeta.OP_ExADRoot:global.ad.toro.com
        <# on curly, from Ps into EMS:
        get-pssession | fl computername,computertype,state,configurationname,availability,name
        ComputerName      : curlyhoward.cmw.internal
        ComputerType      : RemoteMachine
        State             : Opened
        ConfigurationName : Microsoft.Exchange
        Availability      : Available
        Name              : Session1

        ComputerName      : lynms650.global.ad.toro.com
        ComputerType      : RemoteMachine
        State             : Broken
        ConfigurationName : Microsoft.Exchange
        Availability      : None
        Name              : Exchange2010

        "^\w*\.$($CMWMeta.OP_ExADRoot)$"
        => ^\w*\.cmw.internal$
        #>

        #$sTitleBarTag = "EMS" ;
        $CommandPrefix = $null ;

        $TenOrg=get-TenantTag -Credential $Credential ;
        <#if($TenOrg -ne 'TOR'){
            # explicitly leave this tenant (default) untagged
            $sTitleBarTag += $TenOrg ;
        } ;
        #>
        if($TenOrg){
            switch -regex ($TenOrg){
                '^(CMW|TOR)$'{
                    $sTitleBarTag = @("EMS$($TenOrg.substring(0,1).tolower())") ; # 1st char
                }
                '^TOL$'{
                    $sTitleBarTag = @("EMS$($TenOrg.substring(2,1).tolower())") ; # last char
                } ;
                default{
                    throw "$($TenOrg):unsupported `$TenOrg!" ;
                    break ;
                }
            } ;
        } else {
            $sTitleBarTag = @("EMS") ;
        } ;
        write-verbose "`$sTitleBarTag:$($sTitleBarTag)" ; 

        <#
        $credDom = ($Credential.username.split("\"))[0] ;
        $Metas=(get-variable *meta|Where-Object{$_.name -match '^\w{3}Meta$'}) ;
        foreach ($Meta in $Metas){
            if( ($credDom -eq $Meta.value.legacyDomain) -OR ($credDom -eq $Meta.value.o365_TenantDomain) -OR ($credDom -eq $Meta.value.o365_OPDomain)){
                if($Meta.value.OP_ExADRoot){
                    if(!$Meta.value.OP_rgxEMSComputerName){
                        write-verbose "(adding XXXMeta.OP_rgxEMSComputerName value)"
                        # build vari that will match curlyhoward.cmw.internal|lynms650.global.ad.toro.com etc
                        set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'OP_rgxEMSComputerName' = "^\w*\.$([Regex]::Escape($Meta.value.OP_ExADRoot))$"} ) ;
                    } ;
                } else {
                    throw "Missing `$$($Meta.value.o365_Prefix).OP_ExADRoot value.`nProfile hasn't loaded proper tor-incl-infrastrings file)!"
                } ;
            } ; # if-E $credDom
        } ; # loop-E
        #>
        # non-looping vers:
        #$TenOrg = get-TenantTag -Credential $Credential ;
        #.OP_ExADRoot
        if( (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName ){

        } else {
            #.OP_rgxEMSComputerName
            if((Get-Variable  -name "$($TenOrg)Meta").value.OP_ExADRoot){
                set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'OP_rgxEMSComputerName' = "^\w*\.$([Regex]::Escape((Get-Variable  -name "$($TenOrg)Meta").value.OP_ExADRoot))$"} )
            } else {
                $smsg = "Missing `$$((Get-Variable  -name "$($TenOrg)Meta").value.o365_Prefix).OP_ExADRoot value.`nProfile hasn't loaded proper tor-incl-infrastrings file)!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;
        } ;

    } ;  # BEG-E
    PROCESS{
        # if we're using ems-style BasicAuth, clear incompatible existing Rems PSS's
        # ComputerName      : curlyhoward.cmw.internal ;  ComputerType      : RemoteMachine ;  State             : Opened ;  ConfigurationName : Microsoft.Exchange ;  Availability      : Available ;  Name              : Session1 ;   ;
        $rgxRemsPSSName = "^(Session\d|Exchange\d{4})$" ;
        $Rems2Good = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') } ;
        # Computername wrong fqdn suffix
        #$Rems2WrongOrg = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND (-not($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName)) -AND ($_.Availability -eq 'Available') } ;
        # above is seeing outlook EXO conns as wrong org, exempt them too: .ComputerName -match $rgxExoPsHostName
        $Rems2WrongOrg = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND (
            $_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND (
            ( -not($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) ) -AND (
            -not($_.ComputerName -match $rgxExoPsHostName)) ) -AND ($_.Availability -eq 'Available')
        } ;
        $Rems2Broken = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND (
            $_.Name -match $rgxRemsPSSName) -AND ($_.State -like "*Broken*") } ;
        $Rems2Closed = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND (
            $_.Name -match $rgxRemsPSSName) -AND ($_.State -like "*Closed*") } ;

        if ($Rems2Broken.count -gt 0){ for ($index = 0 ;$index -lt $Rems2Broken.count ;$index++){Remove-PSSession -session $Rems2Broken[$index]}  };
        if ($Rems2Closed.count -gt 0){for ($index = 0 ;$index -lt $Rems2Closed.count ; $index++){Remove-PSSession -session $Rems2Closed[$index] } } ;
        if ($Rems2WrongOrg.count -gt 0){for ($index = 0 ;$index -lt $Rems2WrongOrg.count ; $index++){Remove-PSSession -session $Rems2WrongOrg[$index] } } ;
        # preclear until proven *up*
        $bExistingREms = $false ;

        if( Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') } ){
            $bExistingREms= $true ;

        } ;
        if($bExistingREms -eq $false){
            #$TorMeta.Ex10Server: dynamic
            #$TorMeta.Ex10ServerXO: lynms650.global.ad.toro.com
            # force unresolved to dyn
            if((Get-Variable  -name "$($TenOrg)Meta").value.Ex10ServerXO){
                write-host -foregroundcolor darkgray "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Adding EMS (connecting to $($TorMeta.Ex10ServerXO))..." ;
            } ;

            $pltNSess = @{
                ConnectionURI = "http://$((Get-Variable  -name "$($TenOrg)Meta").value.Ex10ServerXO)/powershell";
                ConfigurationName = 'Microsoft.Exchange' ;
                name = 'Exchange2010' ;
            } ;
            if ((Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant) {
              # use variant IIS Webpool
              $pltNSess.ConnectionURI = $pltNSess.ConnectionURI.replace("/powershell", "/$((Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant)") ;
            }
            $pltNSess.Add("Credential", $Credential); # just use the passed $Credential vari
            $cMsg = "Connecting to OP Ex20XX ($($credDom))";
            Write-Host $cMsg ;
            write-verbose "`n$((get-date).ToString('HH:mm:ss')):New-PSSession w`n$(($pltNSess|out-string).trim())" ;

            $error.clear() ;
            TRY { $global:E10Sess = New-PSSession @pltNSess -ea STOP
            } CATCH {
                $ErrTrapd = $_ ;
                write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ((Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant) {
                  # switch to stock pool and retry
                  $pltNSess.ConnectionURI = $pltNSess.ConnectionURI.replace("/$((Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant)", "/powershell") ;
                  write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TARGETING VARIANT POOL`nRETRY W STOCK POOL: New-PSSession w`n$(($pltNSess|out-string).trim())" ;
                  $global:E10Sess = New-PSSession @pltNSess -ea STOP  ;
                } else {
                    STOP ;
                } ;
            } ; # try-E

            if(!$global:E10Sess){
                write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO RETURN PSSESSION!`nAUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                Break ;
            } ;

            $pltIMod=@{Global = $true ;PassThru = $true;DisableNameChecking = $true ; verbose=$true ;} ;
            $pltISess = [ordered]@{
                Session             = $global:E10Sess ;
                DisableNameChecking = $true  ;
                AllowClobber        = $true ;
                ErrorAction         = 'Stop' ;
                Verbose             = $false ;
            } ;
            if ($CommandPrefix) {
                write-host -foregroundcolor white "$((get-date).ToString("HH:mm:ss")):Note: Prefixing this Mod's Cmdlets as [verb]-$($CommandPrefix)[noun]" ;
                $pltIMod.add('Prefix',$CommandPrefix) ;
                $pltISess.add('Prefix',$CommandPrefix) ;
            } ;
            write-verbose "`n$((get-date).ToString('HH:mm:ss')):Import-PSSession w`n$(($pltISess|out-string).trim())`nImport-Module w`n$(($pltIMod|out-string).trim())" ;

            # Verbose:Continue is VERY noisey for module loads. Bracketed suppress:
            if($VerbosePreference = "Continue"){
                $VerbosePrefPrior = $VerbosePreference ;
                $VerbosePreference = "SilentlyContinue" ;
                $verbose = ($VerbosePreference -eq "Continue") ;
            } ;
            Try {
                $Global:E10Mod = Import-Module (Import-PSSession @pltISess) @pltIMod  ;
                #$Global:EOLModule = Import-Module (Import-PSSession @pltISess) -Global -Prefix $CommandPrefix -PassThru -DisableNameChecking   ;
            } catch {
                Write-Warning -Message "Tried but failed to import the EXO PS module.`n`nError message:" ;
                throw $_ ;
            } ;
            # reenable VerbosePreference:Continue, if set, during mod loads
            if($VerbosePrefPrior -eq "Continue"){
                $VerbosePreference = $VerbosePrefPrior ;
                $verbose = ($VerbosePreference -eq "Continue") ;
            } ;
            Add-PSTitleBar $sTitleBarTag -verbose:$($VerbosePreference -eq "Continue")  ;

        } ; #  # if-E $bExistingREms
    } ;  # PROC-E
    END {
        if($bExistingREms -eq $false){
            if( Get-PSSession | where-object {$_.ConfigurationName -eq "Microsoft.Exchange" -AND $_.Name -match $rgxRemsPSSName -AND $_.State -eq "Opened" -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') }  ){
                $bExistingREms= $true ;
            } else {
                write-error "(Credential mismatch:disconnecting from existing Ex201X:$($eEXO.Identity) tenant)" ;
                Disconnect-Ex2010 ;
                $bExistingREms = $false ;
            } ;
        } ;
    } ; # END-E
}

#*------^ Connect-Ex2010XO.ps1 ^------


#*------v Connect-ExchangeServerTDO.ps1 v------
if(-not(get-command Connect-ExchangeServerTDO -ea 0)){
    Function Connect-ExchangeServerTDO {
        <#
        .SYNOPSIS
        Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
        will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellREmote (REMS) connect to each server, 
        stopping at the first successful connection.
        .NOTES
        Version     : 3.0.3
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2024-05-30
        FileName    : Connect-ExchangeServerTDO.ps1
        License     : (none-asserted)
        Copyright   : (none-asserted)
        Github      : https://github.com/tostka/verb-Ex2010
        Tags        : Powershell, ActiveDirectory, Exchange, Discovery
        AddedCredit : Brian Farnsworth
        AddedWebsite: https://codeandkeep.com/
        AddedTwitter: URL
        AddedCredit : David Paulson
        AddedWebsite: https://techcommunity.microsoft.com/t5/exchange-team-blog/exchange-health-checker-has-a-new-home/ba-p/2306671
        AddedTwitter: URL
        REVISIONS
        * 3:54 PM 11/26.2.34 integrated back TLS fixes, and ExVersNum flip from June; syncd dbg & vx10 copies.
        * 12:57 PM 6/11/2024 Validated, Ex2010 & Ex2019, hub, mail & edge roles: tested ☑️ on CMW mail role (Curly); and Jumpbox; 
            copied in CBH from repo copy, which has been updated/debugged compat on CMW Edge 
            includes local snapin detect & load for edge role (simplest EMS load option for Edge role, from David Paulson's original code; no longer published with Ex2010 compat)
        * 1:30 PM 9/5/2024 added  update-SecurityProtocolTDO() SB to begin
        * 12:49 PM 6/21/2024 flipped PSS Name to Exchange$($ExchVers[dd])
        * 11:28 AM 5/30/2024 fixed failure to recognize existing functional PSSession; Made substantial update in logic, validate works fine with other orgs, and in our local orgs.
        * 4:02 PM 8/28/2023 debuged, updated CBH, renamed connect-ExchangeSErver -> Connect-ExchangeServerTDO (avoid name clashes, pretty common verb-noun combo).
        * 12:36 PM 8/24/2023 init

        .DESCRIPTION
        Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
        will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellRemote (REMS) connect to each server, 
        stopping at the first successful connection.

        Relies upon/requires get-ADExchangeServerTDO(), to return a descriptive summary of the Exchange server(s) revision etc, for connectivity logic.
        Supports Exchange 2010 through 2019, as implemented.
        
        Intent, as contrasted with verb-EXOP/Ex2010 is to have no local module dependancies, when running EXOP into other connected orgs, where syncing profile & supporting modules code can be problematic. 
        This uses native ADSI calls, which are supported by Windows itself, without need for external ActiveDirectory module etc.

        The particular approach inspired by BF's demo func that accompanied his take on get-adExchangeServer(), which I hybrided with my own existing code for cred-less connectivity. 
        I added get-OrganizationConfig testing, for connection pre/post confirmation, along with Exchange Server revision code for continutional handling of new-pssession remote powershell EMS connections.
        Also shifted connection code into _connect-EXOP() internal func.
        As this doesn't rely on local module presence, it doesn't have to do the usual local remote/local invocation detection you'd do for non-dehydrated on-server EMS (more consistent this way, anyway; 
        there are only a few cmdlet outputs I'm aware of, that have fundementally broken returns dehydrated, and require local non-remote EMS use to function.

        My core usage would be to paste the function into the BEGIN{} block for a given remote org process, to function as a stricly local ad-hoc function.
        .PARAMETER name
        FQDN of a specific Exchange server[-Name EXSERVER.DOMAIN.COM]
        .PARAMETER discover
        Boolean paraameter that drives auto-discovery of target Exchange servers for connection (defaults `$true)[-discover:`$false]
        .PARAMETER credential
        Use specific Credentials[-Credentials [credential object]
            .PARAMETER Site
        Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-Site 'SITENAME']
        .PARAMETER RoleNames
        Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']
        .PARAMETER TenOrg
        Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
        .INPUTS
        None. Does not accepted piped input.(.NET types, can add description)
        .OUTPUTS
        [system.object] Returns a system object containing a successful PSSession
        System.Boolean
        [| get-member the output to see what .NET obj TypeName is returned, to use here]
        System.Array of System.Object's
        .EXAMPLE
        PS> $PSSession = Connect-ExchangeServerTDO -siteName SITENAME -RoleNames @('HUB','CAS') -verbose 
        Demo's connecting to a functional Hub or CAS server in the SITENAME site with verbose outputs, the `PSSession variable will contain information about the successful connection. Makes automatic Exchangeserver discovery calls into AD (using ADSI) leveraging the separate get-ADExchangeServerTDO()
        .EXAMPLE
        PS> TRY{$Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name}CATCH{$Site=$env:COMPUTERNAME} ;
        PS> $PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
        Demo including support for EdgeRole, which is detected on it's lack of AD Site specification (which gets fed through to call, by setting the Site to the machine itself).
        .LINK
        https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/
        .LINK
        https://github.com/Lucifer1993/PLtools/blob/main/HealthChecker.ps1
        .LINK
        https://microsoft.github.io/CSS-Exchange/Diagnostics/HealthChecker/
        .LINK
        https://bitbucket.org/tostka/powershell/
        .LINK
        https://github.com/tostka/verb-Ex2010
        #>        
        [CmdletBinding(DefaultParameterSetName='discover')]
        PARAM(
            [Parameter(Position=0,Mandatory=$true,ParameterSetName='name',HelpMessage="FQDN of a specific Exchange server[-Name EXSERVER.DOMAIN.COM]")]
                [String]$name,
            [Parameter(Position=0,ParameterSetName='discover',HelpMessage="Boolean paraameter that drives auto-discovery of target Exchange servers for connection (defaults `$true)[-discover:`$false]")]
                [bool]$discover=$true,
            [Parameter(Position=1,HelpMessage = "Use specific Credentials[-Credentials [credential object]")]
                [Management.Automation.PSCredential]$credential,
            [Parameter(Position=1,HelpMessage="Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-Site 'SITENAME']")]
                [Alias('Site')]
                [string]$SiteName = (gv -name "$($TenOrg)_ADSiteDefaultName" -ea 0).Value,
            [Parameter(Position=2,HelpMessage="Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']")]
                [ValidateSet('MBX','CAS','HUB','UM','MBX','EDGE')]
                [string[]]$RoleNames = @('HUB','CAS'),
            [Parameter(Mandatory=$FALSE,HelpMessage="Tenant Tag (3-letter abbrebiation - defaults to global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']")]
                [ValidateNotNullOrEmpty()]
                [string]$TenOrg = $global:o365_TenOrgDefault
        ) ;
        BEGIN{
            $Verbose = ($VerbosePreference -eq 'Continue') ;
            $CurrentVersionTlsLabel = [Net.ServicePointManager]::SecurityProtocol ; # Tls, Tls11, Tls12 ('Tls' == TLS1.0)  ;
			write-verbose "PRE: `$CurrentVersionTlsLabel : $($CurrentVersionTlsLabel )" ;
			# psv6+ already covers, test via the SslProtocol parameter presense
			if ('SslProtocol' -notin (Get-Command Invoke-RestMethod).Parameters.Keys) {
				$currentMaxTlsValue = [Math]::Max([Net.ServicePointManager]::SecurityProtocol.value__,[Net.SecurityProtocolType]::Tls.value__) ;
				write-verbose "`$currentMaxTlsValue : $($currentMaxTlsValue )" ;
				$newerTlsTypeEnums = [enum]::GetValues('Net.SecurityProtocolType') | Where-Object { $_ -gt $currentMaxTlsValue }
				if($newerTlsTypeEnums){
					write-verbose "Appending upgraded/missing TLS `$enums:`n$(($newerTlsTypeEnums -join ','|out-string).trim())" ;
				} else {
					write-verbose "Current TLS `$enums are up to date with max rev available on this machine" ;
				};
				$newerTlsTypeEnums | ForEach-Object {
					[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $_
				} ;
			} ;
            $smsg = "#*------v Function _connect-ExOP v------" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            function _connect-ExOP{
                [CmdletBinding()]
                PARAM(
                    [Parameter(Position=0,Mandatory=$true,HelpMessage="Exchange server AD Summary system object[-Server EXSERVER.DOMAIN.COM]")]
                        [system.object]$Server,
                    [Parameter(Position=1,HelpMessage = "Use specific Credentials[-Credentials [credential object]")]
                        [Management.Automation.PSCredential]$credential
                );
                $verbose = $($VerbosePreference -eq "Continue") ;
                if([double]$ExVersNum = [regex]::match($Server.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
                    switch -regex ([string]$ExVersNum) {
                        '15.2' { $isEx2019 = $true ; $ExVers = 'Ex2019' }
                        '15.1' { $isEx2016 = $true ; $ExVers = 'Ex2016'}
                        '15.0' { $isEx2013 = $true ; $ExVers = 'Ex2013'}
                        '14.*' { $isEx2010 = $true ; $ExVers = 'Ex2010'}
                        '8.*' { $isEx2007 = $true ; $ExVers = 'Ex2007'}
                        '6.5' { $isEx2003 = $true ; $ExVers = 'Ex2003'}
                        '6' {$isEx2000 = $true ; $ExVers = 'Ex2000'} ;
                        default {
                            $smsg = "UNRECOGNIZED ExVersNum.Major.Minor string:$($ExVersNum)! ABORTING!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            THROW $SMSG ;
                            BREAK ;
                        }
                    } ;
                }else {
                    $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$Server.version:$($Server.version)!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                    break ;
                } ;
                if($Server.RoleNames -eq 'EDGE'){
                    if(($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup')) -or
                        ($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup')) -or
                        $ByPassLocalExchangeServerTest)
                    {
                        if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole') -or
                                (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'))
                        {
                            $smsg = "We are on Exchange Edge Transport Server"
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            $IsEdgeTransport = $true
                        }
                        TRY {
                            Get-ExchangeServer -ErrorAction Stop | Out-Null
                            $smsg = "Exchange PowerShell Module already loaded."
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            $passed = $true 
                        }CATCH {
                            $smsg = "Failed to run Get-ExchangeServer"
                            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            if($isLocalExchangeServer){
                                write-host  "Loading Exchange PowerShell Module..."
                                TRY{
                                    if($IsEdgeTransport){
                                        # implement local snapins access on edge role: Only way to get access to EMS commands.
                                        [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exshell.psc1" -ErrorAction Stop
                                        ForEach($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn){
                                            write-verbose ("Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name)
                                            Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop
                                        } ; 
                                        Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop ; 
                                        $passed = $true #We are just going to assume this passed.
                                    }else{
                                        Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                                        Connect-ExchangeServer -Auto -ClientApplication:ManagementShell
                                        $passed = $true #We are just going to assume this passed.
                                    } 
                                }CATCH {
                                    $smsg = "Failed to Load Exchange PowerShell Module..." ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                }                               
                            } ;
                        } FINALLY {
                            if($LoadExchangeVariables -and $passed -and $isLocalExchangeServer){
                                if($ExInstall -eq $null -or $ExBin -eq $null){
                                    if(Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup'){
                                        $Global:ExInstall = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath
                                    }else{
                                        $Global:ExInstall = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath
                                    }
        
                                    $Global:ExBin = $Global:ExInstall + "\Bin"
        
                                    $smsg = ("Set ExInstall: {0}" -f $Global:ExInstall)
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                    $smsg = ("Set ExBin: {0}" -f $Global:ExBin)
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                } ; 
                            } ; 
                        } ; 
                    } else  {
                        $smsg = "Does not appear to be an Exchange 2010 or newer server." ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                    }
                    if(get-command -Name Get-OrganizationConfig -ea 0){
                        $smsg = "Running in connected/Native EMS" ; 
                        if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        Return $true ; 
                    } else { 
                        TRY{
                            $smsg = "Initiating Edge EMS local session (exshell.psc1 & exchange.ps1)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            # 5;36 PM 5/30/2024 didn't work, went off to nowhere for a long time, and exited the script
                            #& (gcm powershell.exe).path -PSConsoleFile "$($env:ExchangeInstallPath)bin\exshell.psc1" -noexit -command ". '$($env:ExchangeInstallPath)bin\Exchange.ps1'"
                            <# [Adding the Transport Server to Exchange - Mark Lewis Blog](https://marklewis.blog/2020/11/19/adding-the-transport-server-to-exchange/)
                            To access the management console on the transport server, I opened PowerShell then ran
                            exshell.psc1
                            Followed by
                            exchange.ps1
                            At this point, I was able to create a new subscription using he following PowerShel
                            #>
                            invoke-command exshell.psc1 ; 
                            invoke-command exchange.ps1
                            if(get-command -Name Get-OrganizationConfig -ea 0){
                                $smsg = "Running in connected/Native EMS" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                Return $true ;
                            } else { return $false };  
                        }CATCH{
                            Write-Error $_ ;
                        } ;
                    } ; 
                } else {
                    $pltNPSS=@{ConnectionURI="http://$($Server.FQDN)/powershell"; ConfigurationName='Microsoft.Exchange' ; name="Exchange$($ExVersNum.tostring())"} ;
                    # use ExVersUnm dd instead of hardcoded (Exchange2010)
                    if($ExVersNum -ge 15){
                        $smsg = "EXOP.15+:Adding -Authentication Kerberos" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $pltNPSS.add('Authentication',"Kerberos") ;
                        $pltNPSS.name = $ExVers ;
                    } ;
                    $smsg = "Adding EMS (connecting to $($Server.FQDN))..." ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    $smsg = "New-PSSession w`n$(($pltNPSS|out-string).trim())" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    $ExPSS = New-PSSession @pltNPSS  ;
                    $ExIPSS = Import-PSSession $ExPSS -allowclobber ;
                    $ExPSS | write-output ;
                    $ExPSS= $ExIPSS = $null ;
                } ; 
            } ;
            $smsg = "#*------^ END Function _connect-ExOP ^------" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            $pltGADX=@{
                ErrorAction='Stop';
            } ;
        } ;
        PROCESS{
            if($PSBoundParameters.ContainsKey('credential')){
                $pltGADX.Add('credential',$credential) ;
            }
            if($SiteName){
                $pltGADX.Add('siteName',$siteName) ;
            } ;
            if($RoleNames){
                $pltGADX.Add('RoleNames',$RoleNames) ;
            } ;
            TRY{
                if($discover){
                    $smsg = "Getting list of Exchange Servers" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    $exchServers=get-ADExchangeServerTDO @pltGADX | sort responsetime ;
                }else{
                    $exchServers=get-ADExchangeServerTDO @pltGADX | sort responsetime ;
                } ;
                $pltTW=@{
                    'ErrorAction'='Stop';
                } ;
                $pltCXOP = @{
                    verbose = $($VerbosePreference -eq "Continue") ;
                } ;
                if($pltGADX.credential){
                    $pltCXOP.Add('Credential',$pltCXOP.Credential) ;
                } ;
                $prpPSS = 'Id','Name','ComputerName','ComputerType','State','ConfigurationName','Availability' ; 
                foreach($exServer in $exchServers){
                    $smsg = "testing conn to:$($exServer.name.tostring())..." ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig' -ea SilentlyContinue){
                        if($pssEXOP = Get-PSSession |  where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ( $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND ( $_.runspace.ConnectionInfo.Port -eq '80') }){
                            if($pssEXOP.State -ne "Opened" -OR $pssEXOP.Availability -ne "Available"){
                                $pssEXOP | remove-pssession ; $pssEXOP = $null ;
                            } ;
                        } ; 
                    } else {
                        $smsg = "(mangled ExOP conn: disconnect/reconnect...)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        if($pssEXOP = Get-PSSession |  where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ( $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND ( $_.runspace.ConnectionInfo.Port -eq '80') }){
                            if($pssEXOP.State -ne "Opened" -OR $pssEXOP.Availability -ne "Available"){
                                $pssEXOP | remove-pssession ; $pssEXOP = $null ;
                            } ;
                        } ; 
                    } ;
                    if(-not $pssEXOP){
                        $smsg = "Connecting to: $($exServer.FQDN)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        if($NoTest){
                            $ExPSS =$ExPSS = _connect-ExOP @pltCXOP -Server $exServer
                        } else {
                            TRY{
                                $smsg = "Testing Connection: $($exServer.FQDN)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                If(test-connection $exServer.FQDN -count 1 -ea 0) {
                                    $smsg = "confirmed pingable..." ;
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                } else {
                                    $smsg = "Unable to Ping $($exServer.FQDN)" ; ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                } ;
                                $smsg = "Testing WinRm: $($exServer.FQDN)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                $winrm=Test-WSMan @pltTW -ComputerName $exServer.FQDN ;
                                if($winrm){
                                    $ExPSS = _connect-ExOP @pltCXOP -Server $exServer;
                                } else {
                                    $smsg = "Unable to Test-WSMan $($exServer.FQDN) (skipping)" ; ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                } ;
                            }CATCH{
                                $errMsg="Server: $($exServer.FQDN)] $($_.Exception.Message)" ;
                                Write-Error -Message $errMsg ;
                                continue ;
                            } ;
                        };
                    } else {
                        $smsg = "$((get-date).ToString('HH:mm:ss')):Accepting first valid connection w`n$(($pssEXOP | ft -a $prpPSS|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $ExPSS = $pssEXOP ; 
                        break ; 
                    }  ;
                } ;
            }CATCH{
                Write-Error $_ ;
            } ;
        } ;
        END{
            if(-not $ExPSS){
                $smsg = "NO SUCCESSFUL CONNECTION WAS MADE, WITH THE SPECIFIED INPUTS!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $smsg = "(returning `$false to the pipeline...)" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                return $false
            } else{
                if($ExPSS.State -eq "Opened" -AND $ExPSS.Availability -eq "Available"){
                    if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                        $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ;
                        throw $smsg ;
                        $smsg | write-warning  ;
                    } else {
                        $smsg = "(connected to EXOP.Org:$($orgName))" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    return $ExPSS
                } ;
            } ; 
        } ;
    } ;
}

#*------^ Connect-ExchangeServerTDO.ps1 ^------


#*------v cx10cmw.ps1 v------
function cx10cmw {
    <#
    .SYNOPSIS
    cx10cmw - Connect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Connect-EX2010 - Connect-EX2010 to specified on-prem Exchange
    .EXAMPLE
    cx10cmw
    #>
    [CmdletBinding()] 
    [Alias('cxOPcmw')]
    Param()
    $Verbose = ($VerbosePreference -eq 'Continue') ;
    $pltGHOpCred=@{TenOrg="CMW" ;userrole=@('ESVC','LSVC','SID') ;verbose=$($verbose)} ;
    if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
        Connect-EX2010 -cred $OPCred -Verbose:($VerbosePreference -eq 'Continue') ; 
    } else {
        $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole $($UserRole -join '|') value!"
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
        exit ;
    } ;
}

#*------^ cx10cmw.ps1 ^------


#*------v cx10tol.ps1 v------
function cx10tol {
    <#
    .SYNOPSIS
    cx10tol - Connect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Connect-EX2010 - Connect-EX2010 to specified on-prem Exchange
    .NOTES
    REVISIONS   :
    # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
    .EXAMPLE
    cx10tol
    #>
    [CmdletBinding()] 
    [Alias('cxOPtol')]
    Param()
    $Verbose = ($VerbosePreference -eq 'Continue') ;
    $pltGHOpCred=@{TenOrg="TOL" ;userrole=@('ESVC','LSVC','SID') ;verbose=$($verbose)} ;
    if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
        Connect-EX2010 -cred $OPCred #-Verbose:($VerbosePreference -eq 'Continue') ; 
    } else {
        $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole $($UserRole -join '|') value!"
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
        exit ;
    } ;
}

#*------^ cx10tol.ps1 ^------


#*------v cx10tor.ps1 v------
function cx10tor {
    <#
    .SYNOPSIS
    cx10tor - Connect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Connect-EX2010 - Connect-EX2010 to specified on-prem Exchange
    .NOTES
    REVISIONS   :
    # 3:18 PM 5/18/2021 somehow lost $credOpTORSID, so flipped lost default $credOPTor -> $credTORSID
    # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
    .EXAMPLE
    cx10tor
    #>
    [CmdletBinding()] 
    [Alias('cxOPtor')]
    Param([Parameter(HelpMessage = 'Credential object')][System.Management.Automation.PSCredential]$Credential = $credTorSID)
    $Verbose = ($VerbosePreference -eq 'Continue') ;
    if(!$Credential){
        $pltGHOpCred=@{TenOrg="TOR" ;userrole=@('SID','ESVC','LSVC') ;verbose=$($verbose)} ;
        if($Credential=(get-HybridOPCredentials @pltGHOpCred).cred){
            #Connect-EX2010 -cred $credTorSID #-Verbose:($VerbosePreference -eq 'Continue') ; 
        } else {
            $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole $($UserRole -join '|') value!"
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Break ;
        } ;
    } ; 
    Connect-EX2010 -cred $Credential #-Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ cx10tor.ps1 ^------


#*------v disable-ForestView.ps1 v------
Function disable-ForestView {
<#
.SYNOPSIS
disable-ForestView.ps1 - Disable Exchange onprem AD ViewEntireForest setting (permits org-wide object access, wo use of proper explicit -domaincontroller sub.domain.com)
.NOTES
Version     : 1.0.2
Author      : Todd Kadrie
Website     : http://www.toddomation.com
Twitter     : @tostka / http://twitter.com/tostka
CreatedDate : 2020-10-26
FileName    :
License     : MIT License
Copyright   : (c) 2020 Todd Kadrie
Github      : https://github.com/tostka/verb-XXX
Tags        : Powershell
REVISIONS
* 11:40 AM 5/14/2021 added -ea 0 to the gv tests (suppresses not-found error when called without logging config)
* 10:56 AM 4/2/2021 cleaned up; added recstat & wlt
* 11:44 AM 3/5/2021 variant of toggle-fv
.DESCRIPTION
disable-ForestView.ps1 - Disable Exchange onprem AD ViewEntireForest setting (permits org-wide object access, wo use of proper explicit -domaincontroller sub.domain.com)
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output
.EXAMPLE
disable-ForestView
.LINK
https://github.com/tostka/verb-ex2010
.LINK
#>
[CmdletBinding()]
PARAM() ;
    # toggle forest view
    if (get-command -name set-AdServerSettings){
        if ((get-AdServerSettings).ViewEntireForest ) {
              write-verbose "(set-AdServerSettings -ViewEntireForest `$False)"
              set-AdServerSettings -ViewEntireForest $False
        } ;
    } else {
        #-=-record a STATUSERROR=-=-=-=-=-=-=
        $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
        if(Get-Variable passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
        if(Get-Variable -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
        #-=-=-=-=-=-=-=-=
        $smsg = "MISSING:set-AdServerSettings`nOPEN an Exchange OnPrem connection FIRST!"
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        BREAK ;
    } ;
}

#*------^ disable-ForestView.ps1 ^------


#*------v Disconnect-Ex2010.ps1 v------
Function Disconnect-Ex2010 {
  <#
    .SYNOPSIS
    Disconnect-Ex2010 - Clear Remote Exch2010 Mgmt Shell connection
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    AddedCredit : Inspired by concept code by ExactMike Perficient, Global Knowl... (Partner)
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Version     : 1.1.0
    CreatedDate : 2020-02-24
    FileName    : Connect-Ex2010()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell,Exchange,ExchangeOnline
    REVISIONS   :
    * 3:11 PM 7/15/2024 needed to change CHKPREREQ to check for presence of prop, not that it had a value (which fails as $false); hadn't cleared $MetaProps = ...,'DOESNTEXIST' ; confirmed cxo working non-based
    * 10:47 AM 7/11/2024 cleared debugging NoSuch etc meta tests
    * 1:34 PM 6/21/2024 ren $Global:E10Sess -> $Global:EXOPSess ; add: prereq checks, and $isBased support, to devert into most connect-exchangeServerTDO, get-ADExchangeServerTDO 100% generic fall back support; sketched in Ex2013 disconnect support
    # 11:12 AM 10/25/2021 added trailing null $Global:E10Sess  (to avoid false conn detects on that test)
    # 9:44 AM 7/27/2021 add -PsTitleBar EMS[ctl] support by dyn gathering range of all 1st & last $Meta.Name[0,2] values
    # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
    * 1:14 PM 3/1/2021 added color reset
    * 4:13 PM 10/22/2020 added pretest of $Global:*'s before running at remove-module (suppresses errors)
    * 12:23 PM 5/27/2020 updated cbh, moved aliases:Disconnect-EMSR','dx10' win func
    * 10:51 AM 2/24/2020 updated attrib
    * 6:59 PM 1/15/2020 cleanup
    * 8:01 AM 11/1/2017 added Remove-PSTitlebar 'EMS', and Disconnect-PssBroken to the bottom - to halt growth of unrepaired broken connections. Updated example to pretest for reqMods
    * 12:54 PM 12/9/2016 cleaned up, add pshelp, implented and debugged as part of verb-Ex2010 set
    * 2:37 PM 12/6/2016 ported to local EMSRemote
    * 2/10/14 posted version
    .DESCRIPTION
    Disconnect-Ex2010 - Clear Remote Exch2010 Mgmt Shell connection
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    $reqMods="Remove-PSTitlebar".split(";") ;
    $reqMods | % {if( !(test-path function:$_ ) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing $($_) function. EXITING." } } ;
    Disconnect-Ex2010 ;
    .LINK
    #>
    [CmdletBinding()]
    [Alias('Disconnect-EMSR','dx10')]
    Param()
    BEGIN{
        
        #region CHKPREREQ ; #*------v CHKPREREQ v------
        # critical dependancy Meta variables
        $MetaNames = ,'TOR','CMW','TOL' #,'NOSUCH' ;
        # critical dependancy Meta variable properties
        $MetaProps = 'Ex10Server','Ex10WebPoolVariant','ExRevision','ExViewForest','ExOPAccessFromToro','legacyDomain' #,'DOESNTEXIST' ;
        # critical dependancy parameters
        $gvNames = 'Credential'
        $isBased = $true ; $gvMiss = @() ; $ppMiss = @() ;
        foreach($met in $metanames){
            write-verbose "chk:`$$($met)Meta" ;
            if(-not (gv -name "$($met)Meta" -ea 0)){$isBased = $false; $gvMiss += "$($met)Meta" } ;
            if($MetaProps){
                foreach($mp in $MetaProps){
                    write-verbose "chk:`$$($met)Meta.$($mp)" ;
                    #if(-not (gv -name "$($met)Meta" -ea 0).value[$mp]){ # testing has a value, not is present as a spec!
                    if(-not (gv -name "$($met)Meta" -ea 0).value.keys -contains $mp){
                        $isBased = $false; $ppMiss += "$($met)Meta.$($mp)"
                    } ;
                } ;
            } ;
        } ;
        if($gvNames){
            foreach($gvN in $gvNames){
                write-verbose "chk:`$$($gvN)" ;
                if(-not (gv -name "$($gvN)" -ea 0)){$isBased = $false; $gvMiss += "$($gvN)" } ;
            } ;
        } ;
        if($gvmiss){write-warning "Missing Dependant Meta variables:`n$(($gvMiss |%{"`$$($_)" }) -join ',')" } ;
        if($ppMiss){write-warning "Missing Dependant Meta vari properties:`n$(($ppMiss |%{"`$$($_)" }) -join ',')" } ;
        if(-not $isBased){ write-warning  "missing critical dependancy profile config!" } ;
        #endregion CHKPREREQ ; #*------^ END CHKPREREQ ^------

        if($Global:E10Mod){$Global:E10Mod | Remove-Module -Force -verbose:$($false) } ;
        if($Global:EXOPSess){$Global:EXOPSess | Remove-PSSession -verbose:$($false)} ;
        if($isBased){
            $Metas=(get-variable *meta|?{$_.name -match '^\w{3}Meta$'}).name ; 
            # 7:56 AM 11/1/2017 remove titlebar tag
            #Remove-PSTitlebar 'EMS' -verbose:$($VerbosePreference -eq "Continue")  ;
            # 9:21 AM 7/27/2021 expand to cover EMS[tlc]
            #Remove-PSTitlebar 'EMS[ctl]' -verbose:$($VerbosePreference -eq "Continue")  ;
            # make it fully dyn: build range of all meta 1sts & last chars
            [array]$chrs = $metas.substring(0,3).substring(0,1) ; 
            $chrs+= $metas.substring(0,3).substring(2,1) ; 
            $chrs=$chrs.tolower()|select -unique ;
            $sTitleBarTag = "EMS$('[' + (($chrs |%{[regex]::escape($_)}) -join '') + ']')" ; 
            write-verbose "remove PSTitleBarstring:$($sTitleBarTag)" ; 
            Remove-PSTitlebar $sTitleBarTag -verbose:$($VerbosePreference -eq "Continue")  ;
        }  ; 
        # should pull TenOrg if no other mounted 
        <#$sXopDesig = 'xp' ;
        $sXoDesig = 'xo' ;
        #>
        #$xxxMeta.rgxOrgSvcs : $ExchangeServer = (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server|get-random ;
        # normally would be org specific, but we don't have a cred or a TenOrg ref to resolve, so just check xx's version
        # -replace 'EMS','' -replace '\(\|','(' -replace '\|\)',')'
        #if($host.ui.RawUI.WindowTitle -notmatch ((Get-Variable  -name "TorMeta").value.rgxOrgSvcs-replace 'EMS','' -replace '\(\|','(' -replace '\|\)',')' )){
        # drop the current tag being removed from the rgx...
        <# # at this point, if we're no longer using explict Org tag (EMS[tlc] instead), don't need to remove, they'll come out with the EMS removel
        [regex]$rgxsvcs = ('(' + (((Get-Variable  -name "TorMeta").value.OrgSvcs |?{$_ -ne 'EMS'} |%{[regex]::escape($_)}) -join '|') + ')') ;
        if($host.ui.RawUI.WindowTitle -notmatch $rgxsvcs){
            write-verbose "(removing TenOrg reference from PSTitlebar)" ; 
            #Remove-PSTitlebar $TenOrg ;
            # split the rgx into an array of tags
            #sTitleBarTag = (((Get-Variable  -name "TorMeta").value.rgxOrgSvcs) -replace '(\\s\(|\)\\s)','').split('|') ; 
            # no remove all meta tenorg tags - shouldn't be cross-org connecting
            #$Metas=(get-variable *meta|?{$_.name -match '^\w{3}Meta$'}).name ; 
            $sTitleBarTag = $metas.substring(0,3) ; 
            Remove-PSTitlebar $sTitleBarTag -verbose:$($VerbosePreference -eq "Continue") ;
        } else {
            write-verbose "(detected matching OrgSvcs in PSTitlebar: *not* removing TenOrg reference)" ; 
        } ; 
        #>
    }  # BEG-E
    PROCESS{
        # kill any other sessions using distinctive name; add verbose, to ensure they're echo'd that they were missed
        #Get-PSSession | Where-Object { $_.name -eq 'Exchange2010' } | Remove-PSSession -verbose:$($false);
        <# Ex2013 [PS] C:\scripts>get-pssession | fl name,configuration
        Name              : Session1
        ConfigurationName : Microsoft.Exchange
        ComputerName      : server.domain.tld
        #>
        Get-PSSession | Where-Object { $_.ConfigurationName='Microsoft.Exchange'} | Remove-PSSession -verbose:$($false); #version agnostic
        Get-PSSession | Where-Object { $_.name -match 'Exchange2010' } | Remove-PSSession -verbose:$($false); # my older customized connection filtering
        

        # should splice in Ex2013/16 support as well
        # kill any broken PSS, self regen's even for L13 leave the original borked and create a new 'Session for implicit remoting module at C:\Users\', toast them, they don't reopen. Same for Ex2010 REMS, identical new PSS, indistinguishable from the L13 regen, except the random tmp_xxxx.psm1 module name. Toast them, it's just a growing stack of broken's
        Disconnect-PssBroken ;
        #[console]::ResetColor()  # reset console colorscheme
        # null $Global:EXOPSess 
        if($Global:EXOPSess){$Global:EXOPSess = $null } ; 
    } ;  # PROC-E
}

#*------^ Disconnect-Ex2010.ps1 ^------


#*------v enable-ForestView.ps1 v------
Function enable-ForestView {
    <#
    .SYNOPSIS
    enable-ForestView.ps1 - Enable Exchange onprem AD ViewEntireForest setting (permits org-wide object access, wo use of proper explicit -domaincontroller sub.domain.com)
    .NOTES
    Version     : 1.0.2
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2020-10-26
    FileName    : enable-ForestView
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    REVISIONS
    * 11:40 AM 5/14/2021 added -ea 0 to the gv tests (suppresses not-found error when called without logging config)
    * 10:56 AM 4/2/2021 cleaned up; added recstat & wlt
    * 11:43 AM 3/5/2021 variant of toggle-fv
    .DESCRIPTION
    enable-ForestView.ps1 - Enable Exchange onprem AD ViewEntireForest setting (permits org-wide object access, wo use of proper explicit -domaincontroller sub.domain.com)
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output
    .EXAMPLE
    enable-ForestView
    .LINK
    https://github.com/tostka/verb-ex2010
    .LINK
    #>
    [CmdletBinding()]
    PARAM() ;
    # toggle forest view
    if (get-command -name set-AdServerSettings){
        if (!(get-AdServerSettings).ViewEntireForest ) {
              write-verbose "(set-AdServerSettings -ViewEntireForest `$False)" ;
              set-AdServerSettings -ViewEntireForest $TRUE  ;
        } ;
    } else {
        #-=-record a STATUSERROR=-=-=-=-=-=-=
        $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
        if(Get-Variable passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
        if(Get-Variable -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
        #-=-=-=-=-=-=-=-=
        $smsg = "MISSING:set-AdServerSettings`nOPEN an Exchange OnPrem connection FIRST!"
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        BREAK ;
    } ;
}

#*------^ enable-ForestView.ps1 ^------


#*------v get-ADExchangeServerTDO.ps1 v------
if(-not(get-command get-ADExchangeServerTDO -ea 0)){
    Function get-ADExchangeServerTDO {
        <#
        .SYNOPSIS
        get-ADExchangeServerTDO.ps1 - Returns Exchangeserver summary(s) from AD records
        .NOTES
        Version     : 3.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2015-09-03
        FileName    : get-ADExchangeServerTDO.ps1
        License     : (none-asserted)
        Copyright   : (none-asserted)
        Github      : https://github.com/tostka/verb-Ex2010
        Tags        : Powershell, ActiveDirectory, Exchange, Discovery
        AddedCredit : Mike Pfeiffer
        AddedWebsite: mikepfeiffer.net
        AddedTwitter: URL
        AddedCredit : Sammy Krosoft 
        AddedWebsite: http://aka.ms/sammy
        AddedTwitter: URL
        AddedCredit : Brian Farnsworth
        AddedWebsite: https://codeandkeep.com/
        AddedTwitter: URL
        REVISIONS
        * 3:57 PM 11/26.2.34 updated simple write-host,write-verbose with full pswlt support;  syncd dbg & vx10 copies.
        * 12:57 PM 6/11/2024 Validated, Ex2010 & Ex2019, hub, mail & edge roles: tested ☑️ on CMW mail role (Curly); and Jumpbox; copied in CBH from repo copy, which has been updated/debugged compat on CMW Edge 
        * 2:05 PM 8/28/2023 REN -> Get-ExchangeServerInSite -> get-ADExchangeServerTDO (aliased orig); to better steer profile-level options - including in cmw org, added -TenOrg, and default Site to constructed vari, targeting new profile $XXX_ADSiteDefault vari; Defaulted -Roles to HUB,CAS as well.
        * 3:42 PM 8/24/2023 spliced together combo of my long-standing, and some of the interesting ideas BF's version had. Functional prod:
            - completely removed ActiveDirectory module dependancies from BF's code, and reimplemented in raw ADSI calls. Makes it fully portable, even into areas like Edge DMZ roles, where ADMS would never be installed.

        * 3:17 PM 8/23/2023 post Edge testing: some logic fixes; add: -Names param to filter on server names; -Site & supporting code, to permit lookup against sites *not* local to the local machine (and bypass lookup on the local machine) ; 
            ren $Ex10siteDN -> $ExOPsiteDN; ren $Ex10configNC -> $ExopconfigNC
        * 1:03 PM 8/22/2023 minor cleanup
        * 10:31 AM 4/7/2023 added CBH expl of postfilter/sorting to draw predictable pattern 
        * 4:36 PM 4/6.2.33 validated Psv51 & Psv20 and Ex10 & 16; added -Roles & -RoleNames params, to perform role filtering within the function (rather than as an external post-filter step). 
        For backward-compat retain historical output field 'Roles' as the msexchcurrentserverroles summary integer; 
        use RoleNames as the text role array; 
            updated for psv2 compat: flipped hash key lookups into properties, found capizliation differences, (psv2 2was all lower case, wouldn't match); 
        flipped the [pscustomobject] with new... psobj, still psv2 doesn't index the hash keys ; updated for Ex13+: Added  16  "UM"; 20  "CAS, UM"; 54  "MBX" Ex13+ ; 16385 "CAS" Ex13+ ; 16439 "CAS, HUB, MBX" Ex13+
        Also hybrided in some good ideas from SammyKrosoft's Get-SKExchangeServers.psm1 
        (emits Version, Site, low lvl Roles # array, and an array of Roles, for post-filtering); 
        # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
        * 12:08 PM 5/15/2020 fixed vpn issue: Try/Catch'd around recently failing $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName qry
        * 11:22 AM 3/13/2020 Get-ExchangeServerInSite added a ping-test, to only return matches that are pingable, added -NoPing param, to permit (faster) untested bypass
        * 6:59 PM 1/15/2020 cleanup
        # 10:03 AM 11/16/2018 Get-ExchangeServerInSite:can't do AD-related functions when not AD authentictaed (home, pre-vpn connect). Added if/then test on status and abort balance when false.
        * 11/18/18 BF's posted rev
        # 12:10 PM 8/1/2017 updated example code at bottom, to accommodate variant sites
        # 11:28 AM 3/31/2016 validated that latest round of updates are still functional
        #1:58 PM 9/3/2015 - added pshelp and some docs
        #April 12, 2010 - web version
        .DESCRIPTION
        get-ADExchangeServerTDO.ps1 - Returns Exchangeserver summary(s) from AD records

        Hybrided together ideas from Brian Farnsworth's blog post
        [PowerShell - ActiveDirectory and Exchange Servers – CodeAndKeep.Com – Code and keep calm...](https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/)
        ... with much older concepts from  Sammy Krosoft, and much earlier Mike Pfeiffer. 

        - Subbed in MP's use of ADSI for ActiveDirectory Ps mod cmds - it's much more dependancy-free; doesn't require explicit install of the AD ps module
        ADSI support is built into windows.
        - spliced over my addition of Roles, RoleNames, Name & NoTest params, for prefiltering and suppressing testing.


        [briansworth · GitHub](https://github.com/briansworth)

        Uses an ADSI DirectorySearcher to search the current Active Directory site for Exchange on-prem servers.
                Intent is to discover connection points for Powershell, wo the need to preload/pre-connect to Exchange.

                But, as a non-Exchange-Management-Shell-dependant info source on Exchange Server configs, it can be used before connection, with solely AD-available data, to check configuration spes on the subject server(s). 

                For example, this query will return sufficient data under Version to indicate which revision of Exchange is in use:


                Returned object (in array):
                Site      : {ADSITENAME}
                Roles     : {64}
                Version   : {Version 15.1 (Build 32375.7)}
                Name      : SERVERNAME
                RoleNames : EDGE
                FQDN      : SERVERNAME.DOMAIN.TLD

                ... includes the post-filterable Role property ($_.Role -contains 'CAS') which reflects the following
                installed-roles ('msExchCurrentServerRoles') on the discovered servers
                    2   {"MBX"} # Ex10
                    4   {"CAS"}
                    16  {"UM"}
                    20  {"CAS, UM" -split ","} # 
                    32  {"HUB"}
                    36  {"CAS, HUB" -split ","}
                    38  {"CAS, HUB, MBX" -split ","}
                    54  {"MBX"} # Ex13+
                    64  {"EDGE"}
                    16385   {"CAS"} # Ex13+
                    16439   {"CAS, HUB, MBX"  -split ","} # Ex13+
        .PARAMETER Roles
        Array of msExchCurrentServerRoles 'role' integers to be filtered against (2|4|16|20|32|36|38|54|64|16385|16439)[-Roles @(38,16385)]
        .PARAMETER RoleNames
        Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']
        .PARAMETER Server
        Array of Server name strings to be filtered against[-Server 'SERVER1','SERVER2']
        .PARAMETER SiteName
        Name of specific AD SiteName to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-SiteName 'SITENAME']
        .PARAMETER TenOrg
        Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
        .PARAMETER NoPing
        Switch to suppress default 'pingable' test (e.g. returns all matches, no testing)[-NoPing]
        .INPUTS
        None. Does not accepted piped input.(.NET types, can add description)
        .OUTPUTS
        None. Returns no objects or output (.NET types)
        System.Boolean
        [| get-member the output to see what .NET obj TypeName is returned, to use here]
        System.Array of System.Object's
        .EXAMPLE
        PS> If(!($ExchangeServer)){$ExchangeServer = (get-ADExchangeServerTDO| ?{$_.RoleNames -contains 'CAS' -OR $_.RoleNames -contains 'HUB' -AND ($_.FQDN -match "^SITECODE") } | Get-Random ).FQDN
        Return a random Hub Cas Role server in the local Site with a fqdn beginning SITECODE
        .EXAMPLE
        PS> $localADExchserver = get-ADExchangeServerTDO -Names $env:computername -SiteName ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().name)
        Demo, if run from an Exchange server, return summary details about the local server (-SiteName isn't required, is default imputed from local server's Site, but demos explicit spec for remote sites)
        .EXAMPLE
        PS> $regex = '(' + [regex]($ADSiteCodeUK,$ADSiteCodeAU -join "|") + ')'
        PS> switch -regex ($($env:computername).substring(0,3)){
        PS>    "$($ADSiteCodeUS)" {$tExRole=36 } ;
        PS>    "$($regex)" {$tExRole= 38 }  default { write-error "$((get-date).ToString('HH:mm:ss')):UNRECOGNIZED SERVER PREFIX!."; } ;
        PS> } ;
        PS> $exhubcas = (get-ADExchangeServerTDO |?{($_.roles -eq $tExRole) -AND ($_.FQDN -match "$($env:computername.substring(0,3)).*")} | Get-Random ).FQDN ;
        Use a switch block to select different role combo targets for a given server fqdn prefix string.
        .EXAMPLE
        PS> $ExchangeServer = get-ADExchangeServerTDO | ?{$_.Roles -match '(4|20|32|36|38|16385|16439)'} | select -expand fqdn | get-random ; 
        Another/Older approach filtering on the Roles integer (targeting combos with Hub or CAS in the mix)
        .EXAMPLE
        PS> $ret = get-ADExchangeServerTDO -Roles @(4,20,32,36,38,16385,16439) -verbose 
        Demo use of the -Roles param, feeding it an array of Role integer values to be filtered against. In this case, the Role integers that include a CAS or HUB role.
        .EXAMPLE
        PS> $ret = get-ADExchangeServerTDO -RoleNames 'HUB','CAS' -verbose ;
        Demo use of the -RoleNames param, feeding it the array 'HUB','CAS' Role name strings to be filtered against
        PS> $ret = get-ADExchangeServerTDO -Names 'SERVERName' -verbose ;
        Demo use of the -RoleNames param, feeding it the array 'HUB','CAS' Role name strings to be filtered against
        .EXAMPLE
        PS> $ExchangeServer = get-ADExchangeServerTDO | sort version,roles,name | ?{$_.rolenames -contains 'CAS'}  | select -last 1 | select -expand fqdn ;
        Demo post sorting & filtering, to deliver a rule-based predictable pattern for server selection: 
        Above will always pick the highest Version, 'CAS' RoleName containing, alphabetically last server name (that is pingable). 
        And should stick to that pattern, until the servers installed change, when it will shift to the next predictable box.
        .EXAMPLE
        PS> $ExOPServer = get-ADExchangeServerTDO -Name LYNMS650 -SiteName Lyndale
        PS> if([double]$ExVersNum = [regex]::match($ExOPServer.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
        PS>     switch -regex ([string]$ExVersNum) {
        PS>         '15\.2' { $isEx2019 = $true ; $ExVers = 'Ex2019' }
        PS>         '15\.1' { $isEx2016 = $true ; $ExVers = 'Ex2016'}
        PS>         '15\.0' { $isEx2013 = $true ; $ExVers = 'Ex2013'}
        PS>         '14\..*' { $isEx2010 = $true ; $ExVers = 'Ex2010'}
        PS>         '8\..*' { $isEx2007 = $true ; $ExVers = 'Ex2007'}
        PS>         '6\.5' { $isEx2003 = $true ; $ExVers = 'Ex2003'}
        PS>         '6|6\.0' {$isEx2000 = $true ; $ExVers = 'Ex2000'} ;
        PS>         default {
        PS>             $smsg = "UNRECOGNIZED ExchangeServer.AdminDisplayVersion.Major.Minor string:$($ExOPServer.version)! ABORTING!" ;
        PS>             write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
        PS>         }
        PS>     } ; 
        PS> }else {
        PS>     $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$ExOPServer.version:$($ExOPServer.version)!" ; 
        PS>     write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ; 
        PS>     throw $smsg ; 
        PS>     break ; 
        PS> } ; 
        Demo of parsing the returned Version property, into the proper Exchange Server revision.      
        .LINK
        https://github.com/tostka/verb-XXX
        .LINK
        https://bitbucket.org/tostka/powershell/
        .LINK
        http://mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
        .LINK
        https://github.com/SammyKrosoft/Search-AD-Using-Plain-PowerShell/blob/master/Get-SKExchangeServers.psm1
        .LINK
        https://github.com/tostka/verb-Ex2010
        .LINK
        https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/
        #>
        [CmdletBinding()]
        [Alias('Get-ExchangeServerInSite')]
        PARAM(
            [Parameter(Position=0,HelpMessage="Array of Server name strings to be filtered against[-Server 'SERVER1','SERVER2']")]
                [string[]]$Server,
            [Parameter(Position=1,HelpMessage="Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-SiteName 'SITENAME']")]
                [Alias('Site')]
                [string]$SiteName = (gv -name "$($TenOrg)_ADSiteDefaultName" -ea 0).Value,
            [Parameter(Position=2,HelpMessage="Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']")]
                [ValidateSet('MBX','CAS','HUB','UM','MBX','EDGE')]
                [string[]]$RoleNames = @('HUB','CAS'),
            [Parameter(HelpMessage="Array of msExchCurrentServerRoles 'role' integers to be filtered against (2|4|16|20|32|36|38|54|64|16385|16439)[-Roles @(38,16385)]")]
                [ValidateSet(2,4,16,20,32,36,38,54,64,16385,16439)]
                [int[]]$Roles,
            [Parameter(HelpMessage="Switch to suppress default 'pingable' test (e.g. returns all matches, no testing)[-NoTest]")]
                [Alias('NoPing')]
                [switch]$NoTest,
            [Parameter(HelpMessage="Milliseconds of max timeout to wait during port 80 test (defaults 100)[-SpeedThreshold 500]")]
                [int]$SpeedThreshold=100,
            [Parameter(Mandatory=$FALSE,HelpMessage="Tenant Tag (3-letter abbrebiation - defaults to global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']")]
                [ValidateNotNullOrEmpty()]
                [string]$TenOrg = $global:o365_TenOrgDefault,
            [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials[-Credentials [credential object]]")]
                [System.Management.Automation.PSCredential]$Credential
        ) ;
        BEGIN{
            ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
            $Verbose = ($VerbosePreference -eq 'Continue') ;
            $_sBnr="#*======v $(${CmdletName}): v======" ;
            $smsg = $_sBnr ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 }
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
        PROCESS{
            TRY{
                $configNC = ([ADSI]"LDAP://RootDse").configurationNamingContext ;
                $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC") ;
                $bLocalEdge = $false ; 
                if($Sitename -eq $env:COMPUTERNAME){
                    $smsg = "`$SiteName -eq `$env:COMPUTERNAME:$($SiteName):$($env:COMPUTERNAME)" ; 
                    $smsg += "`nThis computer appears to be an EdgeRole system (non-ADConnected)" ; 
                    $smsg += "`n(Blanking `$sitename and continuing discovery)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    #$bLocalEdge = $true ; 
                    $SiteName = $null ; 
                    
                } ; 
                If($siteName){
                    $smsg = "Getting Site: $siteName" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    $objectClass = "objectClass=site" ;
                    $objectName = "name=$siteName" ;
                    $search.Filter = "(&($objectClass)($objectName))" ;
                    $site = ($search.Findall()) ;
                    $siteDN = ($site | select -expand properties).distinguishedname  ;
                } else {
                    $smsg = "(No -Site specified, resolving site from local machine domain-connection...)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                    else{ write-host -foregroundcolor green "$($smsg)" } ;
                    TRY{$siteDN = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().GetDirectoryEntry().distinguishedName}
                    CATCH [System.Management.Automation.MethodInvocationException]{
                        $ErrTrapd=$Error[0] ;
                        if(($ErrTrapd.Exception -match 'The computer is not in a site.') -AND $env:ExchangeInstallPath){
                            $smsg = "$($env:computername) is non-ADdomain-connected" ;
                            $smsg += "`nand has `$env:ExchangeInstalled populated: Likely Edge Server" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                            else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $vers = (get-item "$($env:ExchangeInstallPath)\Bin\Setup.exe").VersionInfo.FileVersionRaw ; 
                            $props = @{
                                Name=$env:computername;
                                FQDN = ([System.Net.Dns]::gethostentry($env:computername)).hostname;
                                Version = "Version $($vers.major).$($vers.minor) (Build $($vers.Build).$($vers.Revision))" ; 
                                #"$($vers.major).$($vers.minor)" ; 
                                #$exServer.serialNumber[0];
                                Roles = [System.Object[]]64 ;
                                RoleNames = @('EDGE');
                                DistinguishedName =  "CN=$($env:computername),CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=First Organization,CN=Microsoft Exchange,CN=Services,CN=Configuration,CN={nnnnnnnn-FAKE-GUID-nnnn-nnnnnnnnnnnn}" ;
                                Site = [System.Object[]]'NOSITE'
                                ResponseTime = if($rsp){$rsp.ResponseTime} else { 0} ;
                                NOTE = "This summary object, returned for a non-AD-connected EDGE server, *approximates* what would be returned on an AD-connected server" ;
                            } ;
                            
                            $smsg = "(-NoTest:Defaulting Fast:`$true)" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $props.add('Fast',$true) ;
                            
                            return (New-Object -TypeName PsObject -Property $props) ;
                        }elseif(-not $env:ExchangeInstallPath){
                            $smsg = "Non-Domain Joined machine, with NO ExchangeInstallPath e-vari: `nExchange is not installed locally: local computer resolution fails:`nPlease specify an explicit -Server, or -SiteName" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $false | write-output ;
                        } else {
                            $smsg = "$($env:computername) is both NON-Domain-joined -AND lacks an Exchange install (NO ExchangeInstallPath e-vari)`nPlease specify an explicit -Server, or -SiteName" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $false | write-output ;
                        };
                    } CATCH {
                        $siteDN =$ExOPsiteDN ;
                        write-warning "`$siteDN lookup FAILED, deferring to hardcoded `$ExOPsiteDN string in infra file!" ;
                    } ;
                } ;
                $smsg = "Getting Exservers in Site:$($siteDN)" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC") ;
                $objectClass = "objectClass=msExchExchangeServer" ;
                $version = "versionNumber>=1937801568" ;
                $site = "msExchServerSite=$siteDN" ;
                $search.Filter = "(&($objectClass)($version)($site))" ;
                $search.PageSize = 1000 ;
                [void] $search.PropertiesToLoad.Add("name") ;
                [void] $search.PropertiesToLoad.Add("msexchcurrentserverroles") ;
                [void] $search.PropertiesToLoad.Add("networkaddress") ;
                [void] $search.PropertiesToLoad.Add("msExchServerSite") ;
                [void] $search.PropertiesToLoad.Add("serialNumber") ;
                [void] $search.PropertiesToLoad.Add("DistinguishedName") ;
                $exchServers = $search.FindAll() ;
                $Aggr = @() ;
                foreach($exServer in $exchServers){
                    $fqdn = ($exServer.Properties.networkaddress |
                        Where-Object{$_ -match '^ncacn_ip_tcp:'}).split(':')[1] ;
                    if($NoTest){} else {
                        $rsp = test-connection $fqdn -count 1 -ea 0 ;
                    } ;
                    $props = @{
                        Name = $exServer.Properties.name[0]
                        FQDN=$fqdn;
                        Version = $exServer.Properties.serialnumber
                        Roles = $exserver.Properties.msexchcurrentserverroles
                        RoleNames = $null ;
                        DistinguishedName = $exserver.Properties.distinguishedname;
                        Site = @("$($exserver.Properties.msexchserversite -Replace '^CN=|,.*$')") ;
                        ResponseTime = if($rsp){$rsp.ResponseTime} else { 0} ;
                    } ;
                    $props.RoleNames = switch ($exserver.Properties.msexchcurrentserverroles){
                        2       {"MBX"}
                        4       {"CAS"}
                        16      {"UM"}
                        20      {"CAS;UM".split(';')}
                        32      {"HUB"}
                        36      {"CAS;HUB".split(';')}
                        38      {"CAS;HUB;MBX".split(';')}
                        54      {"MBX"}
                        64      {"EDGE"}
                        16385   {"CAS"}
                        16439   {"CAS;HUB;MBX".split(';')}
                    }
                    if($NoTest){
                        $smsg = "(-NoTest:Defaulting Fast:`$true)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $props.add('Fast',$true) ;
                    }else {
                        $props.add('Fast',[boolean]($rsp.ResponseTime -le $SpeedThreshold)) ;
                    };
                    $Aggr += New-Object -TypeName PsObject -Property $props ;
                } ;
                $httmp = @{} ;
                if($Roles){
                    [regex]$rgxRoles = ('(' + (($roles |%{[regex]::escape($_)}) -join '|') + ')') ;
                    $matched =  @( $aggr | ?{$_.Roles -match $rgxRoles}) ;
                    foreach($m in $matched){
                        if($httmp[$m.name]){} else {
                            $httmp[$m.name] = $m ;
                        } ;
                    } ;
                } ;
                if($RoleNames){
                    foreach ($RoleName in $RoleNames){
                        $matched = @($Aggr | ?{$_.RoleNames -contains $RoleName} ) ;
                        foreach($m in $matched){
                            if($httmp[$m.name]){} else {
                                $httmp[$m.name] = $m ;
                            } ;
                        } ;
                    } ;
                } ;
                if($Server){
                    foreach ($Name in $Server){
                        $matched = @($Aggr | ?{$_.Name -eq $Name} ) ;
                        foreach($m in $matched){
                            if($httmp[$m.name]){} else {
                                $httmp[$m.name] = $m ;
                            } ;
                        } ;
                    } ;
                } ;
                if(($httmp.Values| measure).count -gt 0){
                    $Aggr  = $httmp.Values ;
                } ;
                $smsg = "Returning $((($Aggr|measure).count|out-string).trim()) match summaries to pipeline..." ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                $Aggr | write-output ;
            }CATCH{
                Write-Error $_ ;
            } ;
        } ;
        END{
            $smsg = "$($_sBnr.replace('=v','=^').replace('v=','^='))" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 }
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
    } ;
}

#*------^ get-ADExchangeServerTDO.ps1 ^------


#*------v get-DAGDatabaseCopyStatus.ps1 v------
function get-DAGDatabaseCopyStatus {
    <#
    .SYNOPSIS
    get-DAGDatabaseCopyStatus.ps1 - Retrieve MailboxDatabaseCopyStatus for an entire DatabaseAvailabilityGroup
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2022-04-25
    FileName    : get-DAGDatabaseCopyStatus.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell,Exchange,ExchangeOnline,ActiveDirectory
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 2:16 PM 6/24/2024: rem'd out #Requires -RunasAdministrator; sec chgs in last x mos wrecked RAA detection
    4:23 PM 3/21/2023 typo, missing trailing , on $name param def
    11:08 AM 4/25/2022 init
    .DESCRIPTION
    get-DAGDatabaseCopyStatus.ps1 - Retrieve MailboxDatabaseCopyStatus for an entire DatabaseAvailabilityGroup
    .PARAMETER  users
    Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)
    .PARAMETER ADDisabledOnly
    Switch to exclude users solely on ADUser.disabled (not Disabled OU presense), or with that have the ADUser below an OU matching '*OU=(Disabled|TERMedUsers)'  [-ADDisabledOnly]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    Returns a System.Object[] object to pipeline, with subsets of processed users as 'Enabled' (ADUser.enabled),'Disabled', and 'Contacts' properties.
    .EXAMPLE
    PS> $rpt = get-DAGDatabaseCopyStatus -users 'username1','user2@domain.com','[distinguishedname]' ;
    PS> $rpt | export-csv -nottype ".\pathto\usersummaries-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
    Processes local/remotemailbox & ADUser details on three specified users (alias,email address, DN). Summaries are returned to pipeline, and assigned to the $rpt variable, which is then exported to csv.
    .EXAMPLE
    PS> $rpt = get-DAGDatabaseCopyStatus -users 'username1','user2@domain.com','[distinguishedname]' -ADDisabledOnly ;
    PS> $rpt | export-csv -nottype ".\pathto\usersummaries-ENABLEDUSERS-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
    Processes local/remotemailbox & ADUser details on three specified users (alias,email address, DN).
    And allocate as 'Disabled', accounts that are *solely* ADUser.disabled
    (e.g. considers users below OU's with names like 'OU=Disabled*' as 'Enabled' users),
    and then exports to csv.
    .EXAMPLE
    $rpt = get-DAGDatabaseCopyStatus -users 'username1','user2@domain.com','[distinguishedname]' ;
    $rpt.enabled | export-csv -nottype ".\pathto\usersummaries-ENABLEDUSERS-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
    Process specified identifiers, and export solely the 'Enabled' users returned to csv.
    .LINK
    https://github.com/tostka/verb-ex2010
    .LINK
    #>
    ###Requires -Version 5
    #Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS
    ##Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    ##[Alias('ulu')]
    PARAM(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true,HelpMessage="DatabaseAvailabilityGroup Name")]
        #[ValidateNotNullOrEmpty()]
        #[Alias('ALIAS1', 'ALIAS2')]
        [array]$Name,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2
    ) ;
    BEGIN{
        $Verbose = ($VerbosePreference -eq 'Continue') ;

        $rgxDisabledOUs = '.*OU=(Disabled|TERMedUsers).*' ;
        if(!$users){
            $users= (get-clipboard).trim().replace("'",'').replace('"','') ;
            if($users){
                write-verbose "No -users specified, detected value on clipboard:`n$($users)" ;
            } else {
                write-warning "No -users specified, nothing matching dname, emailaddress or samaccountname, found on clipboard. EXITING!" ;
                Break ;
            } ;
        } ;
        $ttl = ($users|measure).count ;
        if($ttl -lt 10){
            write-verbose "($(($users|measure).count)) user(s) specified:`n'$($users -join "','")'" ;
        } else {
            write-verbose "($(($users|measure).count)) user(s) specified (-gt 10, suppressing details)" ;
        }

        rx10 -Verbose:$false ; #rxo  -Verbose:$false ; #cmsol  -Verbose:$false ;
        connect-ad -Verbose:$false;

        $propsmbx='Database','UseDatabaseRetentionDefaults','SingleItemRecoveryEnabled','RetentionPolicy','ProhibitSendQuota',
            'ProhibitSendReceiveQuota','SamAccountName','ServerName','UseDatabaseQuotaDefaults','IssueWarningQuota','Office',
            'UserPrincipalName','Alias','OrganizationalUnit  global.ad.toro.com/TERMedUsers','DisplayName','EmailAddresses',
            'HiddenFromAddressListsEnabled','LegacyExchangeDN','PrimarySmtpAddress','RecipientType','RecipientTypeDetails',
            'WindowsEmailAddress','DistinguishedName','CustomAttribute1','CustomAttribute2','CustomAttribute3','CustomAttribute4',
            'CustomAttribute5','CustomAttribute6','CustomAttribute7','CustomAttribute8','CustomAttribute9','CustomAttribute10',
            'CustomAttribute11','CustomAttribute12','CustomAttribute13','CustomAttribute14','CustomAttribute15''EmailAddressPolicyEnabled',
            'WhenChanged','WhenCreated' ;
        $propsadu = "accountExpires","CannotChangePassword","Company","Compound","Country","countryCode","Created","Department",
            "Description","DisplayName","DistinguishedName","Division","EmployeeID","EmployeeNumber","employeeType","Enabled","Fax",
            "GivenName","homeMDB","homeMTA","info","Initials","lastLogoff","lastLogon","LastLogonDate","mail","mailNickname","Manager",
            "mobile","MobilePhone","Modified","Name","Office","OfficePhone","Organization","physicalDeliveryOfficeName","POBox","PostalCode",
            "SamAccountName","sAMAccountType","State","StreetAddress","Surname","Title","UserPrincipalName",'CustomAttribute1',
            'CustomAttribute2','CustomAttribute3','CustomAttribute4','CustomAttribute5','CustomAttribute6','CustomAttribute7',
            'CustomAttribute8','CustomAttribute9','CustomAttribute10','CustomAttribute11','CustomAttribute12','CustomAttribute13',
            'CustomAttribute14','CustomAttribute15','EmailAddressPolicyEnabled',"whenChanged","whenCreated" ;
        $propsMC = 'ExternalEmailAddress','Alias','DisplayName','EmailAddresses','PrimarySmtpAddress','RecipientType',
            'RecipientTypeDetails','WindowsEmailAddress','Name','DistinguishedName','Identity','CustomAttribute1','CustomAttribute2',
            'CustomAttribute3','CustomAttribute4','CustomAttribute5','CustomAttribute6','CustomAttribute7','CustomAttribute8',
            'CustomAttribute9','CustomAttribute10','CustomAttribute11','CustomAttribute12','CustomAttribute13','CustomAttribute14',
            'CustomAttribute15','EmailAddressPolicyEnabled','whenChanged','whenCreated' ;
    }
    PROCESS{
        $Procd=0 ;$pct = 0 ;
        $aggreg =@() ; $contacts =@() ; $UnResolved = @() ; $Failed = @() ;
        $pltGRcp=[ordered]@{identity=$null;erroraction='STOP';resultsize=1;} ;
        $pltGMbx=[ordered]@{identity=$null;erroraction='STOP'} ;
        $pltGadu=[ordered]@{Identity = $null ; Properties='*' ;errorAction='STOP'} ;
        foreach ($usr in $users){
            $procd++ ; $pct = '{0:p0}' -f ($procd/$ttl) ;
            $rrcp = $mbx = $mc = $mbxspecs = $adspecs = $summary = $NULL ;
            #write-verbose "processing:$($usr)" ;
            $sBnrS="`n#*------v PROCESSING ($($procd)/$($ttl)):$($usr)`t($($pct)) v------" ;
            if($verbose){
                write-verbose "$((get-date).ToString('HH:mm:ss')):$($sBnrS)" ;
            } else {
                write-host "." -NoNewLine ;
            } ;

            TRY {
                $pltGRcp.identity = $usr ;
                write-verbose "get-recipient  w`n$(($pltGRcp|out-string).trim())" ;
                $rrcp = get-recipient @pltGRcp ;
                if($rrcp){
                    $pltgmbx.identity = $rrcp.PrimarySmtpAddress ;
                    switch ($rrcp.recipienttype){
                        'MailUser'{
                            write-verbose "get-remotemailbox  w`n$(($pltgmbx|out-string).trim())" ;
                            $mbx = get-remotemailbox @pltgmbx
                        }
                        'UserMailbox' {
                            write-verbose "get-mailbox w`n$(($pltgmbx|out-string).trim())" ;
                            $mbx = get-mailbox @pltgmbx ;
                        }
                        'MailContact' {
                            write-verbose "get-mailcontact w`n$(($pltgmbx|out-string).trim())" ;
                            $mc = get-mailcontact @pltgmbx ;
                        }
                        default {throw "$($rrcp.alias):Unsupported RecipientType:$($rrcp.recipienttype)" }
                    } ;
                    if(-not($mc)){
                        $mbxspecs =  $mbx| select $propsmbx ;
                        $pltGadu.identity = $mbx.samaccountname ;
                        write-verbose "Get-ADUser w`n$(($pltGadu|out-string).trim())" ;
                        Try {
                            $adspecs =Get-ADUser @pltGadu | select $propsadu ;
                        } CATCH [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                            $smsg = "(no matching ADuser found:$($pltGadu.identity))" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Continue ;
                        } catch {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Continue ;
                        } ;
                        $summary = @{} ;
                        foreach($object_properties in $mbxspecs.PsObject.Properties) {
                            $summary.add($object_properties.Name,$object_properties.Value) ;
                        } ;
                        foreach($object_properties in $adspecs.PsObject.Properties) {
                            $summary.add("AD$($object_properties.Name)",$object_properties.Value) ;
                        } ;
                        $aggreg+= New-Object PSObject -Property $summary ;

                    } else {
                        $smsg = "Resolved user for $($usr) is RecipientType:$($mc.RecipientType)`nIt is not a local mail object, or AD object, and simply reflects a pointer to an external mail recipient.`nThis object is being added to the 'Contacts' section of the output.." ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $contacts += $mc | select $propsMC ;
                    } ;
                } else {
                    # in ISE $Error[0] is empty
                    #$ErrTrapd=$Error[0] ;
                    #if($ErrTrapd.Exception -match "couldn't\sbe\sfound\son"){
                        $UnResolved += $pltGRcp.identity ;
                    #} else {
                        #$Failed += $pltGRcp.identity ;
                    #} ;
                } ;
            } CATCH [System.Management.Automation.RemoteException] {
                # catch error never gets here (at least not in ISE)
                $ErrTrapd=$Error[0] ;
                if($ErrTrapd.Exception -match "couldn't\sbe\sfound\son"){
                    $UnResolved += $pltGRcp.identity ;
                } else {
                    $Failed += $pltGRcp.identity ;
                } ;
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $Failed += $pltGRcp.identity ;
            } ;

            if($verbose){
                write-verbose "$((get-date).ToString('HH:mm:ss')):$($sBnrS.replace('-v','-^').replace('v-','^-'))" ; ;
            } ;

        } ;
    }
    END{
        if(-not($ADDisabledOnly)){
            $Report = [ordered]@{
                Enabled = $Aggreg|?{($_.ADEnabled -eq $true ) -AND -not($_.distinguishedname -match $rgxDisabledOUs) } ;#?{$_.adDisabled -ne $true -AND -not($_.distinguishedname -match $rgxDisabledOUs)}
                Disabled = $Aggreg|?{($_.ADEnabled -eq $False) } ;
                Contacts = $contacts ;
                Unresolved = $Unresolved ;
                Failed = $Failed;
            } ;
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):outputing $(($Report.Enabled|measure).count) Enabled User summaries,`nand $(($Report.Disabled|measure).count) ADUser.Disabled or Disabled/TERM-OU account summaries`nand $(($Report.Contacts|measure).count) users resolved to MailContacts" ;
        } else {
            $Report = [ordered]@{
                Enabled = $Aggreg|?{($_.ADEnabled -eq $true) -AND -not($_.distinguishedname -match $rgxDisabledOUs) } ;#?{$_.adDisabled -ne $true -AND -not($_.distinguishedname -match $rgxDisabledOUs)}
                Disabled = $Aggreg|?{($_.ADEnabled -eq $False) -OR ($_.distinguishedname -match $rgxDisabledOUs) } ;
                Contacts = $contacts ;
                Unresolved = $Unresolved ;
                Failed = $Failed;
            } ;
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):outputing $(($Report.Enabled|measure).count) Enabled User summaries,`nand $(($Report.Disabled|measure).count) ADUser.Disabled`nand $(($Report.Contacts|measure).count) users resolved to MailContacts" ;
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):-ADDisabledOnly specified: 'Disabled' output are *solely* ADUser.Disabled (no  Disabled/TERM-OU account filtering applied)`nand $(($Report.Contacts|measure).count) users resolved to MailContacts" ;
        } ;
        New-Object PSObject -Property $Report | write-output ;

     } ;
 }

#*------^ get-DAGDatabaseCopyStatus.ps1 ^------


#*------v Get-ExchServerFromExServersGroup.ps1 v------
Function Get-ExchServerFromExServersGroup {
  <#
    .SYNOPSIS
    Get-ExchServerFromExServersGroup - Returns the name of an Exchange server by drawing a random box from ad.DOMAIN.com\Exchange Servers grp & regex matches for desired site hubCas names.
    .NOTES
    Author: Todd Kadrie
    Website:	http://tintoys.blogspot.com
    REVISIONS   :
    * 10:02 AM 5/15/2020 pushed the post regex into a infra string & defaulted param, so this could work with any post-filter ;ren Get-ExchServerInLYN -> Get-ExchServerFromExServersGroup
    * 6:59 PM 1/15/2020 cleanup
    # 10:44 AM 9/2/2016 - initial tweak
    .PARAMETER  ServerRegex
    Server filter Regular Expression[-ServerRegex '^CN=(SITE1|SITE2)BOX1[0,1].*$']
    .DESCRIPTION
    Get-ExchServerFromExServersGroup - Returns the name of an Exchange server by drawing a random box from ad.DOMAIN.com\Exchange Servers grp & regex matches for desired site hubCas names.
    Leverages the ActiveDirectory module Get-ADGroupMember cmdlet
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns the name of an Exchange server in the local AD site.
    .EXAMPLE
    .\Get-ExchServerFromExServersGroup
    Draw random matching ex server with defaulted settings
    .EXAMPLE
    .\Get-ExchServerFromExServersGroup -ServerRegex '^CN=SITEPREFIX.*$'
    Draw random matching ex server with explicit ServerRegex match
    .LINK
    #>
    #Requires -Modules ActiveDirectory
    PARAM(
        [Parameter(HelpMessage="Server filter Regular Expression[-ServerRegex '^CN=(SITE1|SITE2)BOX1[0,1].*$']")]
        $ServerRegex=$rgxLocalHubCAS,
        [Parameter(HelpMessage="AD ParentDomain fqdn [-ADParentDomain 'ROOTDOMAIN.DOMAIN.com']")]
        $ADParentDomain=$DomTORParentfqdn
    ) ;
    (Get-ADGroupMember -Identity 'Exchange Servers' -server $DomTORParentfqdn |
        Where-Object { $_.distinguishedname -match $ServerRegex }).name |
            get-random | write-output ;
}

#*------^ Get-ExchServerFromExServersGroup.ps1 ^------


#*------v get-ExRootSiteOUs.ps1 v------
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
        [ValidateNotNullOrEmpty()][string]$Regex = '^OU=(\w{3}|PACRIM),DC=global,DC=ad,DC=toro((lab)*),DC=com$',
        [Parameter(Position = 0, HelpMessage = "OU DistinguishedName regex, to EXCLUDE non-legitimate 'Site' OUs [-RegexBanned [regularexpression]]")]
        [ValidateNotNullOrEmpty()][string]$RegexBanned = '^OU=(BCC|EDC|NC1|NDS|TAC),DC=global,DC=ad,DC=toro((lab)*),DC=com$',
        #[Parameter(HelpMessage = "Domain Controller [-domaincontroller server.fqdn.com]")]
        #[string] $domaincontroller,
        [Parameter(HelpMessage = "Debugging Flag [-showDebug]")]
        [switch] $showDebug
    ) # PARAM BLOCK END
    $verbose = ($VerbosePreference -eq "Continue") ; 
    $error.clear() ;
    TRY {
        #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
        $SiteOUs = Get-OrganizationalUnit |?{($_.distinguishedname -match $Regex) -AND ($_.distinguishedname -notmatch $RegexBanned) }|sort distinguishedname ; 

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
        $smsg= "Unable to retrieve OUs matching specified rgx:`n$($regex)";
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $false | write-output ;
    }
}

#*------^ get-ExRootSiteOUs.ps1 ^------


#*------v get-MailboxDatabaseQuotas.ps1 v------
function get-MailboxDatabaseQuotas {
<#
    .SYNOPSIS
    get-MailboxDatabaseQuotas - Queries all on-prem mailbox databases (get-mailboxdatabase) for default quota settings, and returns an indexed hashtable summarizing the values per database (indexed to each database 'name' value).
    .PARAMETER Mailboxes
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2022-02-25
    FileName    : get-MailboxDatabaseQuotas.ps1
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell
    REVISIONS
    * 4:27 PM 2/25/2022 init vers
    .DESCRIPTION
    get-MailboxDatabaseQuotas - Queries all on-prem mailbox databases (get-mailboxdatabase) for default quota settings, and returns an indexed hashtable summarizing the name and quotas per database (indexed to each database 'name' value).
    .PARAMETER TenOrg
TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']
    .EXAMPLE
    PS> $hQuotas = get-MailboxDatabaseQuotas -verbose ; 
    PS> $hQuotas['database2']
    Name           ProhibitSendReceiveQuotaGB ProhibitSendQuotaGB IssueWarningQuotaGB
    ----           -------------------------- ------------------- -------------------
    database2      12.000                     10.000              9.000
    Retrieve local org on-prem MailboxDatabase quotas and assign to a variable, with verbose outputs. Then output the retrieved quotas from the indexed hash returned, for the mailboxdatabase named 'database2'.
    .EXAMPLE
    PS> $pltGMDQ=[ordered]@{
            TenOrg= $TenOrg;
            verbose=$($VerbosePreference -eq "Continue") ;
            credential= $pltRXO.credential ;
            #(Get-Variable -name cred$($tenorg) ).value ;
        } ;
    PS> $smsg = "$($tenorg):get-MailboxDatabaseQuotas w`n$(($pltGMDQ|out-string).trim())" ;
    PS> if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS> $objRet = $null ;
    PS> $objRet = get-MailboxDatabaseQuotas @pltGMDQ ;
    PS> switch -regex ($objRet.GetType().FullName){
            "(System.Collections.Hashtable|System.Collections.Specialized.OrderedDictionary)" {
                if( ($objRet|Measure-Object).count ){
                    $smsg = "get-MailboxDatabaseQuotas:$($tenorg):returned populated MailboxDatabaseQuotas" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $mdbquotas = $objRet ;
                } else {
                    $smsg = "get-MailboxDatabaseQuotas:$($tenorg):FAILED TO RETURN populated MailboxDatabaseQuotas" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    THROW $SMSG ; 
                    break ; 
                } ;
            }
            default {
                $smsg = "get-MailboxDatabaseQuotas:$($tenorg):RETURNED UNDEFINED OBJECT TYPE!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Exit ;
            } ;
        } ;  
    PS> $smsg = "$(($mdbquotas|measure).count) quota summaries returned)" ;
    PS> if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    # given populuated $mbx 'mailbox object', lookup demo:
    PS> if($mbx.UseDatabaseQuotaDefaults){
            $MbxProhibitSendQuota = $mdbquotas[$mbx.database].ProhibitSendQuota ;
            $MbxProhibitSendReceiveQuota = $mdbquotas[$mbx.database].ProhibitSendReceiveQuota ;
            $MbxIssueWarningQuota = $mdbquotas[$mbx.database].IssueWarningQuota ;
        } else {
            write-verbose "(Custom Mbx Quotas configured...)" ;
            $MbxProhibitSendQuota = $mbx.ProhibitSendQuota ;
            $MbxProhibitSendReceiveQuota = $mbx.ProhibitSendReceiveQuota ;
            $MbxIssueWarningQuota = $mbx.IssueWarningQuota ;
        } ;    
    Expanded example with testing of returned object, and demoes use of the returned hash against a mailbox spec, steering via .UseDatabaseQuotaDefaults
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    #Requires -Modules verb-IO, verb-logging, verb-Text
    [OutputType('System.Collections.Hashtable')]
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        [string]$TenOrg = 'TOR',
        [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credTORSID
    ) ;
    
    ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
    $verbose = ($VerbosePreference -eq "Continue") ;
    
    # select db properties (converts dehydrated bytes string values to decimal gigabytes, via my verb-io module's convert-DehydratedBytesToGB())
    $propsMDB = 'Name',@{Name='ProhibitSendReceiveQuotaGB';Expression={$_.ProhibitSendReceiveQuota | convert-DehydratedBytesToGB }},
    @{Name='ProhibitSendQuotaGB';Expression={$_.ProhibitSendQuota | convert-DehydratedBytesToGB }},
    @{Name='IssueWarningQuotaGB';Expression={$_.IssueWarningQuota | convert-DehydratedBytesToGB }} ; 
    #'ProhibitSendReceiveQuota','ProhibitSendQuota','IssueWarningQuota' ; 
    
    #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======
#region GENERIC_EXOP_CREDS_&_SRVR_CONN #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
    # steer all onprem code on $XXXMeta.ExOPAccessFromToro & Ex10Server values
    $UseExOP=$true ;
    <# no onprem dep
    if((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro -AND (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server){
        $UseExOP = $true ;
        $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`ENABLING use of OnPrem Ex system this pass." ;
        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
    } else {
        $UseExOP = $false ;
        $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`nDISABLING use of OnPrem Ex system this pass." ;
        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
    } ;
    #>
    if($UseExOP){
        #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
        # do the OP creds too
        $OPCred=$null ;
        # default to the onprem svc acct
        $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC','SID'; verbose=$($verbose)} ;
        if($Credential){
            $pltGHOpCred.add('Credential',$Credential) ;
            if(get-Variable -Name "cred$($tenorg)OP" -scope Script -ea 0){
                set-Variable -Name "cred$($tenorg)OP" -scope Script -Value $Credential ;
            } else { New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $Credential } ;
        } else { 
            if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name "cred$($tenorg)OP" -scope Script -ea 0 ){ remove-Variable -Name "cred$($tenorg)OP" -scope Script } ;
                New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ;
                if($verbose){
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                Break ;
            } ;
        } ; 
        $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        <# CALLS ARE IN FORM: (cred$($tenorg))
            $pltRX10 = @{
            Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
            #verbose = $($verbose) ;
            Verbose = $FALSE ; Silent = $true ; } ;
        Reconnect-Ex2010 @pltRX10 ; # local org conns
        #$pltRx10 creds & .username can also be used for local ADMS connections
        #>
        $pltRX10 = @{
            Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
            #verbose = $($verbose) ;
            Verbose = $FALSE ; Silent = $false ; } ;
        if($1stConn){
            $pltRX10.silent = $false ; 
        } else { 
            $pltRX10.silent = $true ; 
        } ; 
        # defer cx10/rx10, until just before get-recipients qry
        #endregion GENERIC_EXOP_CREDS_&_SRVR_CONN #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
        # connect to ExOP X10
        if($pltRX10){
            #ReConnect-Ex2010XO @pltRX10 ;
            ReConnect-Ex2010 @pltRX10 ;
        } else { Reconnect-Ex2010 ; } ;
    } ;  # if-E $useEXOP

    # check if using Pipeline input or explicit params:
    if ($PSCmdlet.MyInvocation.ExpectingInput) {
        write-verbose "Data received from pipeline input: '$($InputObject)'" ;
    } else {
        # doesn't actually return an obj in the echo
        #write-verbose "Data received from parameter input: '$($InputObject)'" ;
    } ;
    
    # building a CustObj (actually an indexed hash) with the default quota specs from all db's. The 'index' for each db, is the db's Name (which is also stored as Database on the $mbx)
    if($host.version.major -gt 2){$dbQuotas = [ordered]@{} } 
    else { $dbQuotas = @{} } ;
    
    $smsg = "(querying quotas from all local-org mailboxdatabases)" ; 
    if($verbose){
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    } ; 
    
    $error.clear() ;
    TRY {
        $dbQuotaDefaults=(get-mailboxdatabase -erroraction 'STOP' | sort server,name | select $propsMDB ) ;
    } CATCH {
        $ErrTrapd=$Error[0] ;
        $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #-=-record a STATUSWARN=-=-=-=-=-=-=
        $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
        if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
        if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
        #-=-=-=-=-=-=-=-=
        $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
    } ; 
    
    $ttl = ($dbQuotaDefaults|measure).count ; $Procd = 0 ; 
    foreach ($db in $dbQuotaDefaults){
        $Procd ++ ; 
        $sBnrS="`n#*------v PROCESSING : ($($Procd)/$($ttl)) $($db.name) v------" ; 
        $smsg = $sBnrS ; 
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        
        $name =$($db | select -expand Name) ; 
        $dbQuotas[$name] = $db ; 

        $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
    } ;  # loop-E

    if($dbQuotas){
        $smsg = "(Returning summary objects to pipeline)" ; 
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        $dbQuotas | Write-Output ; 
    } else {
        $smsg = "NO RETURNABLE `$dbQuotas OBJECT!" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        THROW $smsg ;
    } ; 
}

#*------^ get-MailboxDatabaseQuotas.ps1 ^------


#*------v Get-MessageTrackingLogTDO.ps1 v------
function Get-MessageTrackingLogTDO {
    <#
    .SYNOPSIS
    Get-MessageTrackingLogTDO - Wrapper that stages everything needed to discover ADSite & Servers, and open REMS connection to mail servers, to run a Get-MessageTrackingLog pass: has all comments pulled: *should* unwrap, but can run stacked as well. Also runs natively in EMS. Center unwrapped block is stock 7psmsgboxall
    .NOTES
    Version     : 0.0.1
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 20240530-1042AM
    FileName    : Get-MessageTrackingLogTDO.ps1
    License     : MIT License
    Copyright   : (c) 2024 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell,Exchange,MessageTracking,Get-MessageTrackingLog,ActiveDirectory
    REVISIONS
    * 3:23 PM 12/2/2024 throwing param transform errs on start & end (wyhen typed): pull typing, and do it post assignh, can't assign empty '' or $null to t a datetime coerced vary ;pre-reduce splat hash to populated values, in exmpl & BP use;
         rem out the parameterset code, and just do manual conflicting -start/-end -days tests and errors
    * 2:34 PM 11/26.2.34 updated to latest 'Connect-ExchangeServerTDO()','get-ADExchangeServerTDO()', set to defer to existing
    * 4:20 PM 11/25/2024 updated from get-exomessagetraceexportedtdo(), more silent suppression, integrated dep-less ExOP conn supportadd delimters to echos, to space more, readability ;  fixed typo in eventid histo output
    * 3:16 PM 11/21/2024 working: added back Connectorid (postfiltered from results); add: $DaysLimit = 30 ; added: MsgsFail, MsgsDefer, MsgsFailRcpStat; 
    * 2:00 PM 11/20/2024 rounded out to iflv level, no dbg yet
    * 5:00 PM 10/14/2024 at this point roughed in ported updates from get-exomsgtracedetailed, no debugging/testing yet; updated params & cbh to semi- match vso\get-exomsgtracedetailed(); convert to a function (from ps1)
    * 11:30 AM 7/16.2.34 CBH example typo fix
    * 2:41 PM 7/10/2024 spliced in notices for plus-addressing, ren'd $Tix -> $Ticket (matches gbug fi in psb.)
    * 12:57 PM 6/11/2024 Validated, Ex2010 & Ex2019, hub, mail & edge roles: tested ☑️ on CMW mail role (Curly); and Jumpbox;  finished test/debug on CMW edge: appears full functioning, including get-ADExchangeServerTDO() & Connect-ExchangeServerTDO()
    * 3:46 PM 6/3/2024 WIP for edge, latest chgs: rounded out params to subst cover full range of underlying Get-MessageTrackingLog: MessageID ; InternalMessageID; NetworkMessageID; Reference; ResultSize; 
        incl CH Example splatted call; shift conflicting sub bnr into $_sBnr; Also added param valid accomodating ResultSize is int32, $null or ''
    * 6;09 PM 5/30/2024 WIP for edge, finally got edge EMS code spliced in (from https://github.com/Lucifer1993/PLtools/blob/main/HealthChecker.ps1); 
    now connects see Get-MessageTrackingLog  runs & returns content on Edge; has trailing bug on exit
    You cannot call a method on a null-valued expression.
    At C:\scripts\Get-MessageTrackingLogTDO.ps1:952 char:16
    +     $smsg = "$($sBnr.replace('=v','=^').replace('v=','^='))" ;
    * 11:56 AM 5/30/2024 init; add: out-clipboard() ; spliced in conditional ordered hash code; transplanted psMsgTrkCMW.cbp into a full blown function; bonus: it runs fine in either org/enviro, as it's a full self contained solution to discover the local Exchange org from AD, then connect to the systems. 
    .DESCRIPTION
    Get-MessageTrackingLogTDO - Wrapper that stages everything needed to discover ADSite & Exchange Servers in the site; open REMS connection to a local HubCAS mail server;and then run a Get-MessageTrackingLog pass: has all comments pulled: *should* unwrap, but can run stacked as well. Also runs natively in EMS. Center unwrapped block is stock 7psmsgboxall

    SET DAYS=0 IF USING START/END (they only get used when days is non-0); isplt.
    TAG is appended to ticketNO for output vari $vn, and $ofile

    Returns Summary object to pipeline:
    [obj].MTMessagesCSVFile: full path to exported MTMessages as csv file 
    [obj].MTMessages: MessageTracking messages matched
    [obj].EventIDHisto: Histogram of EventID entries for MTMessages array
    [obj].MsgLast: Last Message returned on track
    [obj].MsgsFail: EventID:Fail messages returned on track
    [obj].MsgsDefer: EventID:Defer messages returned on track
    [obj].MsgsFailRcpStatHisto: Histogram of RecipientStatus entries for MTMessages array

    .PARAMETER ticket
    Ticket Number [-Ticket 'TICKETNO']
    .PARAMETER Requestor
    Ticket Customer email identifier. [-Requestor 'fname.lname@domain.com']
    .PARAMETER Days
    Days to be searched, back from current time(Alt to use of StartDate & EndDate)[-Days 30]
    .PARAMETER Start
    Optional search Start timestamp[-Start '8:55 AM 5/30/2024']
    .PARAMETER End
    Optional search End timestamp[-Start '9:55 AM 5/30/2024']
    .PARAMETER Sender
    Sender Address[-Sender email@domain.com]
    .PARAMETER Recipients
    Recipient Addresses[-Recipients 'eml1@domain.com','eml2@domain.com']
    .PARAMETER MessageSubject
    MessageSubject string[-MessageSubject 'Subject string']
    .PARAMETER MessageId
    Corresponds to the value of the Message-Id: header field in the message. Be sure to include the full MessageId string (which may include angle brackets) and enclose the value in quotation marks 
    .PARAMETER InternalMessageId
    The InternalMessageId parameter filters the message tracking log entries by the value of the InternalMessageId field. The InternalMessageId value is a message identifier that's assigned by the Exchange server that's currently processing the message.  The value of the internal-message-id for a specific message is different in the message tracking log of every Exchange server that's involved in the delivery of the message.
    .PARAMETER NetworkMessageId
    This field contains a unique message ID value that persists across copies of the message that may be created due to bifurcation or distribution group expansion. 
    .PARAMETER EventID
    The EventId parameter filters the results by the delivery status of the message (RECEIVE|DELIVER|FAIL|SEND|RESOLVE|EXPAND|TRANSFER|DEFER). [-Status 'Failed']")]
    .PARAMETER Reference
    The Reference field contains additional information for specific types of events. For example, the Reference field value for a DSN message tracking entry contains the InternalMessageId value of the message that caused the DSN. For many types of events, the value of Reference is blank
    .PARAMETER Source
    Source (STOREDRIVER|SMTP|DNS|ROUTING)[-Source STOREDRIVER]
    .PARAMETER Server
    The Server parameter specifies the Exchange server where you want to run this command. You can use any value that uniquely identifies the server. For example: Name, FQDN, Distinguished name (DN), Exchange Legacy DN. If you don't use this parameter, the command is run on the local server[-Server Servername]
    .PARAMETER TransportTrafficType
    The TransportTrafficType parameter filters the message tracking log entries by the value of the TransportTrafficType field. However, this field isn't interesting for on-premises Exchange organizations[-TransportTrafficType xxx]
    .PARAMETER Resultsize
    The ResultSize parameter specifies the maximum number of results to return. If you want to return all requests that match the query, use unlimited for the value of this parameter. The default value is 1000.
    .PARAMETER Tag
    Tag string to be used in export filenames and output variablename[-Tag 'FromSenderX']
    .PARAMETER SimpleTrack
    switch to just return the net messages on the initial track (no Fail/Quarantine, MTDetail or other post-processing summaries) [-simpletrack]
    .PARAMETER DoExports
    switch to perform configured csv exports of results (defaults true) [-DoExports]
    .PARAMETER TenOrg
    Tenant Tag (3-letter abbrebiation)[-TenOrg 'XYZ']
    .PARAMETER Credential
    Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]
    .PARAMETER UserRole
    Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]
    .PARAMETER Silent
    Suppress echoes.
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    None. Returns no objects or output (.NET types)
    System.Boolean
    [| get-member the output to see what .NET obj TypeName is returned, to use here]
    .EXAMPLE
    PS> $pltI=[ordered]@{
    PS>     ticket="TICKETNO" ;
    PS>     Requestor="USERID";
    PS>     days=7 ;
    PS>     Start="" ;
    PS>     End="" ;
    PS>     Sender="" ;
    PS>     Recipients="" ;
    PS>     MessageSubject="" ;
    PS>     MessageId="" ;
    PS>     InternalMessageId="" ;
    PS>     Reference="" ;
    PS>     EventID='' ;
    PS>     ConnectorId="" ;
    PS>     Source="" ;
    PS>     ResultSize="" ;
    PS>     Tag='' ;
    PS> } ;
    PS> $pltGMTL = [ordered]@{} ;
    PS> $pltI.GetEnumerator() | ?{ $_.value}  | ForEach-Object { $pltGMTL.Add($_.Key, $_.Value) } ;
    PS> $vn = (@("xopMsgs$($pltI.ticket)",$pltI.Tag) | ?{$_}) -join '_' write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Get-MessageTrackingLogTDO w`n$(($pltGMTL|out-string).trim())`n(assign to `$$($vn))" ;
    PS> if(gv $vn -ea 0){rv $vn} ;
    PS> if($tmsgs = Get-MessageTrackingLogTDO @pltGMTL){sv -na $vn -va $tmsgs ;
    PS> write-host "(assigned to `$$vn)"} ;
    Demo run fed by splatted parameters with Days specified and Start/End blank (matches 7PSMsgTrkAll splat layout)
    .EXAMPLE
    PS> $platIn=[ordered]@{
    PS>     ticket="TICKETNO" ;
    PS>     Requestor="USERID";
    PS>     days=0 ;
    PS>     Start= (get-date '5/31/2024 1:01:10 PM').adddays(-1)  ;
    PS>     End= (get-date '5/31/2024 1:01:10 PM').adddays(1)  ; ;
    PS>     Sender="" ;
    PS>     Recipients="" ;
    PS>     MessageSubject="" ;
    PS>     MessageId="" ;
    PS>     InternalMessageId="" ;
    PS>     Reference="" ;
    PS>     EventID='' ;
    PS>     ConnectorId="" ;
    PS>     Source="" ;
    PS>     ResultSize="" ;
    PS>     Tag='' ;
    PS> } ;
    PS> $pltGMTL = [ordered]@{} ;
    PS> $pltI.GetEnumerator() | ?{ $_.value}  | ForEach-Object { $pltGMTL.Add($_.Key, $_.Value) } ;
    PS> $vn = (@("xopMsgs$($pltI.ticket)",$pltI.Tag) | ?{$_}) -join '_' write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Get-MessageTrackingLogTDO w`n$(($pltGMTL|out-string).trim())`n(assign to `$$($vn))" ;
    PS> if(gv $vn -ea 0){rv $vn} ;
    PS> if($tmsgs = Get-MessageTrackingLogTDO @pltGMTL){sv -na $vn -va $tmsgs ;
    PS> write-host "(assigned to `$$vn)"} ;
    Demo run fed by splatted parameters with Days set to 0 and Start/End using explicit timestamps, calculated to bracket a specific timestamp(matches 7PSMsgTrkAll splat layout)
    .EXAMPLE
    PS> gci \\tsclient\d\scripts\Get-MessageTrackingLogTDO* | copy-item -dest c:\scripts\ -Verbose
    Copy in via RDP (includes exported psbreakpoint file etc)
    .LINK
    https://bitbucket.org/tostka/powershell/
    #>
    #[CmdletBinding(DefaultParameterSetName='Days')]
    [CmdletBinding()]
    ## PSV3+ whatif support:[CmdletBinding(SupportsShouldProcess)]
    ###[Alias('Alias','Alias2')]
    PARAM(
        [Parameter(Mandatory=$True,HelpMessage="Ticket Number [-Ticket 'TICKETNO']")]
            [Alias('tix')]
            [string]$ticket,
        [Parameter(Mandatory=$False,HelpMessage="Ticket Customer email identifier. [-Requestor 'fname.lname@domain.com']")]
            [Alias('UID')]
            [string]$Requestor,
        #[Parameter(ParameterSetName='Dates',HelpMessage="Start of range to be searched[-StartDate '11/5/2021 2:16 PM']")]
        [Parameter(HelpMessage="Start of range to be searched[-StartDate '11/5/2021 2:16 PM']")]
            #[Alias('Start')]
            #[DateTime]$StartDate,
            [Alias('StartDate')]
            [DateTime]$Start,
        #[Parameter(ParameterSetName='Dates',HelpMessage="End of range to be searched (defaults to current time if unspecified)[-EndDate '11/5/2021 5:16 PM']")]
        [Parameter(HelpMessage="End of range to be searched (defaults to current time if unspecified)[-EndDate '11/5/2021 5:16 PM']")]
            [Alias('EndDate')]
            [DateTime]$End,
        #[Parameter(ParameterSetName='Days',HelpMessage="Days to be searched, back from current time(Alt to use of StartDate & EndDate; Note:MS won't search -gt 10 days)[-Days 7]")]
        [Parameter(HelpMessage="Days to be searched, back from current time(Alt to use of StartDate & EndDate; Note:MS won't search -gt 10 days)[-Days 7]")]
            #[ValidateRange(0,[int]::MaxValue)]
            [ValidateRange(0,30)] # EXOP log retn is 2g or 30d whichever comes first
            [int]$Days,
        [Parameter(HelpMessage="SenderAddress (an array runs search on each)[-SenderAddress addr@domain.com]")]
            [Alias('SenderAddress')]
            [string]$Sender,
        [Parameter(HelpMessage="RecipientAddress (an array runs search on each)[-RecipientAddress addr@domain.com]")]
                [Alias('RecipientAddress')]
                [string[]]$Recipients, # MultiValuedProperty
        [Parameter(HelpMessage="Subject of target message (emulated via post filtering, not supported param of Get-xoMessageTrace) [-Subject 'Some subject']")]
                [Alias('subject')]
                [string]$MessageSubject,
        [Parameter(Mandatory=$False,HelpMessage="Corresponds to the value of the Message-Id: header field in the message. Be sure to include the full MessageId string (which may include angle brackets) and enclose the value in quotation marks[-MessageId `"<nnnn-nnn...>`"]")]
            [string]$MessageId,
        [Parameter(Mandatory=$False,HelpMessage="The InternalMessageId parameter filters the message tracking log entries by the value of the InternalMessageId field. The InternalMessageId value is a message identifier that's assigned by the Exchange server that's currently processing the message.  The value of the internal-message-id for a specific message is different in the message tracking log of every Exchange server that's involved in the delivery of the message.")]
            [string]$InternalMessageId,
        [Parameter(Mandatory=$False,HelpMessage="This field contains a unique message ID value that persists across copies of the message that may be created due to bifurcation or distribution group expansion.(Ex16,19)")]
            [string]$NetworkMessageId,
        [Parameter(Mandatory=$False,HelpMessage="The Reference field contains additional information for specific types of events. For example, the Reference field value for a DSN message tracking entry contains the InternalMessageId value of the message that caused the DSN. For many types of events, the value of Reference is blank")]
            [string]$Reference,
        [Parameter(HelpMessage="The Status parameter filters the results by the delivery status of the message (None|GettingStatus|Failed|Pending|Delivered|Expanded|Quarantined|FilteredAsSpam),an array runs search on each). [-Status 'Failed']")]
            [Alias('DeliveryStatus','Status')]
            [ValidateSet("RECEIVE","DELIVER","FAIL","SEND","RESOLVE","EXPAND","TRANSFER","DEFER","")]
            [string[]]$EventId, # MultiValuedProperty
        [Parameter(Mandatory=$False,HelpMessage="Source (STOREDRIVER|SMTP|DNS|ROUTING)[-Source STOREDRIVER]")]
            [ValidateSet("STOREDRIVER","SMTP","DNS","ROUTING","")]
            [string]$Source,
        [Parameter(Mandatory=$False,HelpMessage="The TransportTrafficType parameter filters the message tracking log entries by the value of the TransportTrafficType field. However, this field isn't interesting for on-premises Exchange organizations[-TransportTrafficType xxx]")]
            [string]$TransportTrafficType, 
        [Parameter(Mandatory=$False,HelpMessage="Connector ID string to be post-filtered from results[-Connectorid xxx]")]
            [string]$Connectorid,
        [Parameter(Mandatory=$False,HelpMessage="The Server parameter specifies the Exchange server where you want to run this command. You can use any value that uniquely identifies the server. For example: Name, FQDN, Distinguished name (DN), Exchange Legacy DN. If you don't use this parameter, the command is run on the local server[-Server Servername]")]
            [string]$Server,
        [Parameter(Mandatory=$False,HelpMessage="The ResultSize parameter specifies the maximum number of results to return. If you want to return all requests that match the query, use unlimited for the value of this parameter. The default value is 1000.")]
            [ValidateScript({
              if( ($_ -match 'unlimited') -OR ($_.gettype().fullname -eq 'System.Int32') -OR ($null -eq $_) -OR ('' -eq $_) ){
                  $true ; 
              } else { 
                  throw "Resultsize must be an integer or the string 'unlimited' (or blank)" ; 
              } ;
            })]
            $ResultSize,
        [Parameter(HelpMessage="Tag string (Variable Name compatible: no spaces A-Za-z0-9_ only) that is used for Variables and export file name construction. [-Tag 'LastDDGSend']")] 
            [ValidatePattern('^[A-Za-z0-9_]*$')]
            [string]$Tag,
        [Parameter(HelpMessage="switch to just return the net messages on the initial track (no Fail/Quarantine, MTDetail or other post-processing summaries) [-simpletrack]")]
            [switch]$SimpleTrack,
        [Parameter(HelpMessage="switch to perform configured csv exports of results (defaults true) [-DoExports]")]
            [switch]$DoExports=$TRUE,
        # Service Connection Supporting Varis (AAD, EXO, EXOP)
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
            [ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [string]$TenOrg = $global:o365_TenOrgDefault,
        [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]")]
            [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false, HelpMessage = "Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]")]
            # sourced from get-admincred():#182: $targetRoles = 'SID', 'CSID', 'ESVC','CSVC','UID','ESvcCBA','CSvcCBA','SIDCBA' ;
            #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
            # pulling the pattern from global vari w friendly err
            [ValidateScript({
                if(-not $rgxPermittedUserRoles){$rgxPermittedUserRoles = '(SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)'} ;
                if(-not ($_ -match $rgxPermittedUserRoles)){throw "'$($_)' doesn't match `$rgxPermittedUserRoles:`n$($rgxPermittedUserRoles.tostring())" ; } ;
                return $true ;
            })]
            [string[]]$UserRole = @('SIDCBA','SID','CSVC'),
            #@('SID','CSVC'),
            # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent
    ) ;
    BEGIN{
        #region CONSTANTS_AND_ENVIRO #*======v CONSTANTS_AND_ENVIRO v======
        # Debugger:proxy automatic variables that aren't directly accessible when debugging (must be assigned and read back from another vari) ; 
        $rPSCmdlet = $PSCmdlet ; 
        $rPSScriptRoot = $PSScriptRoot ; 
        $rPSCommandPath = $PSCommandPath ; 
        $rMyInvocation = $MyInvocation ; 
        $rPSBoundParameters = $PSBoundParameters ; 
        [array]$score = @() ; 
        if($rPSCmdlet.MyInvocation.InvocationName){
            if($rPSCmdlet.MyInvocation.InvocationName -match '\.ps1$'){
                $score+= 'ExternalScript' 
            }elseif($rPSCmdlet.MyInvocation.InvocationName  -match '^\.'){
                write-warning "dot-sourced invocation detected!:$($rPSCmdlet.MyInvocation.InvocationName)`n(will be unable to leverage script path etc from MyInvocation objects)" ; 
                # dot sourcing is implicit scripot exec
                $score+= 'ExternalScript' ; 
            } else {$score+= 'Function' };
        } ; 
        if($rPSCmdlet.CommandRuntime){
            if($rPSCmdlet.CommandRuntime.tostring() -match '\.ps1$'){$score+= 'ExternalScript' } else {$score+= 'Function' }
        } ; 
        $score+= $rMyInvocation.MyCommand.commandtype.tostring() ; 
        $grpSrc = $score | group-object -NoElement | sort count ;
        if( ($grpSrc |  measure | select -expand count) -gt 1){
            write-warning  "$score mixed results:$(($grpSrc| ft -a count,name | out-string).trim())" ;
            if($grpSrc[-1].count -eq $grpSrc[-2].count){
                write-warning "Deadlocked non-majority results!" ;
            } else {
                $runSource = $grpSrc | select -last 1 | select -expand name ;
            } ;
        } else {
            write-verbose "consistent results" ;
            $runSource = $grpSrc | select -last 1 | select -expand name ;
        };
        write-verbose  "Calculated `$runSource:$($runSource)" ;
        'score','grpSrc' | get-variable | remove-variable ; # cleanup temp varis

        # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
        ${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ;
        $PSParameters = New-Object -TypeName PSObject -Property $rPSBoundParameters ;
        write-verbose "`$rPSBoundParameters:`n$(($rPSBoundParameters|out-string).trim())" ;
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        # pre psv2, no $rPSBoundParameters autovari to check, so back them out:
        if($rPSCmdlet.MyInvocation.InvocationName){
            if($rPSCmdlet.MyInvocation.InvocationName  -match '^\.'){
                $smsg = "detected dot-sourced invocation: Skipping `$PSCmdlet.MyInvocation.InvocationName-tied cmds..." ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } else { 
                write-verbose 'Collect all non-default Params (works back to psv2 w CmdletBinding)'
                $ParamsNonDefault = (Get-Command $rPSCmdlet.MyInvocation.InvocationName).parameters | Select-Object -expand keys | Where-Object{$_ -notmatch '(Verbose|Debug|ErrorAction|WarningAction|ErrorVariable|WarningVariable|OutVariable|OutBuffer)'} ;
            } ; 
        } else { 
            $smsg = "(blank `$rPSCmdlet.MyInvocation.InvocationName, skipping Parameters collection)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; 
        #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
        <#
        # Debugger:proxy automatic variables that aren't directly accessible when debugging ; 
        $rPSScriptRoot = $PSScriptRoot ; 
        $rPSCommandPath = $PSCommandPath ; 
        $rMyInvocation = $MyInvocation ; 
        $rPSBoundParameters = $PSBoundParameters ; 
        #>
        $ScriptDir = $scriptName = '' ;     
        if($ScriptDir -eq '' -AND ( (get-variable -name rPSScriptRoot -ea 0) -AND (get-variable -name rPSScriptRoot).value.length)){
            $ScriptDir = $rPSScriptRoot
        } ; # populated rPSScriptRoot
        if( (get-variable -name rPSCommandPath -ea 0) -AND (get-variable -name rPSCommandPath).value.length){
            $ScriptName = $rPSCommandPath
        } ; # populated rPSCommandPath
        if($ScriptDir -eq '' -AND $runSource -eq 'ExternalScript'){$ScriptDir = (Split-Path -Path $rMyInvocation.MyCommand.Source -Parent)} # Running from File
        # when $runSource:'Function', $rMyInvocation.MyCommand.Source is empty,but on functions also tends to pre-hit from the rPSCommandPath entFile.FullPath ;
        if( $scriptname -match '\.psm1$' -AND $runSource -eq 'Function'){
            write-host "MODULE-HOMED FUNCTION:Use `$CmdletName to reference the running function name for transcripts etc (under a .psm1 `$ScriptName will reflect the .psm1 file  fullname)"
            if(-not $CmdletName){write-warning "MODULE-HOMED FUNCTION with BLANK `$CmdletNam:$($CmdletNam)" } ;
        } # Running from .psm1 module
        if($ScriptDir -eq '' -AND (Test-Path variable:psEditor)) {
            write-verbose "Running from VSCode|VS" ; 
            $ScriptDir = (Split-Path -Path $psEditor.GetEditorContext().CurrentFile.Path -Parent) ; 
                if($ScriptName -eq ''){$ScriptName = $psEditor.GetEditorContext().CurrentFile.Path }; 
        } ;
        if ($ScriptDir -eq '' -AND $host.version.major -lt 3 -AND $rMyInvocation.MyCommand.Path.length -gt 0){
            $ScriptDir = $rMyInvocation.MyCommand.Path ; 
            write-verbose "(backrev emulating `$rPSScriptRoot, `$rPSCommandPath)"
            $ScriptName = split-path $rMyInvocation.MyCommand.Path -leaf ;
            $rPSScriptRoot = Split-Path $ScriptName -Parent ;
            $rPSCommandPath = $ScriptName ;
        } ;
        if ($ScriptDir -eq '' -AND $rMyInvocation.MyCommand.Path.length){
            if($ScriptName -eq ''){$ScriptName = $rMyInvocation.MyCommand.Path} ;
            $ScriptDir = $rPSScriptRoot = Split-Path $rMyInvocation.MyCommand.Path -Parent ;
        }
        if ($ScriptDir -eq ''){throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$rMyInvocation IS BLANK!" } ;
        if($ScriptName){
            if(-not $ScriptDir ){$ScriptDir = Split-Path -Parent $ScriptName} ; 
            $ScriptBaseName = split-path -leaf $ScriptName ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($ScriptName) ;
        } ; 
        # blank $cmdlet name comming through, patch it for Scripts:
        if(-not $CmdletName -AND $ScriptBaseName){
            $CmdletName = $ScriptBaseName
        }
        # last ditch patch the values in if you've got a $ScriptName
        if($rPSScriptRoot.Length -ne 0){}else{ 
            if($ScriptName){$rPSScriptRoot = Split-Path $ScriptName -Parent }
            else{ throw "Unpopulated, `$rPSScriptRoot, and no populated `$ScriptName from which to emulate the value!" } ; 
        } ; 
        if($rPSCommandPath.Length -ne 0){}else{ 
            if($ScriptName){$rPSCommandPath = $ScriptName }
            else{ throw "Unpopulated, `$rPSCommandPath, and no populated `$ScriptName from which to emulate the value!" } ; 
        } ; 
        if(-not ($ScriptDir -AND $ScriptBaseName -AND $ScriptNameNoExt  -AND $rPSScriptRoot  -AND $rPSCommandPath )){ 
            throw "Invalid Invocation. Blank `$ScriptDir/`$ScriptBaseName/`ScriptNameNoExt" ; 
            BREAK ; 
        } ; 
        # echo results dyn aligned:
        $tv = 'runSource','CmdletName','ScriptName','ScriptBaseName','ScriptNameNoExt','ScriptDir','PSScriptRoot','PSCommandPath','rPSScriptRoot','rPSCommandPath' ; 
        $tvmx = ($tv| Measure-Object -Maximum -Property Length).Maximum * -1 ; 
        if($silent){}else{
            #$tv | get-variable | %{  write-host -fore yellow ("`${0,$tvmx} : {1}" -f $_.name,$_.value) } ; # w-h
            $tv | get-variable | %{  write-verbose ("`${0,$tvmx} : {1}" -f $_.name,$_.value) } ; # w-v
        }
        'tv','tvmx'|get-variable | remove-variable ; # cleanup temp varis        

        #endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------

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

        $DaysLimit = 30 ; # technically no specific limit to Get-MessageTrackingLog, but practical matter they're limited to 30d on the drive
        $rgxIsPlusAddrSmtpAddr = "[+].*@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}" ; 

        #$ComputerName = $env:COMPUTERNAME ;
        #$NoProf = [bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
        # XXXMeta derived constants:
        # - AADU Licensing group checks
        # calc the rgxLicGrpName fr the existing $xxxmeta.rgxLicGrpDN: (get-variable tormeta).value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
        #$rgxLicGrpName = (get-variable -name "$($tenorg)meta").value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
        # use the dn vers LicGrouppDN = $null ; # | ?{$_ -match $tormeta.rgxLicGrpDN}
        #$rgxLicGrpDN = (get-variable -name "$($tenorg)meta").value.rgxLicGrpDN

        # email trigger vari, it will be semi-delimd list of mail-triggering events
        #$script:PassStatus = $null ;
        #[array]$SmtpAttachment = $null ;

        # local Constants:
        $prpMTLfta = 'Timestamp','EventId','Sender','Recipients','MessageSubject' ; 
        $prpXCsv = "Timestamp",@{N='TimestampLocal'; E={$_.Timestamp.ToLocalTime()}},"Source","EventId","RelatedRecipientAddress","Sender",@{N='Recipients'; E={$_.Recipients}},"RecipientCount",@{N='RecipientStatus'; E={$_.RecipientStatus}},"MessageSubject","TotalBytes",@{N='Reference'; E={$_.Reference}},"MessageLatency","MessageLatencyType","InternalMessageId","MessageId","ReturnPath","ClientIp","ClientHostname","ServerIp","ServerHostname","ConnectorId","SourceContext","MessageInfo",@{N='EventData'; E={$_.EventData}} ;
        $prpMTFailFL = 'Timestamp','ClientHostname','Source','EventId','Recipients','RecipientStatus','MessageSubject','ReturnPath' ;
        $s24HTimestamp = 'yyyyMMdd-HHmm'
        $sFiletimestamp =  $s24HTimestamp

        #endregion CONSTANTS_AND_ENVIRO ; #*------^ END CONSTANTS_AND_ENVIRO ^------
        
        #region FUNCTIONS ; #*======v FUNCTIONS v======
        # Pull the CUser mod dir out of psmodpaths:
        #$CUModPath = $env:psmodulepath.split(';')|?{$_ -like '*\Users\*'} ;

        #region RVARIINVALIDCHARS ; #*------v RVARIINVALIDCHARS v------
        #*------v Function Remove-InvalidVariableNameChars v------
        if(-not (gcm Remove-InvalidVariableNameChars -ea 0)){
            Function Remove-InvalidVariableNameChars ([string]$Name) {
                ($Name.tochararray() -match '[A-Za-z0-9_]') -join '' | write-output ;
            };
        } ;
        #*------^ END Function Remove-InvalidVariableNameChars ^------
        #endregion RVARIINVALIDCHARS ; #*------^ END RVARIINVALIDCHARS ^------

        #*------v Function Connect-ExchangeServerTDO v------
        if(-not(get-command Connect-ExchangeServerTDO -ea 0)){
            Function Connect-ExchangeServerTDO {
                <#
                .SYNOPSIS
                Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
                will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellREmote (REMS) connect to each server, 
                stopping at the first successful connection.
                .NOTES
                Version     : 3.0.3
                Author      : Todd Kadrie
                Website     : http://www.toddomation.com
                Twitter     : @tostka / http://twitter.com/tostka
                CreatedDate : 2024-05-30
                FileName    : Connect-ExchangeServerTDO.ps1
                License     : (none-asserted)
                Copyright   : (none-asserted)
                Github      : https://github.com/tostka/verb-Ex2010
                Tags        : Powershell, ActiveDirectory, Exchange, Discovery
                AddedCredit : Brian Farnsworth
                AddedWebsite: https://codeandkeep.com/
                AddedTwitter: URL
                AddedCredit : David Paulson
                AddedWebsite: https://techcommunity.microsoft.com/t5/exchange-team-blog/exchange-health-checker-has-a-new-home/ba-p/2306671
                AddedTwitter: URL
                REVISIONS
                * 3:54 PM 11/26.2.34 integrated back TLS fixes, and ExVersNum flip from June; syncd dbg & vx10 copies.
                * 12:57 PM 6/11/2024 Validated, Ex2010 & Ex2019, hub, mail & edge roles: tested ☑️ on CMW mail role (Curly); and Jumpbox; 
                    copied in CBH from repo copy, which has been updated/debugged compat on CMW Edge 
                    includes local snapin detect & load for edge role (simplest EMS load option for Edge role, from David Paulson's original code; no longer published with Ex2010 compat)
                * 1:30 PM 9/5/2024 added  update-SecurityProtocolTDO() SB to begin
                * 12:49 PM 6/21/2024 flipped PSS Name to Exchange$($ExchVers[dd])
                * 11:28 AM 5/30/2024 fixed failure to recognize existing functional PSSession; Made substantial update in logic, validate works fine with other orgs, and in our local orgs.
                * 4:02 PM 8/28/2023 debuged, updated CBH, renamed connect-ExchangeSErver -> Connect-ExchangeServerTDO (avoid name clashes, pretty common verb-noun combo).
                * 12:36 PM 8/24/2023 init

                .DESCRIPTION
                Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
                will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellRemote (REMS) connect to each server, 
                stopping at the first successful connection.

                Relies upon/requires get-ADExchangeServerTDO(), to return a descriptive summary of the Exchange server(s) revision etc, for connectivity logic.
                Supports Exchange 2010 through 2019, as implemented.
        
                Intent, as contrasted with verb-EXOP/Ex2010 is to have no local module dependancies, when running EXOP into other connected orgs, where syncing profile & supporting modules code can be problematic. 
                This uses native ADSI calls, which are supported by Windows itself, without need for external ActiveDirectory module etc.

                The particular approach inspired by BF's demo func that accompanied his take on get-adExchangeServer(), which I hybrided with my own existing code for cred-less connectivity. 
                I added get-OrganizationConfig testing, for connection pre/post confirmation, along with Exchange Server revision code for continutional handling of new-pssession remote powershell EMS connections.
                Also shifted connection code into _connect-EXOP() internal func.
                As this doesn't rely on local module presence, it doesn't have to do the usual local remote/local invocation detection you'd do for non-dehydrated on-server EMS (more consistent this way, anyway; 
                there are only a few cmdlet outputs I'm aware of, that have fundementally broken returns dehydrated, and require local non-remote EMS use to function.

                My core usage would be to paste the function into the BEGIN{} block for a given remote org process, to function as a stricly local ad-hoc function.
                .PARAMETER name
                FQDN of a specific Exchange server[-Name EXSERVER.DOMAIN.COM]
                .PARAMETER discover
                Boolean paraameter that drives auto-discovery of target Exchange servers for connection (defaults `$true)[-discover:`$false]
                .PARAMETER credential
                Use specific Credentials[-Credentials [credential object]
                    .PARAMETER Site
                Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-Site 'SITENAME']
                .PARAMETER RoleNames
                Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']
                .PARAMETER TenOrg
                Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
                .INPUTS
                None. Does not accepted piped input.(.NET types, can add description)
                .OUTPUTS
                [system.object] Returns a system object containing a successful PSSession
                System.Boolean
                [| get-member the output to see what .NET obj TypeName is returned, to use here]
                System.Array of System.Object's
                .EXAMPLE
                PS> $PSSession = Connect-ExchangeServerTDO -siteName SITENAME -RoleNames @('HUB','CAS') -verbose 
                Demo's connecting to a functional Hub or CAS server in the SITENAME site with verbose outputs, the `PSSession variable will contain information about the successful connection. Makes automatic Exchangeserver discovery calls into AD (using ADSI) leveraging the separate get-ADExchangeServerTDO()
                .EXAMPLE
                PS> TRY{$Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name}CATCH{$Site=$env:COMPUTERNAME} ;
                PS> $PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
                Demo including support for EdgeRole, which is detected on it's lack of AD Site specification (which gets fed through to call, by setting the Site to the machine itself).
                .LINK
                https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/
                .LINK
                https://github.com/Lucifer1993/PLtools/blob/main/HealthChecker.ps1
                .LINK
                https://microsoft.github.io/CSS-Exchange/Diagnostics/HealthChecker/
                .LINK
                https://bitbucket.org/tostka/powershell/
                .LINK
                https://github.com/tostka/verb-Ex2010
                #>        
                [CmdletBinding(DefaultParameterSetName='discover')]
                PARAM(
                    [Parameter(Position=0,Mandatory=$true,ParameterSetName='name',HelpMessage="FQDN of a specific Exchange server[-Name EXSERVER.DOMAIN.COM]")]
                        [String]$name,
                    [Parameter(Position=0,ParameterSetName='discover',HelpMessage="Boolean paraameter that drives auto-discovery of target Exchange servers for connection (defaults `$true)[-discover:`$false]")]
                        [bool]$discover=$true,
                    [Parameter(Position=1,HelpMessage = "Use specific Credentials[-Credentials [credential object]")]
                        [Management.Automation.PSCredential]$credential,
                    [Parameter(Position=1,HelpMessage="Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-Site 'SITENAME']")]
                        [Alias('Site')]
                        [string]$SiteName = (gv -name "$($TenOrg)_ADSiteDefaultName" -ea 0).Value,
                    [Parameter(Position=2,HelpMessage="Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']")]
                        [ValidateSet('MBX','CAS','HUB','UM','MBX','EDGE')]
                        [string[]]$RoleNames = @('HUB','CAS'),
                    [Parameter(Mandatory=$FALSE,HelpMessage="Tenant Tag (3-letter abbrebiation - defaults to global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']")]
                        [ValidateNotNullOrEmpty()]
                        [string]$TenOrg = $global:o365_TenOrgDefault
                ) ;
                BEGIN{
                    $Verbose = ($VerbosePreference -eq 'Continue') ;
                    $CurrentVersionTlsLabel = [Net.ServicePointManager]::SecurityProtocol ; # Tls, Tls11, Tls12 ('Tls' == TLS1.0)  ;
			        write-verbose "PRE: `$CurrentVersionTlsLabel : $($CurrentVersionTlsLabel )" ;
			        # psv6+ already covers, test via the SslProtocol parameter presense
			        if ('SslProtocol' -notin (Get-Command Invoke-RestMethod).Parameters.Keys) {
				        $currentMaxTlsValue = [Math]::Max([Net.ServicePointManager]::SecurityProtocol.value__,[Net.SecurityProtocolType]::Tls.value__) ;
				        write-verbose "`$currentMaxTlsValue : $($currentMaxTlsValue )" ;
				        $newerTlsTypeEnums = [enum]::GetValues('Net.SecurityProtocolType') | Where-Object { $_ -gt $currentMaxTlsValue }
				        if($newerTlsTypeEnums){
					        write-verbose "Appending upgraded/missing TLS `$enums:`n$(($newerTlsTypeEnums -join ','|out-string).trim())" ;
				        } else {
					        write-verbose "Current TLS `$enums are up to date with max rev available on this machine" ;
				        };
				        $newerTlsTypeEnums | ForEach-Object {
					        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $_
				        } ;
			        } ;
                    $smsg = "#*------v Function _connect-ExOP v------" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    function _connect-ExOP{
                        [CmdletBinding()]
                        PARAM(
                            [Parameter(Position=0,Mandatory=$true,HelpMessage="Exchange server AD Summary system object[-Server EXSERVER.DOMAIN.COM]")]
                                [system.object]$Server,
                            [Parameter(Position=1,HelpMessage = "Use specific Credentials[-Credentials [credential object]")]
                                [Management.Automation.PSCredential]$credential
                        );
                        $verbose = $($VerbosePreference -eq "Continue") ;
                        if([double]$ExVersNum = [regex]::match($Server.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
                            switch -regex ([string]$ExVersNum) {
                                '15.2' { $isEx2019 = $true ; $ExVers = 'Ex2019' }
                                '15.1' { $isEx2016 = $true ; $ExVers = 'Ex2016'}
                                '15.0' { $isEx2013 = $true ; $ExVers = 'Ex2013'}
                                '14.*' { $isEx2010 = $true ; $ExVers = 'Ex2010'}
                                '8.*' { $isEx2007 = $true ; $ExVers = 'Ex2007'}
                                '6.5' { $isEx2003 = $true ; $ExVers = 'Ex2003'}
                                '6' {$isEx2000 = $true ; $ExVers = 'Ex2000'} ;
                                default {
                                    $smsg = "UNRECOGNIZED ExVersNum.Major.Minor string:$($ExVersNum)! ABORTING!" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    THROW $SMSG ;
                                    BREAK ;
                                }
                            } ;
                        }else {
                            $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$Server.version:$($Server.version)!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            throw $smsg ;
                            break ;
                        } ;
                        if($Server.RoleNames -eq 'EDGE'){
                            if(($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup')) -or
                                ($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup')) -or
                                $ByPassLocalExchangeServerTest)
                            {
                                if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole') -or
                                        (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'))
                                {
                                    $smsg = "We are on Exchange Edge Transport Server"
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                    $IsEdgeTransport = $true
                                }
                                TRY {
                                    Get-ExchangeServer -ErrorAction Stop | Out-Null
                                    $smsg = "Exchange PowerShell Module already loaded."
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                    $passed = $true 
                                }CATCH {
                                    $smsg = "Failed to run Get-ExchangeServer"
                                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                    if($isLocalExchangeServer){
                                        write-host  "Loading Exchange PowerShell Module..."
                                        TRY{
                                            if($IsEdgeTransport){
                                                # implement local snapins access on edge role: Only way to get access to EMS commands.
                                                [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exshell.psc1" -ErrorAction Stop
                                                ForEach($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn){
                                                    write-verbose ("Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name)
                                                    Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop
                                                } ; 
                                                Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop ; 
                                                $passed = $true #We are just going to assume this passed.
                                            }else{
                                                Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                                                Connect-ExchangeServer -Auto -ClientApplication:ManagementShell
                                                $passed = $true #We are just going to assume this passed.
                                            } 
                                        }CATCH {
                                            $smsg = "Failed to Load Exchange PowerShell Module..." ; 
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                        }                               
                                    } ;
                                } FINALLY {
                                    if($LoadExchangeVariables -and $passed -and $isLocalExchangeServer){
                                        if($ExInstall -eq $null -or $ExBin -eq $null){
                                            if(Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup'){
                                                $Global:ExInstall = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath
                                            }else{
                                                $Global:ExInstall = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath
                                            }
        
                                            $Global:ExBin = $Global:ExInstall + "\Bin"
        
                                            $smsg = ("Set ExInstall: {0}" -f $Global:ExInstall)
                                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                            $smsg = ("Set ExBin: {0}" -f $Global:ExBin)
                                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                        } ; 
                                    } ; 
                                } ; 
                            } else  {
                                $smsg = "Does not appear to be an Exchange 2010 or newer server." ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                            }
                            if(get-command -Name Get-OrganizationConfig -ea 0){
                                $smsg = "Running in connected/Native EMS" ; 
                                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                Return $true ; 
                            } else { 
                                TRY{
                                    $smsg = "Initiating Edge EMS local session (exshell.psc1 & exchange.ps1)" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                    # 5;36 PM 5/30/2024 didn't work, went off to nowhere for a long time, and exited the script
                                    #& (gcm powershell.exe).path -PSConsoleFile "$($env:ExchangeInstallPath)bin\exshell.psc1" -noexit -command ". '$($env:ExchangeInstallPath)bin\Exchange.ps1'"
                                    <# [Adding the Transport Server to Exchange - Mark Lewis Blog](https://marklewis.blog/2020/11/19/adding-the-transport-server-to-exchange/)
                                    To access the management console on the transport server, I opened PowerShell then ran
                                    exshell.psc1
                                    Followed by
                                    exchange.ps1
                                    At this point, I was able to create a new subscription using he following PowerShel
                                    #>
                                    invoke-command exshell.psc1 ; 
                                    invoke-command exchange.ps1
                                    if(get-command -Name Get-OrganizationConfig -ea 0){
                                        $smsg = "Running in connected/Native EMS" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                        Return $true ;
                                    } else { return $false };  
                                }CATCH{
                                    Write-Error $_ ;
                                } ;
                            } ; 
                        } else {
                            $pltNPSS=@{ConnectionURI="http://$($Server.FQDN)/powershell"; ConfigurationName='Microsoft.Exchange' ; name="Exchange$($ExVersNum.tostring())"} ;
                            # use ExVersUnm dd instead of hardcoded (Exchange2010)
                            if($ExVersNum -ge 15){
                                $smsg = "EXOP.15+:Adding -Authentication Kerberos" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                $pltNPSS.add('Authentication',"Kerberos") ;
                                $pltNPSS.name = $ExVers ;
                            } ;
                            $smsg = "Adding EMS (connecting to $($Server.FQDN))..." ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $smsg = "New-PSSession w`n$(($pltNPSS|out-string).trim())" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $ExPSS = New-PSSession @pltNPSS  ;
                            $ExIPSS = Import-PSSession $ExPSS -allowclobber ;
                            $ExPSS | write-output ;
                            $ExPSS= $ExIPSS = $null ;
                        } ; 
                    } ;
                    $smsg = "#*------^ END Function _connect-ExOP ^------" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $pltGADX=@{
                        ErrorAction='Stop';
                    } ;
                } ;
                PROCESS{
                    if($PSBoundParameters.ContainsKey('credential')){
                        $pltGADX.Add('credential',$credential) ;
                    }
                    if($SiteName){
                        $pltGADX.Add('siteName',$siteName) ;
                    } ;
                    if($RoleNames){
                        $pltGADX.Add('RoleNames',$RoleNames) ;
                    } ;
                    TRY{
                        if($discover){
                            $smsg = "Getting list of Exchange Servers" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $exchServers=get-ADExchangeServerTDO @pltGADX | sort responsetime ;
                        }else{
                            $exchServers=get-ADExchangeServerTDO @pltGADX | sort responsetime ;
                        } ;
                        $pltTW=@{
                            'ErrorAction'='Stop';
                        } ;
                        $pltCXOP = @{
                            verbose = $($VerbosePreference -eq "Continue") ;
                        } ;
                        if($pltGADX.credential){
                            $pltCXOP.Add('Credential',$pltCXOP.Credential) ;
                        } ;
                        $prpPSS = 'Id','Name','ComputerName','ComputerType','State','ConfigurationName','Availability' ; 
                        foreach($exServer in $exchServers){
                            $smsg = "testing conn to:$($exServer.name.tostring())..." ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig' -ea SilentlyContinue){
                                if($pssEXOP = Get-PSSession |  where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ( $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND ( $_.runspace.ConnectionInfo.Port -eq '80') }){
                                    if($pssEXOP.State -ne "Opened" -OR $pssEXOP.Availability -ne "Available"){
                                        $pssEXOP | remove-pssession ; $pssEXOP = $null ;
                                    } ;
                                } ; 
                            } else {
                                $smsg = "(mangled ExOP conn: disconnect/reconnect...)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                if($pssEXOP = Get-PSSession |  where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ( $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND ( $_.runspace.ConnectionInfo.Port -eq '80') }){
                                    if($pssEXOP.State -ne "Opened" -OR $pssEXOP.Availability -ne "Available"){
                                        $pssEXOP | remove-pssession ; $pssEXOP = $null ;
                                    } ;
                                } ; 
                            } ;
                            if(-not $pssEXOP){
                                $smsg = "Connecting to: $($exServer.FQDN)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                if($NoTest){
                                    $ExPSS =$ExPSS = _connect-ExOP @pltCXOP -Server $exServer
                                } else {
                                    TRY{
                                        $smsg = "Testing Connection: $($exServer.FQDN)" ;
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                        If(test-connection $exServer.FQDN -count 1 -ea 0) {
                                            $smsg = "confirmed pingable..." ;
                                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                        } else {
                                            $smsg = "Unable to Ping $($exServer.FQDN)" ; ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        } ;
                                        $smsg = "Testing WinRm: $($exServer.FQDN)" ;
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                        $winrm=Test-WSMan @pltTW -ComputerName $exServer.FQDN ;
                                        if($winrm){
                                            $ExPSS = _connect-ExOP @pltCXOP -Server $exServer;
                                        } else {
                                            $smsg = "Unable to Test-WSMan $($exServer.FQDN) (skipping)" ; ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        } ;
                                    }CATCH{
                                        $errMsg="Server: $($exServer.FQDN)] $($_.Exception.Message)" ;
                                        Write-Error -Message $errMsg ;
                                        continue ;
                                    } ;
                                };
                            } else {
                                $smsg = "$((get-date).ToString('HH:mm:ss')):Accepting first valid connection w`n$(($pssEXOP | ft -a $prpPSS|out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $ExPSS = $pssEXOP ; 
                                break ; 
                            }  ;
                        } ;
                    }CATCH{
                        Write-Error $_ ;
                    } ;
                } ;
                END{
                    if(-not $ExPSS){
                        $smsg = "NO SUCCESSFUL CONNECTION WAS MADE, WITH THE SPECIFIED INPUTS!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $smsg = "(returning `$false to the pipeline...)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        return $false
                    } else{
                        if($ExPSS.State -eq "Opened" -AND $ExPSS.Availability -eq "Available"){
                            if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                                $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ;
                                throw $smsg ;
                                $smsg | write-warning  ;
                            } else {
                                $smsg = "(connected to EXOP.Org:$($orgName))" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                            return $ExPSS
                        } ;
                    } ; 
                } ;
            } ;
        } ; 
        #*------^ END Function Connect-ExchangeServerTDO ^------

        #*------v Function get-ADExchangeServerTDO v------
        if(-not(get-command get-ADExchangeServerTDO -ea 0)){
            Function get-ADExchangeServerTDO {
                <#
                .SYNOPSIS
                get-ADExchangeServerTDO.ps1 - Returns Exchangeserver summary(s) from AD records
                .NOTES
                Version     : 3.0.1
                Author      : Todd Kadrie
                Website     : http://www.toddomation.com
                Twitter     : @tostka / http://twitter.com/tostka
                CreatedDate : 2015-09-03
                FileName    : get-ADExchangeServerTDO.ps1
                License     : (none-asserted)
                Copyright   : (none-asserted)
                Github      : https://github.com/tostka/verb-Ex2010
                Tags        : Powershell, ActiveDirectory, Exchange, Discovery
                AddedCredit : Mike Pfeiffer
                AddedWebsite: mikepfeiffer.net
                AddedTwitter: URL
                AddedCredit : Sammy Krosoft 
                AddedWebsite: http://aka.ms/sammy
                AddedTwitter: URL
                AddedCredit : Brian Farnsworth
                AddedWebsite: https://codeandkeep.com/
                AddedTwitter: URL
                REVISIONS
                * 3:57 PM 11/26.2.34 updated simple write-host,write-verbose with full pswlt support;  syncd dbg & vx10 copies.
                * 12:57 PM 6/11/2024 Validated, Ex2010 & Ex2019, hub, mail & edge roles: tested ☑️ on CMW mail role (Curly); and Jumpbox; copied in CBH from repo copy, which has been updated/debugged compat on CMW Edge 
                * 2:05 PM 8/28/2023 REN -> Get-ExchangeServerInSite -> get-ADExchangeServerTDO (aliased orig); to better steer profile-level options - including in cmw org, added -TenOrg, and default Site to constructed vari, targeting new profile $XXX_ADSiteDefault vari; Defaulted -Roles to HUB,CAS as well.
                * 3:42 PM 8/24/2023 spliced together combo of my long-standing, and some of the interesting ideas BF's version had. Functional prod:
                    - completely removed ActiveDirectory module dependancies from BF's code, and reimplemented in raw ADSI calls. Makes it fully portable, even into areas like Edge DMZ roles, where ADMS would never be installed.

                * 3:17 PM 8/23/2023 post Edge testing: some logic fixes; add: -Names param to filter on server names; -Site & supporting code, to permit lookup against sites *not* local to the local machine (and bypass lookup on the local machine) ; 
                    ren $Ex10siteDN -> $ExOPsiteDN; ren $Ex10configNC -> $ExopconfigNC
                * 1:03 PM 8/22/2023 minor cleanup
                * 10:31 AM 4/7/2023 added CBH expl of postfilter/sorting to draw predictable pattern 
                * 4:36 PM 4/6.2.33 validated Psv51 & Psv20 and Ex10 & 16; added -Roles & -RoleNames params, to perform role filtering within the function (rather than as an external post-filter step). 
                For backward-compat retain historical output field 'Roles' as the msexchcurrentserverroles summary integer; 
                use RoleNames as the text role array; 
                    updated for psv2 compat: flipped hash key lookups into properties, found capizliation differences, (psv2 2was all lower case, wouldn't match); 
                flipped the [pscustomobject] with new... psobj, still psv2 doesn't index the hash keys ; updated for Ex13+: Added  16  "UM"; 20  "CAS, UM"; 54  "MBX" Ex13+ ; 16385 "CAS" Ex13+ ; 16439 "CAS, HUB, MBX" Ex13+
                Also hybrided in some good ideas from SammyKrosoft's Get-SKExchangeServers.psm1 
                (emits Version, Site, low lvl Roles # array, and an array of Roles, for post-filtering); 
                # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
                * 12:08 PM 5/15/2020 fixed vpn issue: Try/Catch'd around recently failing $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName qry
                * 11:22 AM 3/13/2020 Get-ExchangeServerInSite added a ping-test, to only return matches that are pingable, added -NoPing param, to permit (faster) untested bypass
                * 6:59 PM 1/15/2020 cleanup
                # 10:03 AM 11/16/2018 Get-ExchangeServerInSite:can't do AD-related functions when not AD authentictaed (home, pre-vpn connect). Added if/then test on status and abort balance when false.
                * 11/18/18 BF's posted rev
                # 12:10 PM 8/1/2017 updated example code at bottom, to accommodate variant sites
                # 11:28 AM 3/31/2016 validated that latest round of updates are still functional
                #1:58 PM 9/3/2015 - added pshelp and some docs
                #April 12, 2010 - web version
                .DESCRIPTION
                get-ADExchangeServerTDO.ps1 - Returns Exchangeserver summary(s) from AD records

                Hybrided together ideas from Brian Farnsworth's blog post
                [PowerShell - ActiveDirectory and Exchange Servers – CodeAndKeep.Com – Code and keep calm...](https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/)
                ... with much older concepts from  Sammy Krosoft, and much earlier Mike Pfeiffer. 

                - Subbed in MP's use of ADSI for ActiveDirectory Ps mod cmds - it's much more dependancy-free; doesn't require explicit install of the AD ps module
                ADSI support is built into windows.
                - spliced over my addition of Roles, RoleNames, Name & NoTest params, for prefiltering and suppressing testing.


                [briansworth · GitHub](https://github.com/briansworth)

                Uses an ADSI DirectorySearcher to search the current Active Directory site for Exchange on-prem servers.
                        Intent is to discover connection points for Powershell, wo the need to preload/pre-connect to Exchange.

                        But, as a non-Exchange-Management-Shell-dependant info source on Exchange Server configs, it can be used before connection, with solely AD-available data, to check configuration spes on the subject server(s). 

                        For example, this query will return sufficient data under Version to indicate which revision of Exchange is in use:


                        Returned object (in array):
                        Site      : {ADSITENAME}
                        Roles     : {64}
                        Version   : {Version 15.1 (Build 32375.7)}
                        Name      : SERVERNAME
                        RoleNames : EDGE
                        FQDN      : SERVERNAME.DOMAIN.TLD

                        ... includes the post-filterable Role property ($_.Role -contains 'CAS') which reflects the following
                        installed-roles ('msExchCurrentServerRoles') on the discovered servers
                            2   {"MBX"} # Ex10
                            4   {"CAS"}
                            16  {"UM"}
                            20  {"CAS, UM" -split ","} # 
                            32  {"HUB"}
                            36  {"CAS, HUB" -split ","}
                            38  {"CAS, HUB, MBX" -split ","}
                            54  {"MBX"} # Ex13+
                            64  {"EDGE"}
                            16385   {"CAS"} # Ex13+
                            16439   {"CAS, HUB, MBX"  -split ","} # Ex13+
                .PARAMETER Roles
                Array of msExchCurrentServerRoles 'role' integers to be filtered against (2|4|16|20|32|36|38|54|64|16385|16439)[-Roles @(38,16385)]
                .PARAMETER RoleNames
                Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']
                .PARAMETER Server
                Array of Server name strings to be filtered against[-Server 'SERVER1','SERVER2']
                .PARAMETER SiteName
                Name of specific AD SiteName to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-SiteName 'SITENAME']
                .PARAMETER TenOrg
                Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
                .PARAMETER NoPing
                Switch to suppress default 'pingable' test (e.g. returns all matches, no testing)[-NoPing]
                .INPUTS
                None. Does not accepted piped input.(.NET types, can add description)
                .OUTPUTS
                None. Returns no objects or output (.NET types)
                System.Boolean
                [| get-member the output to see what .NET obj TypeName is returned, to use here]
                System.Array of System.Object's
                .EXAMPLE
                PS> If(!($ExchangeServer)){$ExchangeServer = (get-ADExchangeServerTDO| ?{$_.RoleNames -contains 'CAS' -OR $_.RoleNames -contains 'HUB' -AND ($_.FQDN -match "^SITECODE") } | Get-Random ).FQDN
                Return a random Hub Cas Role server in the local Site with a fqdn beginning SITECODE
                .EXAMPLE
                PS> $localADExchserver = get-ADExchangeServerTDO -Names $env:computername -SiteName ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().name)
                Demo, if run from an Exchange server, return summary details about the local server (-SiteName isn't required, is default imputed from local server's Site, but demos explicit spec for remote sites)
                .EXAMPLE
                PS> $regex = '(' + [regex]($ADSiteCodeUK,$ADSiteCodeAU -join "|") + ')'
                PS> switch -regex ($($env:computername).substring(0,3)){
                PS>    "$($ADSiteCodeUS)" {$tExRole=36 } ;
                PS>    "$($regex)" {$tExRole= 38 }  default { write-error "$((get-date).ToString('HH:mm:ss')):UNRECOGNIZED SERVER PREFIX!."; } ;
                PS> } ;
                PS> $exhubcas = (get-ADExchangeServerTDO |?{($_.roles -eq $tExRole) -AND ($_.FQDN -match "$($env:computername.substring(0,3)).*")} | Get-Random ).FQDN ;
                Use a switch block to select different role combo targets for a given server fqdn prefix string.
                .EXAMPLE
                PS> $ExchangeServer = get-ADExchangeServerTDO | ?{$_.Roles -match '(4|20|32|36|38|16385|16439)'} | select -expand fqdn | get-random ; 
                Another/Older approach filtering on the Roles integer (targeting combos with Hub or CAS in the mix)
                .EXAMPLE
                PS> $ret = get-ADExchangeServerTDO -Roles @(4,20,32,36,38,16385,16439) -verbose 
                Demo use of the -Roles param, feeding it an array of Role integer values to be filtered against. In this case, the Role integers that include a CAS or HUB role.
                .EXAMPLE
                PS> $ret = get-ADExchangeServerTDO -RoleNames 'HUB','CAS' -verbose ;
                Demo use of the -RoleNames param, feeding it the array 'HUB','CAS' Role name strings to be filtered against
                PS> $ret = get-ADExchangeServerTDO -Names 'SERVERName' -verbose ;
                Demo use of the -RoleNames param, feeding it the array 'HUB','CAS' Role name strings to be filtered against
                .EXAMPLE
                PS> $ExchangeServer = get-ADExchangeServerTDO | sort version,roles,name | ?{$_.rolenames -contains 'CAS'}  | select -last 1 | select -expand fqdn ;
                Demo post sorting & filtering, to deliver a rule-based predictable pattern for server selection: 
                Above will always pick the highest Version, 'CAS' RoleName containing, alphabetically last server name (that is pingable). 
                And should stick to that pattern, until the servers installed change, when it will shift to the next predictable box.
                .EXAMPLE
                PS> $ExOPServer = get-ADExchangeServerTDO -Name LYNMS650 -SiteName Lyndale
                PS> if([double]$ExVersNum = [regex]::match($ExOPServer.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
                PS>     switch -regex ([string]$ExVersNum) {
                PS>         '15\.2' { $isEx2019 = $true ; $ExVers = 'Ex2019' }
                PS>         '15\.1' { $isEx2016 = $true ; $ExVers = 'Ex2016'}
                PS>         '15\.0' { $isEx2013 = $true ; $ExVers = 'Ex2013'}
                PS>         '14\..*' { $isEx2010 = $true ; $ExVers = 'Ex2010'}
                PS>         '8\..*' { $isEx2007 = $true ; $ExVers = 'Ex2007'}
                PS>         '6\.5' { $isEx2003 = $true ; $ExVers = 'Ex2003'}
                PS>         '6|6\.0' {$isEx2000 = $true ; $ExVers = 'Ex2000'} ;
                PS>         default {
                PS>             $smsg = "UNRECOGNIZED ExchangeServer.AdminDisplayVersion.Major.Minor string:$($ExOPServer.version)! ABORTING!" ;
                PS>             write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                PS>         }
                PS>     } ; 
                PS> }else {
                PS>     $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$ExOPServer.version:$($ExOPServer.version)!" ; 
                PS>     write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ; 
                PS>     throw $smsg ; 
                PS>     break ; 
                PS> } ; 
                Demo of parsing the returned Version property, into the proper Exchange Server revision.      
                .LINK
                https://github.com/tostka/verb-XXX
                .LINK
                https://bitbucket.org/tostka/powershell/
                .LINK
                http://mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
                .LINK
                https://github.com/SammyKrosoft/Search-AD-Using-Plain-PowerShell/blob/master/Get-SKExchangeServers.psm1
                .LINK
                https://github.com/tostka/verb-Ex2010
                .LINK
                https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/
                #>
                [CmdletBinding()]
                [Alias('Get-ExchangeServerInSite')]
                PARAM(
                    [Parameter(Position=0,HelpMessage="Array of Server name strings to be filtered against[-Server 'SERVER1','SERVER2']")]
                        [string[]]$Server,
                    [Parameter(Position=1,HelpMessage="Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-SiteName 'SITENAME']")]
                        [Alias('Site')]
                        [string]$SiteName = (gv -name "$($TenOrg)_ADSiteDefaultName" -ea 0).Value,
                    [Parameter(Position=2,HelpMessage="Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']")]
                        [ValidateSet('MBX','CAS','HUB','UM','MBX','EDGE')]
                        [string[]]$RoleNames = @('HUB','CAS'),
                    [Parameter(HelpMessage="Array of msExchCurrentServerRoles 'role' integers to be filtered against (2|4|16|20|32|36|38|54|64|16385|16439)[-Roles @(38,16385)]")]
                        [ValidateSet(2,4,16,20,32,36,38,54,64,16385,16439)]
                        [int[]]$Roles,
                    [Parameter(HelpMessage="Switch to suppress default 'pingable' test (e.g. returns all matches, no testing)[-NoTest]")]
                        [Alias('NoPing')]
                        [switch]$NoTest,
                    [Parameter(HelpMessage="Milliseconds of max timeout to wait during port 80 test (defaults 100)[-SpeedThreshold 500]")]
                        [int]$SpeedThreshold=100,
                    [Parameter(Mandatory=$FALSE,HelpMessage="Tenant Tag (3-letter abbrebiation - defaults to global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']")]
                        [ValidateNotNullOrEmpty()]
                        [string]$TenOrg = $global:o365_TenOrgDefault,
                    [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials[-Credentials [credential object]]")]
                        [System.Management.Automation.PSCredential]$Credential
                ) ;
                BEGIN{
                    ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
                    $Verbose = ($VerbosePreference -eq 'Continue') ;
                    $_sBnr="#*======v $(${CmdletName}): v======" ;
                    $smsg = $_sBnr ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
                PROCESS{
                    TRY{
                        $configNC = ([ADSI]"LDAP://RootDse").configurationNamingContext ;
                        $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC") ;
                        $bLocalEdge = $false ; 
                        if($Sitename -eq $env:COMPUTERNAME){
                            $smsg = "`$SiteName -eq `$env:COMPUTERNAME:$($SiteName):$($env:COMPUTERNAME)" ; 
                            $smsg += "`nThis computer appears to be an EdgeRole system (non-ADConnected)" ; 
                            $smsg += "`n(Blanking `$sitename and continuing discovery)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            #$bLocalEdge = $true ; 
                            $SiteName = $null ; 
                    
                        } ; 
                        If($siteName){
                            $smsg = "Getting Site: $siteName" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $objectClass = "objectClass=site" ;
                            $objectName = "name=$siteName" ;
                            $search.Filter = "(&($objectClass)($objectName))" ;
                            $site = ($search.Findall()) ;
                            $siteDN = ($site | select -expand properties).distinguishedname  ;
                        } else {
                            $smsg = "(No -Site specified, resolving site from local machine domain-connection...)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                            else{ write-host -foregroundcolor green "$($smsg)" } ;
                            TRY{$siteDN = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().GetDirectoryEntry().distinguishedName}
                            CATCH [System.Management.Automation.MethodInvocationException]{
                                $ErrTrapd=$Error[0] ;
                                if(($ErrTrapd.Exception -match 'The computer is not in a site.') -AND $env:ExchangeInstallPath){
                                    $smsg = "$($env:computername) is non-ADdomain-connected" ;
                                    $smsg += "`nand has `$env:ExchangeInstalled populated: Likely Edge Server" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                                    else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    $vers = (get-item "$($env:ExchangeInstallPath)\Bin\Setup.exe").VersionInfo.FileVersionRaw ; 
                                    $props = @{
                                        Name=$env:computername;
                                        FQDN = ([System.Net.Dns]::gethostentry($env:computername)).hostname;
                                        Version = "Version $($vers.major).$($vers.minor) (Build $($vers.Build).$($vers.Revision))" ; 
                                        #"$($vers.major).$($vers.minor)" ; 
                                        #$exServer.serialNumber[0];
                                        Roles = [System.Object[]]64 ;
                                        RoleNames = @('EDGE');
                                        DistinguishedName =  "CN=$($env:computername),CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=First Organization,CN=Microsoft Exchange,CN=Services,CN=Configuration,CN={nnnnnnnn-FAKE-GUID-nnnn-nnnnnnnnnnnn}" ;
                                        Site = [System.Object[]]'NOSITE'
                                        ResponseTime = if($rsp){$rsp.ResponseTime} else { 0} ;
                                        NOTE = "This summary object, returned for a non-AD-connected EDGE server, *approximates* what would be returned on an AD-connected server" ;
                                    } ;
                            
                                    $smsg = "(-NoTest:Defaulting Fast:`$true)" ;
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                    $props.add('Fast',$true) ;
                            
                                    return (New-Object -TypeName PsObject -Property $props) ;
                                }elseif(-not $env:ExchangeInstallPath){
                                    $smsg = "Non-Domain Joined machine, with NO ExchangeInstallPath e-vari: `nExchange is not installed locally: local computer resolution fails:`nPlease specify an explicit -Server, or -SiteName" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    $false | write-output ;
                                } else {
                                    $smsg = "$($env:computername) is both NON-Domain-joined -AND lacks an Exchange install (NO ExchangeInstallPath e-vari)`nPlease specify an explicit -Server, or -SiteName" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    $false | write-output ;
                                };
                            } CATCH {
                                $siteDN =$ExOPsiteDN ;
                                write-warning "`$siteDN lookup FAILED, deferring to hardcoded `$ExOPsiteDN string in infra file!" ;
                            } ;
                        } ;
                        $smsg = "Getting Exservers in Site:$($siteDN)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC") ;
                        $objectClass = "objectClass=msExchExchangeServer" ;
                        $version = "versionNumber>=1937801568" ;
                        $site = "msExchServerSite=$siteDN" ;
                        $search.Filter = "(&($objectClass)($version)($site))" ;
                        $search.PageSize = 1000 ;
                        [void] $search.PropertiesToLoad.Add("name") ;
                        [void] $search.PropertiesToLoad.Add("msexchcurrentserverroles") ;
                        [void] $search.PropertiesToLoad.Add("networkaddress") ;
                        [void] $search.PropertiesToLoad.Add("msExchServerSite") ;
                        [void] $search.PropertiesToLoad.Add("serialNumber") ;
                        [void] $search.PropertiesToLoad.Add("DistinguishedName") ;
                        $exchServers = $search.FindAll() ;
                        $Aggr = @() ;
                        foreach($exServer in $exchServers){
                            $fqdn = ($exServer.Properties.networkaddress |
                                Where-Object{$_ -match '^ncacn_ip_tcp:'}).split(':')[1] ;
                            if($NoTest){} else {
                                $rsp = test-connection $fqdn -count 1 -ea 0 ;
                            } ;
                            $props = @{
                                Name = $exServer.Properties.name[0]
                                FQDN=$fqdn;
                                Version = $exServer.Properties.serialnumber
                                Roles = $exserver.Properties.msexchcurrentserverroles
                                RoleNames = $null ;
                                DistinguishedName = $exserver.Properties.distinguishedname;
                                Site = @("$($exserver.Properties.msexchserversite -Replace '^CN=|,.*$')") ;
                                ResponseTime = if($rsp){$rsp.ResponseTime} else { 0} ;
                            } ;
                            $props.RoleNames = switch ($exserver.Properties.msexchcurrentserverroles){
                                2       {"MBX"}
                                4       {"CAS"}
                                16      {"UM"}
                                20      {"CAS;UM".split(';')}
                                32      {"HUB"}
                                36      {"CAS;HUB".split(';')}
                                38      {"CAS;HUB;MBX".split(';')}
                                54      {"MBX"}
                                64      {"EDGE"}
                                16385   {"CAS"}
                                16439   {"CAS;HUB;MBX".split(';')}
                            }
                            if($NoTest){
                                $smsg = "(-NoTest:Defaulting Fast:`$true)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                $props.add('Fast',$true) ;
                            }else {
                                $props.add('Fast',[boolean]($rsp.ResponseTime -le $SpeedThreshold)) ;
                            };
                            $Aggr += New-Object -TypeName PsObject -Property $props ;
                        } ;
                        $httmp = @{} ;
                        if($Roles){
                            [regex]$rgxRoles = ('(' + (($roles |%{[regex]::escape($_)}) -join '|') + ')') ;
                            $matched =  @( $aggr | ?{$_.Roles -match $rgxRoles}) ;
                            foreach($m in $matched){
                                if($httmp[$m.name]){} else {
                                    $httmp[$m.name] = $m ;
                                } ;
                            } ;
                        } ;
                        if($RoleNames){
                            foreach ($RoleName in $RoleNames){
                                $matched = @($Aggr | ?{$_.RoleNames -contains $RoleName} ) ;
                                foreach($m in $matched){
                                    if($httmp[$m.name]){} else {
                                        $httmp[$m.name] = $m ;
                                    } ;
                                } ;
                            } ;
                        } ;
                        if($Server){
                            foreach ($Name in $Server){
                                $matched = @($Aggr | ?{$_.Name -eq $Name} ) ;
                                foreach($m in $matched){
                                    if($httmp[$m.name]){} else {
                                        $httmp[$m.name] = $m ;
                                    } ;
                                } ;
                            } ;
                        } ;
                        if(($httmp.Values| measure).count -gt 0){
                            $Aggr  = $httmp.Values ;
                        } ;
                        $smsg = "Returning $((($Aggr|measure).count|out-string).trim()) match summaries to pipeline..." ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $Aggr | write-output ;
                    }CATCH{
                        Write-Error $_ ;
                    } ;
                } ;
                END{
                    $smsg = "$($_sBnr.replace('=v','=^').replace('v=','^='))" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            } ;
        }
        #*------^ END Function get-ADExchangeServerTDO ^------ ;

        $smsg = #*------v out-Clipboard.ps1 v------" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        Function out-Clipboard {
            [CmdletBinding()]
            Param (
                [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Content to be copied to clipboard [-Content `$object]")]
                [ValidateNotNullOrEmpty()]$Content,
                [Parameter(HelpMessage="Switch to suppress the default 'append `n' clip.exe-emulating behavior[-NoLegacy]")]
                [switch]$NoLegacy
            ) ;
            PROCESS {
                if($host.version.major -lt 3){
                    # provide clipfunction downrev
                    if(-not (get-command out-clipboard)){
                        # build the alias if not pre-existing
                        $tClip = "$((Resolve-Path $env:SystemRoot\System32\clip.exe).path)" ;
                        #$input | "($tClip)" ; 
                        #$content | ($tClip) ; 
                        Set-Alias -Name 'Out-Clipboard' -Value $tClip -scope script ;
                    } ;
                    $content | out-clipboard ;
                } else {
                    # emulate clip.exe's `n-append behavior on ps3+
                    if(-not $NoLegacy){
                        $content = $content | foreach-object {"$($_)$([Environment]::NewLine)"} ; 
                    } ; 
                    $content | set-clipboard ;
                } ; 
            } ; 
        }
        $smsg = #*------^ out-Clipboard.ps1 ^------" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

        # remove-SmtpPlusAddress.ps1


        #*------v remove-SmtpPlusAddress.ps1 v------
        function remove-SmtpPlusAddress {
            <#
            .SYNOPSIS
            remove-SmtpPlusAddress - Strips any Plus address Tag present in an smtp address, and returns the base address
            .NOTES
            Version     : 1.0.0
            Author      : Todd Kadrie
            Website     : http://www.toddomation.com
            Twitter     : @tostka / http://twitter.com/tostka
            CreatedDate : 2024-05-22
            FileName    : remove-SmtpPlusAddress
            License     : (none asserted)
            Copyright   : (none asserted)
            Github      : https://github.com/tostka/verb-Ex2010
            Tags        : Powershell,EmailAddress,Version
            AddedCredit : Bruno Lopes (brunokktro )
            AddedWebsite: https://www.linkedin.com/in/blopesinfo
            AddedTwitter: @brunokktro / https://twitter.com/brunokktro
            REVISIONS
            * 1:47 PM 7/9/2024 CBA github field correction
            * 1:22 PM 5/22/2024init
            .DESCRIPTION
            remove-SmtpPlusAddress - Strips any Plus address Tag present in an smtp address, and returns the base address

            Plus Addressing is supported in Exchange Online, Gmail, and other select hosts. 
            It is *not* supported for Exchange Server onprem. Any + addressed email will read as an unresolvable email address. 
            Supporting systems will truncate the local part (in front of the @), after the +, to resolve the email address for normal routing:

            monitoring+whatever@domain.tld, is cleaned down to: monitor@domain.tld. 

            .PARAMETER EmailAddress
            SMTP Email Address
            .OUTPUT
            String
            .EXAMPLE
            PS> 
            PS> $returned = remove-SmtpPlusAddress -EmailAddress 'monitoring+SolarWinds@toro.com';  
            PS> $returned ; 
            Demo retrieving get-EmailAddress, assigning to output, processing it for version info, and expanding the populated returned values to local variables. 
            .EXAMPLE
            ps> remove-SmtpPlusAddress -EmailAddress 'monitoring+SolarWinds@toro.com;notanemailaddresstoro.com,todd+spam@kadrie.net' -verbose ;
            Demo with comma and semicolon delimiting, and an invalid address (to force a regex match fail error).
            .LINK
            https://github.com/brunokktro/EmailAddress/blob/master/Get-ExchangeEnvironmentReport.ps1
            .LINK
            https://github.com/tostka/verb-Ex2010
            #>
            [CmdletBinding()]
            #[Alias('rvExVers')]
            PARAM(
                [Parameter(Mandatory = $true,Position=0,HelpMessage="Object returned by a get-EmailAddress command[-EmailAddress `$ExObject]")]
                    [string[]]$EmailAddress
            ) ;
            BEGIN {
                ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
                $verbose = $($VerbosePreference -eq "Continue")
                $rgxSMTPAddress = "([0-9a-zA-Z]+[-._+&='])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}" ; 
                $sBnr="#*======v $($CmdletName): v======" ;
                write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
                if($EmailAddress -match ','){
                    $smsg = "(comma detected, attempting split on commas)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $EmailAddress = $EmailAddress.split(',') ; 
                } ; 
                if($EmailAddress -match ';'){
                    $smsg = "(semi-colon detected, attempting split on semicolons)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $EmailAddress = $EmailAddress.split(';') ; 
                } ; 
            }
            PROCESS {
                foreach ($item in $EmailAddress){
                    if($item -match $rgxSMTPAddress){
                        if($item.split('@')[0].contains('+')){
                            write-verbose  "Remove Plus Addresses from: $($item)" ; 
                            $lpart,$domain = $item.split('@') ; 
                            $item = "$($lpart.split('+')[0])@$($domain)" ; 
                            $smsg = "Cleaned Address: $($item)" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        }
                        $item | write-output ; 
                    } else { 
                        write-warning  "$($item)`ndoes not match a standard SMTP Email Address (skipping):`n$($rgxSmtpAddress)" ; 
                        continue ;
                    } ; 
                } ;     
            
            } # PROC-E
            END{
                write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
            }
        }; 
        #*------^ remove-SmtpPlusAddress.ps1 ^------


        #endregion FUNCTIONS ; #*======^ END FUNCTIONS ^======

        #region SUBMAIN ; #*======v SUB MAIN v======
        $smsg = #*------v Function SUB MAIN v------"
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

        #region BANNER ; #*------v BANNER v------
        $sBnr="#*======v $(${CmdletName}): v======" ;
        $smsg = $sBnr ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #endregion BANNER ; #*------^ END BANNER ^------
            

        #region START-LOG-HOLISTIC #*------v START-LOG-HOLISTIC v------
        # Single log for script/function example that accomodates detect/redirect from AllUsers scope'd installed code, and hunts a series of drive letters to find an alternate logging dir (defers to profile variables)
        #${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ;
        if(!(get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
        foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
        if(!(get-variable rgxPSAllUsersScope -ea 0)){
            $rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps(((d|m))*)1|dll)$" ;
        } ;
        if(!(get-variable rgxPSCurrUserScope -ea 0)){
            $rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps((d|m)*)1|dll)$" ;
        } ;
        $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($whatif) ;} ;
        # if using [CmdletBinding(SupportsShouldProcess)] + -WhatIf:$($WhatIfPreference):
        #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
        #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag="$($ticket)-$($TenOrg)-LASTPASS-" ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
        #$pltSL.Tag = $ModuleName ; 
        if($ticket){$pltSL.Tag = $ticket} ; 
        if($script:rPSCommandPath){ $prxPath = $script:rPSCommandPath }
        elseif($script:PSCommandPath){$prxPath = $script:PSCommandPath}
        if($rMyInvocation.MyCommand.Definition){$prxPath2 = $rMyInvocation.MyCommand.Definition }
        elseif($MyInvocation.MyCommand.Definition){$prxPath2 = $MyInvocation.MyCommand.Definition } ; 
        if($prxPath){
            if(($prxPath -match $rgxPSAllUsersScope) -OR ($prxPath -match $rgxPSCurrUserScope)){
                $bDivertLog = $true ; 
                switch -regex ($prxPath){
                    $rgxPSAllUsersScope{$smsg = "AllUsers"} 
                    $rgxPSCurrUserScope{$smsg = "CurrentUser"}
                } ;
                $smsg += " context script/module, divert logging into [$budrv]:\scripts" 
                write-verbose $smsg  ;
                if($bDivertLog){
                    if((split-path $prxPath -leaf) -ne $cmdletname){
                        # function in a module/script installed to allusers|cu - defer name to Cmdlet/Function name
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
                    } else {
                        # installed allusers|CU script, use the hosting script name
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $prxPath -leaf)) ;
                    }
                } ;
            } else {
                $pltSL.Path = $prxPath ;
            } ;
       }elseif($prxPath2){
            if(($prxPath2 -match $rgxPSAllUsersScope) -OR ($prxPath2 -match $rgxPSCurrUserScope) ){
                 $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $prxPath2 -leaf)) ;
            } elseif(test-path $prxPath2) {
                $pltSL.Path = $prxPath2 ;
            } elseif($cmdletname){
                $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
            } else {
                $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$CMDLETNAME, FROM WHICH TO BUILD A START-LOG.PATH!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                BREAK ;
            } ; 
        } else{
            $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$CMDLETNAME, FROM WHICH TO BUILD A START-LOG.PATH!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            BREAK ;
        }  ;
        write-verbose "start-Log w`n$(($pltSL|out-string).trim())" ; 
        $logspec = start-Log @pltSL ;
        $error.clear() ;
        TRY {
            if($logspec){
                #$logging=$logspec.logging ;
                $logging= $false ; # explicitly turned logfile writing off, just want to use it's path for exports
                $logfile=$logspec.logfile ;
                $transcript=$logspec.transcript ;
                <# 2:30 PM 9/27/2024 no transcript, just want solid logging path discovery
                $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                if($stopResults){
                    $smsg = "Stop-transcript:$($stopResults)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } ; 
                $startResults = start-Transcript -path $transcript ;
                if($startResults){
                    $smsg = "start-transcript:$($startResults)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
                #>
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
        } ;
        #endregion START-LOG-HOLISTIC #*------^ END START-LOG-HOLISTIC ^------


        #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======
        # PRETUNE STEERING separately *before* pasting in balance of region
        #*------v STEERING VARIS v------
        $useO365 = $false ;
        $useEXO = $false ; 
        $UseOP=$true ; 
        $UseExOP=$true ;
        $useExopNoDep = $true ; # switch to use Connect-ExchangeServerTDO, vs connect-ex2010 (creds are assumed inherent to the account)
        $useForestWide = $true ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
        $UseOPAD = $false ; 
        $UseMSOL = $false ; # should be hard disabled now in o365
        $UseAAD = $false  ; 
        $useO365 = [boolean]($useO365 -OR $useEXO -OR $UseMSOL -OR $UseAAD)
        $UseOP = [boolean]($UseOP -OR $UseExOP -OR $UseOPAD) ;
        #*------^ END STEERING VARIS ^------
        #*------v EXO V2/3 steering constants v------
        $EOMModName =  'ExchangeOnlineManagement' ;
        $EOMMinNoWinRMVersion = $MinNoWinRMVersion = '3.0.0' ; # support both names
        #*------^ END EXO V2/3 steering constants ^------
        # assert Org from Credential specs (if not param'd)
        # 1:36 PM 7/7/2023 and revised again -  revised the -AND, for both, logic wasn't working
        if($TenOrg){    
            $smsg = "Confirmed populated `$TenOrg" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } elseif(-not($tenOrg) -and $Credential){
            $smsg = "(unconfigured `$TenOrg: asserting from credential)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $TenOrg = get-TenantTag -Credential $Credential ;
        } else { 
            # if not using Credentials or a TargetTenants/TenOrg loop, default the $TenOrg on the $env:USERDOMAIN
            $smsg = "(unconfigured `$TenOrg & *NO* `$Credential: fallback asserting from `$env:USERDOMAIN)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            switch -regex ($env:USERDOMAIN){
                ([regex]('(' + (( @($TORMeta.legacyDomain,$CMWMeta.legacyDomain)  |foreach-object{[regex]::escape($_)}) -join '|') + ')')).tostring() {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
                $TOLMeta.legacyDomain {$TenOrg = 'TOL' }
                default {throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ; exit ; } ;
            } ; 
        } ; 
        #region useO365 ; #*------v useO365 v------
        #$useO365 = $false ; # non-dyn setting, drives variant EXO reconnect & query code
        #if($CloudFirst){ $useO365 = $true } ; # expl: steering on a parameter
        if($useO365){
            #region GENERIC_EXO_CREDS_&_SVC_CONN #*------v GENERIC EXO CREDS & SVC CONN BP v------
            # o365/EXO creds
            <### Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile*
            $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
            Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
            Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
            Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            ###>
            $o365Cred = $null ;
            if($Credential){
                $smsg = "`Credential:Explicit credentials specified, deferring to use..." ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                # get-TenantCredentials() return format: (emulating)
                $o365Cred = [ordered]@{
                    Cred=$Credential ; 
                    credType=$null ; 
                } ; 
                $uRoleReturn = resolve-UserNameToUserRole -UserName $Credential.username -verbose:$($VerbosePreference -eq "Continue") ; # Username
                #$uRoleReturn = resolve-UserNameToUserRole -Credential $Credential -verbose = $($VerbosePreference -eq "Continue") ;   # full Credential support
                if($uRoleReturn.UserRole){
                    $o365Cred.credType = $uRoleReturn.UserRole ; 
                } else { 
                    $smsg = "Unable to resolve `$credential.username ($($credential.username))"
                    $smsg += "`nto a usable 'UserRole' spec!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                    Break ;
                } ; 
            } else { 
                $pltGTCred=@{TenOrg=$TenOrg ; UserRole=$null; verbose=$($verbose)} ;
                if((get-command get-TenantCredentials).Parameters.keys -contains 'silent'){
                    $pltGTCred.add('silent',$silent) ; 
                } ;
                if($UserRole){
                    $smsg = "(`$UserRole specified:$($UserRole -join ','))" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $pltGTCred.UserRole = $UserRole; 
                } else { 
                    $smsg = "(No `$UserRole found, defaulting to:'CSVC','SID' " ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    $pltGTCred.UserRole = 'CSVC','SID' ; 
                } ; 

                $smsg = "get-TenantCredentials w`n$(($pltGTCred|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $o365Cred = get-TenantCredentials @pltGTCred
            } ; 
            if($o365Cred.credType -AND $o365Cred.Cred -AND $o365Cred.Cred.gettype().fullname -eq 'System.Management.Automation.PSCredential'){
                $smsg = "(validated `$o365Cred contains .credType:$($o365Cred.credType) & `$o365Cred.Cred.username:$($o365Cred.Cred.username)" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                # 9:58 AM 6/13/2024 populate $credential with return, if not populated (may be required for follow-on calls that pass common $Credentials through)
                if((gv Credential) -AND $Credential -eq $null){
                    $credential = $o365Cred.Cred ;
                }elseif($credential.gettype().fullname -eq 'System.Management.Automation.PSCredential'){
                    $smsg = "(`$Credential is properly populated; explicit -Credential was in initial call)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } else {
                    $smsg = "`$Credential is `$NULL, AND $o365Cred.Cred is unusable to populate!" ;
                    $smsg = "downstream commands will *not* properly pass through usable credentials!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                    break ;
                } ;
            } else {
                $smsg = "UNABLE TO RESOLVE FUNCTIONAL CredType/UserRole from specified explicit -Credential:$($Credential.username)!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                break ;
            } ; 
            if($o365Cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name cred$($tenorg) -scope Script -ea 0 ){ remove-Variable -Name cred$($tenorg) -scope Script } ;
                New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
                $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatettus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ;
            # if we get here, wo a $Credential, w resolved $o365Cred, assign it 
            if(-not $Credential -AND $o365Cred){$Credential = $o365Cred.cred } ; 
            # configure splat for connections: (see above useage)
            # downstream commands
            $pltRXO = [ordered]@{
                Credential = $Credential ;
                verbose = $($VerbosePreference -eq "Continue")  ;
            } ;
            if((get-command Reconnect-EXO).Parameters.keys -contains 'silent'){
                $pltRxo.add('Silent',$silent) ;
            } ;
            # default connectivity cmds - force silent false
            $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$true) ; 
            if((get-command Reconnect-EXO).Parameters.keys -notcontains 'silent'){
                $pltRxo.remove('Silent') ;
            } ; 
            #region EOMREV ; #*------v EOMREV Check v------
            #$EOMmodname = 'ExchangeOnlineManagement' ;
            $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
            # do a gmo first, faster than gmo -list
            if([version]$EOMMv = (Get-Module @pltIMod).version){}
            elseif([version]$EOMMv = (Get-Module -ListAvailable @pltIMod).version){}
            else {
                $smsg = "$($EOMmodname) PowerShell v$($MinNoWinRMVersion) module is required, do you want to install it?" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                else{ $smsg = "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bRet = Read-Host "Enter YYY to continue. Anything else will exit"  ;
                if ($bRet.ToUpper() -eq "YYY") {
                    $smsg = "Installing $($EOMmodname) module..." ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Install-Module $EOMmodname -Repository PSGallery -AllowClobber -Force ;
                } else {
                    $smsg = "Please install $($EOMmodname) PowerShell v$($MinNoWinRMVersion)  module." ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #exit 1
                    break ;
                }  ;
            } ;
            $smsg = "(Checking for WinRM support in this EOM rev...)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            if([version]$EOMMv -ge [version]$MinNoWinRMVersion){
                $MinNoWinRMVersion = $EOMMv.tostring() ;
                $IsNoWinRM = $true ;
            }elseif([version]$EOMMv -lt [version]$MinimumVersion){
                $smsg = "Installed $($EOMmodname) is v$($MinNoWinRMVersion): This module is obsolete!" ;
                $smsg += "`nAnd unsupported by this function!" ;
                $smsg += "`nPlease install $($EOMmodname) PowerShell v$($MinNoWinRMVersion)  module!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Break ;
            } else {
                $IsNoWinRM = $false ;
            } ;
            [boolean]$UseConnEXO = [boolean]([version]$EOMMv -ge [version]$MinNoWinRMVersion) ;
            #endregion EOMREV ; #*------^ END EOMREV Check  ^------
            #-=-=-=-=-=-=-=-=
            <### CALLS ARE IN FORM: (cred$($tenorg))
            # downstream commands
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; 
            } ;
            if((get-command Reconnect-EXO).Parameters.keys -contains 'silent'){
                $pltRxo.add('Silent',$false) ;
            } ; 
            # default connectivity cmds - force silent false
            $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$true) ;
            if((get-command Reconnect-EXO).Parameters.keys -notcontains 'silent'){
                $pltRxo.remove('Silent') ;
            } ; 
            #$pltRXO creds & .username can also be used for AzureAD connections:
            #Connect-AAD @pltRXOC ;
            ###>
            #endregion GENERIC_EXO_CREDS_&_SVC_CONN #*------^ END GENERIC EXO CREDS & SVC CONN BP ^------

        } else {
            $smsg = "(`$useO365:$($useO365))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; # if-E if($useO365 ){
        #endregion useO365 ; #*------^ END useO365 ^------

        #region useEXO ; #*------v useEXO v------
        # 1:29 PM 9/15/2022 as of MFA & v205, have to load EXO *before* any EXOP, or gen get-steppablepipeline suffix conflict error
        if($useEXO){
            if ($script:useEXOv2 -OR $useEXOv2) { reconnect-eXO2 @pltRXOC }
            else { reconnect-EXO @pltRXOC } ;
        } else {
            $smsg = "(`$useEXO:$($useEXO))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; # if-E 
        #endregion  ; #*------^ END useEXO ^------
      
        #region GENERIC_EXOP_CREDS_&_SRVR_CONN #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
        # steer all onprem code on $XXXMeta.ExOPAccessFromToro & Ex10Server values
        #$UseOP=$true ; 
        #$UseExOP=$true ;
        #$useExopNoDep = $true # switch to use Connect-ExchangeServerTDO, vs connect-ex2010 (creds are assumed inherent to the account) 
        #$useForestWide = $true ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
        <# no onprem dep
        if((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro -AND (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server){
            $UseOP = $UseExOP = $true ;
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`ENABLING use of OnPrem Ex system this pass." ;
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        } else {
            $UseOP = $UseExOP = $false ;
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`nDISABLING use of OnPrem Ex system this pass." ;
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        } ;
        #>
        if($UseOP){
            if($useExopNoDep){
                # Connect-ExchangeServerTDO use: creds are implied from the PSSession creds; assumed to have EXOP perms
            } else {
                #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
                # do the OP creds too
                $OPCred=$null ;
                # default to the onprem svc acct
                # userrole='ESVC','SID'
                #$pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC','SID'; verbose=$($verbose)} ;
                # userrole='SID','ESVC'
                $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='SID','ESVC'; verbose=$($verbose)} ;
                $smsg = "get-HybridOPCredentials w`n$(($pltGHOpCred|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                    # make it script scope, so we don't have to predetect & purge before using new-variable
                    if(get-Variable -Name "cred$($tenorg)OP" -scope Script -ea 0 ){ remove-Variable -Name "cred$($tenorg)OP" -scope Script } ;
                    New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                    $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } else {
                    $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    $script:PassStatus += $statusdelta ;
                    set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                    $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                    Break ;
                } ;
                $smsg= "Using OnPrem/EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                                                            <### CALLS ARE IN FORM: (cred$($tenorg))
                $pltRX10 = @{
                    Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                    #verbose = $($verbose) ;
                    Verbose = $FALSE ; 
                } ;
                $1stConn = $false ; # below uses silent suppr for both x10 & xo!
                if($1stConn){
                    $pltRX10.silent = $pltRXO.silent = $false ;
                } else {
                    $pltRX10.silent = $pltRXO.silent =$true ;
                } ;
                if($pltRX10){ReConnect-Ex2010 @pltRX10 }
                else {ReConnect-Ex2010 }
                #$pltRx10 creds & .username can also be used for local ADMS connections
                ###>
                $pltRX10 = @{
                    Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                    #verbose = $($verbose) ;
                    Verbose = $FALSE ; 
                } ;
                if((get-command Reconnect-Ex2010).Parameters.keys -contains 'silent'){
                    $pltRX10.add('Silent',$false) ;
                } ;
            } ; 
            # defer cx10/rx10, until just before get-recipients qry
            # connect to ExOP X10
            if($useEXOP){
                if($useExopNoDep){ 
                    $smsg = "(Using ExOP:Connect-ExchangeServerTDO())" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;           
                    TRY{
                        $Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name 
                    }CATCH{$Site=$env:COMPUTERNAME} ;
                    $PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
                } else {
                    if($pltRX10){
                        #ReConnect-Ex2010XO @pltRX10 ;
                        ReConnect-Ex2010 @pltRX10 ;
                    } else { Reconnect-Ex2010 ; } ;
                    #Add-PSSnapin -Name 'Microsoft.Exchange.Management.PowerShell.SnapIn'
                    #TK: add: test Exch & AD functional connections
                    TRY{
                        if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig'){} else {
                            $smsg = "(mangled Ex10 conn: dx10,rx10...)" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            disconnect-ex2010 ; reconnect-ex2010 ; 
                        } ; 
                        if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                            $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ; 
                            throw $smsg ; 
                            $smsg | write-warning  ; 
                        } ; 
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = $ErrTrapd ;
                        $smsg += "`n";
                        $smsg += $ErrTrapd.Exception.Message ;
                        if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        CONTINUE ;
                    } ;
                }
            } ; 
            if($useForestWide){
                #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT v------
                $smsg = "(`$useForestWide:$($useForestWide)):Enabling EXoP Forestwide)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Set-AdServerSettings -ViewEntireForest $True ;
                #endregion  ; #*------^ END OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT ^------
            } ;
        } else {
            $smsg = "(`$useOP:$($UseOP))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        }  ;  # if-E $UseOP
        #endregion GENERIC_EXOP_CREDS_&_SRVR_CONN #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
            
        #region UseOPAD #*------v UseOPAD v------
        if($UseOP -OR $UseOPAD){
            #region GENERIC_ADMS_CONN_&_XO #*------v GENERIC ADMS CONN & XO  v------
            $smsg = "(loading ADMS...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # always capture load-adms return, it outputs a $true to pipeline on success
            $ADMTLoaded = load-ADMS -Verbose:$FALSE ;
            # 9:32 AM 4/20/2023 trimmed disabled/fw-borked cross-org code
            TRY {
                if(-not(Get-ADDomain  -ea STOP).DNSRoot){
                    $smsg = "Missing AD Connection! (no (Get-ADDomain).DNSRoot returned)" ; 
                    throw $smsg ; 
                    $smsg | write-warning  ; 
                } ; 
                $objforest = get-adforest -ea STOP ; 
                # Default new UPNSuffix to the UPNSuffix that matches last 2 elements of the forestname.
                $forestdom = $UPNSuffixDefault = $objforest.UPNSuffixes | ?{$_ -eq (($objforest.name.split('.'))[-2..-1] -join '.')} ; 
                if($useForestWide){
                    #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE AD GC QRY SUPPORT v------
                    $smsg = "(`$useForestWide:$($useForestWide)):Enabling AD Forestwide)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #TK 9:44 AM 10/6.2.32 need org wide for rolegrps in parent dom (only for onprem RBAC, not EXO)
                    $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;        
                    #endregion  ; #*------^ END  OPTIONAL CODE TO ENABLE FOREST-WIDE AD GC QRY SUPPORT  ^------
                } ;    
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = $ErrTrapd ;
                $smsg += "`n";
                $smsg += $ErrTrapd.Exception.Message ;
                if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                CONTINUE ;
            } ;        
            #endregion GENERIC_ADMS_CONN_&_XO #*------^ END GENERIC ADMS CONN & XO ^------
        } else {
            $smsg = "(`$UseOP:$($UseOP))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        }  ;
        #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
        #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller = get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
        # use new get-GCFastXO cross-org dc finde
        # default to Op_ExADRoot forest from $TenOrg Meta
        #if($UseOP -AND -not $domaincontroller){
        if($UseOP -AND -not (get-variable domaincontroller -ea 0)){
            #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((get-variable -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};
            # need to debug the above, credential issue?
            # just get it done
            $domaincontroller = get-GCFast
        }  else { 
            # have to defer to get-azuread, or use EXO's native cmds to poll grp members
            # TODO 1/15/2021
            $useEXOforGroups = $true ; 
            $smsg = "$($TenOrg):HAS NO ON-PREM ACTIVEDIRECTORY, DEFERRING ALL GROUP ACCESS & MGMT TO NATIVE EXO CMDS!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        if($useForestWide -AND -not $GcFwide){
            #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT: v------
            $smsg = "`$GcFwide = Get-ADDomainController -Discover -Service GlobalCatalog" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;
            $smsg = "Discovered `$GcFwide:$($GcFwide)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #endregion  ; #*------^ END OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT ^------
        } ;
        #endregion UseOPAD #*------^ END UseOPAD ^------

        #region MSOL_CONNECTION ; #*------v  MSOL CONNECTION v------
        #$UseMSOL = $false 
        if($UseMSOL){
            #$reqMods += "connect-msol".split(";") ;
            #if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
            $smsg = "(loading MSOL...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #connect-msol ;
            connect-msol @pltRXOC ;
        } else {
            $smsg = "(`$UseMSOL:$($UseMSOL))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;
        #endregion MSOL_CONNECTION ; #*------^  MSOL CONNECTION ^------

        #region AZUREAD_CONNECTION ; #*------v AZUREAD CONNECTION v------
        #$UseAAD = $false 
        if($UseAAD){
            #$reqMods += "Connect-AAD".split(";") ;
            #if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
            $smsg = "(loading AAD...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Connect-AAD @pltRXOC ;
        } else {
            $smsg = "(`$UseAAD:$($UseAAD))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;
        #endregion AZUREAD_CONNECTION ; #*------^ AZUREAD CONNECTION ^------
      
        <# defined above
        # EXO connection
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            verbose = $($verbose) ; } ;
        #>
        <#
        if($VerbosePreference = "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        disconnect-exo ;
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXOC }
        else { reconnect-EXO @pltRXOC } ;
        # reenable VerbosePreference:Continue, if set, during mod loads
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        #>
        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        #endregion SERVICE_CONNECTIONS #*======^ END SERVICE_CONNECTIONS ^======

        # SET DAYS=0 IF USING START/END (they only get used when days is non-0); $platIn.TAG is appended to ticketNO for output vari $vn, and $ofile
        if($Days -AND ($Start -OR $End)){
            write-warning "specified -Days with (-Start -OR -End); If using Start/End, specify -Days 0!" ; 
            Break ; 
        } ; 
        if($Start){[datetime]$Start = $Start } ; 
        if($End){[datetime]$End = $End} ; 
        if( ($null -eq $ResultSize) -OR ('' -eq $ResultSize) ){
            $ResultSize = 'unlimited' ; 
        } else { 
            throw "Resultsize must be an integer or the string 'unlimited' (or blank)" ; 
        } ;
        $pltParams=@{
            ticket=$ticket ;
            Requestor=$Requestor;
            days=$days ;
            Start=$Start ;
            End= $End ;
            Sender=$Sender ;
            Recipients=$Recipients ;
            MessageSubject=$MessageSubject ;
            EventID=$EventID ;
            MessageID=$MessageID;
            InternalMessageId=$InternalMessageId;
            NetworkMessageId=$NetworkMessageId;
            Reference=$Reference ; 
            ResultSize=$ResultSize ;
            Source=$Source ;
            Tag=$Tag ;
            ErrorAction = 'STOP' ;
            verbose = $($VerbosePreference -eq "Continue") ;
        } ; 

        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        # Configure the Get-xoMessageTrace splat 
        # w Ex2010 in the mix, and Ps2, [ordered] hashes won't work, config for:
        if($host.version.major -ge 3){$pltGMTL=[ordered]@{Dummy = $null ;} }
        #else {$pltGMTL = New-Object Collections.Specialized.OrderedDictionary} ;
        else{$pltGMTL=@{Dummy = $null ;} }
        If($pltGMTL.Contains("Dummy")){$pltGMTL.remove("Dummy")} ;
        $pltGMTL.add('ErrorAction',"STOP") ; 
        $pltGMTL.add('verbose',$($VerbosePreference -eq "Continue")) ; 

        <#$pltGMTL=[ordered]@{
            #SenderAddress=$SenderAddress;
            #RecipientAddress=$RecipientAddress;
            #Start=(get-date $Start);
            #Start= $Start;
            #End=(get-date $End);
            #End=$End;
            #Page= 1 ; # default it to 1 vs $null as we'll be purging empties further down
            ErrorAction = 'STOP' ;
            verbose = $($VerbosePreference -eq "Continue") ;
        } ;
        #>

        if ($PSCmdlet.ParameterSetName -eq 'Dates') {
            if($End -and -not $Start){
                $Start = (get-date $End).addDays(-1 * $DaysLimit) ; 
            } ; 
            if($Start -and -not ($End)){
                $smsg = "(Start w *NO* End, asserting currenttime)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $End=(get-date) ;
            } ;
        } else {
            if (-not $Days) {
                $Start = (get-date $End).addDays(-1 * $DaysLimit) ; 
                $smsg = "No Days, Start or End specified. Defaulting to $($DaysLimit)day Search window:$((get-date).adddays(-1 * $DaysLimit))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $End = (get-date) ;
                $Start = (get-date $End).addDays(-1 * $Days) ; 
                $smsg = "-Days:$($Days) specified: "
                #$smsg += "calculated Start:$((get-date $Start -format $sFulltimeStamp ))" ; 
                #$smsg += ", calculated End:$((get-date $End -format $sFulltimeStamp ))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #(get-date -format $sFiletimestamp);
            } ; 
        } ;

        $smsg = "(converting `$Start & `$End to UTC, using input as `$StartLocal & `$EndLocal)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        # convert dates to GMT .ToUniversalTime(
        $Start = ([datetime]$Start).ToUniversalTime() ; 
        $End = ([datetime]$End).ToUniversalTime() ; 
        $StartLocal = ([datetime]$Start).ToLocalTime() ; 
        $EndLocal = ([datetime]$End).ToLocalTime() ; 
        
        # sanity test the start/end dates, just in case (won't throw an error in gxmt)
        if($Start -gt $End){
            $smsg = "`-Start:$($Start) is GREATER THAN -End:($End)!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            throw $smsg ; 
            break ; 
        } ;

        $smsg = "`$Start:$(get-date -Date $StartLocal -format $sFulltimeStamp )" ;
        $smsg += "`n`$End:$(get-date -Date $EndLocal -format $sFulltimeStamp )" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        if((New-TimeSpan -Start $Start -End $End).days -gt $DaysLimit){
            $smsg = "Search span (between -Start & -End, or- Days in use) *exceeds* MS supported days history limit!`nReduce the window below a historical 10d, or use get-HistoricalSearch instead!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            Break ; 
        } ; 
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

        <#
        [Parameter(Mandatory=$True,HelpMessage="Ticket Number [-Ticket 'TICKETNO']")]
            [Alias('tix')]
            [string]$ticket,
        [Parameter(Mandatory=$False,HelpMessage="Ticket Customer email identifier. [-Requestor 'fname.lname@domain.com']")]
            [Alias('UID')]
            [string]$Requestor,
        [Parameter(ParameterSetName='Dates',HelpMessage="Start of range to be searched[-StartDate '11/5/2021 2:16 PM']")]
            [Alias('StartDate')]
            [DateTime]$Start,
        [Parameter(ParameterSetName='Dates',HelpMessage="End of range to be searched (defaults to current time if unspecified)[-EndDate '11/5/2021 5:16 PM']")]
            [Alias('EndDate')]
            [DateTime]$End=(get-date),
        [Parameter(ParameterSetName='Days',HelpMessage="Days to be searched, back from current time(Alt to use of StartDate & EndDate; Note:MS won't search -gt 10 days)[-Days 7]")]
            #[ValidateRange(0,[int]::MaxValue)]
            [ValidateRange(0,10)] # MS won't search beyond 10, and silently returns incomplete results
            [int]$Days,
        [Parameter(HelpMessage="SenderAddress (an array runs search on each)[-SenderAddress addr@domain.com]")]
            [Alias('SenderAddress')]
            [string]$Sender, # MultiValuedProperty
        [Parameter(HelpMessage="RecipientAddress (an array runs search on each)[-RecipientAddress addr@domain.com]")]
                [Alias('RecipientAddress')]
                [string[]]$Recipients, # MultiValuedProperty
        [Parameter(HelpMessage="Subject of target message (emulated via post filtering, not supported param of Get-xoMessageTrace) [-Subject 'Some subject']")]
                [Alias('subject')]
                [string]$MessageSubject,
        [Parameter(Mandatory=$False,HelpMessage="Corresponds to the value of the Message-Id: header field in the message. Be sure to include the full MessageId string (which may include angle brackets) and enclose the value in quotation marks[-MessageId `"<nnnn-nnn...>`"]")]
            [string]$MessageId,
        [Parameter(Mandatory=$False,HelpMessage="The InternalMessageId parameter filters the message tracking log entries by the value of the InternalMessageId field. The InternalMessageId value is a message identifier that's assigned by the Exchange server that's currently processing the message.  The value of the internal-message-id for a specific message is different in the message tracking log of every Exchange server that's involved in the delivery of the message.")]
            [string]$InternalMessageId,
        [Parameter(Mandatory=$False,HelpMessage="This field contains a unique message ID value that persists across copies of the message that may be created due to bifurcation or distribution group expansion.(Ex16,19)")]
            [string]$NetworkMessageId,
        [Parameter(Mandatory=$False,HelpMessage="The Reference field contains additional information for specific types of events. For example, the Reference field value for a DSN message tracking entry contains the InternalMessageId value of the message that caused the DSN. For many types of events, the value of Reference is blank")]
            [string]$Reference,
        [Parameter(HelpMessage="The Status parameter filters the results by the delivery status of the message (None|GettingStatus|Failed|Pending|Delivered|Expanded|Quarantined|FilteredAsSpam),an array runs search on each). [-Status 'Failed']")]
                [Alias('DeliveryStatus','Status')]
                [ValidateSet("RECEIVE","DELIVER","FAIL","SEND","RESOLVE","EXPAND","TRANSFER","DEFER","")]
                [string[]]$EventId, # MultiValuedProperty
        [Parameter(Mandatory=$False,HelpMessage="The TransportTrafficType parameter filters the message tracking log entries by the value of the TransportTrafficType field. However, this field isn't interesting for on-premises Exchange organizations[-TransportTrafficType xxx]")]
            [string]$TransportTrafficType, 
        [Parameter(Mandatory=$False,HelpMessage="Source (STOREDRIVER|SMTP|DNS|ROUTING)[-Source STOREDRIVER]")]
            [ValidateSet("STOREDRIVER","SMTP","DNS","ROUTING","")]
            [string]$Source,
        [Parameter(Mandatory=$False,HelpMessage="The Server parameter specifies the Exchange server where you want to run this command. You can use any value that uniquely identifies the server. For example: Name, FQDN, Distinguished name (DN), Exchange Legacy DN. If you don't use this parameter, the command is run on the local server[-Server Servername]")]
            [string]$Server,
        [Parameter(Mandatory=$False,HelpMessage="The ResultSize parameter specifies the maximum number of results to return. If you want to return all requests that match the query, use unlimited for the value of this parameter. The default value is 1000.")]
            [ValidateScript({
                if( ($_ -match 'unlimited') -OR ($_.gettype().fullname -eq 'System.Int32') -OR ($null -eq $_) -OR ('' -eq $_) ){
                    $true ; 
                } else { 
                    throw "Resultsize must be an integer or the string 'unlimited' (or blank)" ; 
                } ;
            })]
            $ResultSize,
        [Parameter(HelpMessage="Tag string (Variable Name compatible: no spaces A-Za-z0-9_ only) that is used for Variables and export file name construction. [-Tag 'LastDDGSend']")] 
            [ValidatePattern('^[A-Za-z0-9_]*$')]
            [string]$Tag,
        [Parameter(HelpMessage="switch to just return the net messages on the initial track (no Fail/Quarantine, MTDetail or other post-processing summaries) [-simpletrack]")]
            [switch]$SimpleTrack,
        [Parameter(HelpMessage="switch to perform configured csv exports of results (defaults true) [-DoExports]")]
            [switch]$DoExports=$TRUE,
        [Parameter(HelpMessage="switch to perform MessageTrackingDetail pass, after intial MessageTrace (up to limit specified in -MessageTraceDetailLimit (defaults true) [-Detailed]")]
            [switch]$Detailed,
        # Service Connection Supporting Varis (AAD, EXO, EXOP)
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
            [ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [string]$TenOrg = $global:o365_TenOrgDefault,
        [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]")]
            [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false, HelpMessage = "Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]")]
            # sourced from get-admincred():#182: $targetRoles = 'SID', 'CSID', 'ESVC','CSVC','UID','ESvcCBA','CSvcCBA','SIDCBA' ;
            #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
            # pulling the pattern from global vari w friendly err
            [ValidateScript({
                if(-not $rgxPermittedUserRoles){$rgxPermittedUserRoles = '(SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)'} ;
                if(-not ($_ -match $rgxPermittedUserRoles)){throw "'$($_)' doesn't match `$rgxPermittedUserRoles:`n$($rgxPermittedUserRoles.tostring())" ; } ;
                return $true ;
            })]
            [string[]]$UserRole = @('SIDCBA','SID','CSVC'),
            #@('SID','CSVC'),
            # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent
        #>

        #if(-not(gcm Remove-InvalidVariableNameChars -ea 0)){ Function Remove-InvalidVariableNameChars ([string]$Name) {($Name.tochararray() -match '[A-Za-z0-9_]') -join '' | write-output } } ;
        #$prpMTLfta = 'Timestamp','EventId','Sender','Recipients','MessageSubject' ; 
        #$prpXCsv = "Timestamp",@{N='TimestampLocal'; E={$_.Timestamp.ToLocalTime()}},"Source","EventId","RelatedRecipientAddress","Sender",@{N='Recipients'; E={$_.Recipients}},"RecipientCount",@{N='RecipientStatus'; E={$_.RecipientStatus}},"MessageSubject","TotalBytes",@{N='Reference'; E={$_.Reference}},"MessageLatency","MessageLatencyType","InternalMessageId","MessageId","ReturnPath","ClientIp","ClientHostname","ServerIp","ServerHostname","ConnectorId","SourceContext","MessageInfo",@{N='EventData'; E={$_.EventData}} ;
        #$prpMTFailFL = 'Timestamp','ClientHostname','Source','EventId','Recipients','RecipientStatus','MessageSubject','ReturnPath' ;
        #$pltGMTL=@{  resultsize="UNLIMITED" } ;

        #$ofile ="$($ticket)-$($uid)-$($Site.substring(0,3))-MsgTrk" ;
        #$ofile ="$($ticket)-MsgTrk" ;
        #2:39 PM 7/10/2024 revised for plus addressing
        if($Sender){
            if($Sender -match $rgxIsPlusAddrSmtpAddr){write-warning "WARNING! Sender $($Sender) HAS PLUS-ADDRESSING, WON'T WORK FOR EXOP RECIPIENTS!"} ; 
            #$pltGMTL.add("Sender",$Sender) ;
            $pltGMTL.add('Sender',($Sender -split ' *, *')) ;
            #$ofile+=",From-$($pltGMTL.Sender.replace("*","ANY"))"  ; 
        } ;
        if($Recipients){
            #$pltGMTL.add("Recipients",$Recipients) ;
            if($Recipients -match $rgxIsPlusAddrSmtpAddr){write-warning "WARNING! RecipientS $($Recipients) HAS PLUS-ADDRESSING, WON'T WORK FOR EXOP!"} ; 
            #$ofile+=",To-$($Recipients)" ;
            $pltGMTL.add('Recipients',($Recipients -split ' *, *')) ;
            #$ofile+=",To-$($pltGMTL.RecipientAddress.replace("*","ANY"))" ;
        } ;
        if($Start){
            $pltGMTL.add('Start',$Start) ; 
            #$ofile+= "-$(get-date $pltGMTL.Start -format $sFiletimestamp)-"  ;
        } ;
        if($End){
            $pltGMTL.add('End',$End) ; 
            #$ofile+= "$(get-date $pltGMTL.End -format $sFiletimestamp)" ;
        } ;
        if($MessageSubject){
            #$ofile+=",Subj-$($MessageSubject.substring(0,[System.Math]::Min(15,$MessageSubject.Length)))..." 
        } ;
        if($EventID){
            $pltGMTL.add('EventID',($EventID -split ' *, *')) ; 
            #$ofile+= "-Evt-$($EventID -join ',')" ;
        } ;
        if($MessageId){
            $pltGMTL.add('MessageId',($MessageId -split ' *, *')) ; 
            #$ofile+=",MsgId-$($pltGMTL.MessageId.replace('<','').replace('>',''))" ;
        } ;
        if($InternalMessageId){
            $pltGMTL.add("InternalMessageId",$InternalMessageId)  ;
            #$ofile+=",MsgID-$($InternalMessageId.replace('<','').replace('>','').substring(0,10))" ;
        } ;
        if($NetworkMessageId){
            $pltGMTL.add("NetworkMessageId",$NetworkMessageId)  ;
            #$ofile+=",MsgID-$($NetworkMessageId.replace('<','').replace('>','').substring(0,10))" ;
        } ;
         # Reference
        if($Reference){     $pltGMTL.add("Reference",$Reference)  ;
            #$ofile+=",Ref-$($Reference.replace('<','').replace('>','').substring(0,10))" ;
        } ;
        if($TransportTrafficType){     $pltGMTL.add("TransportTrafficType",$TransportTrafficType)  ;
            #$ofile+=",TTT-$(Remove-IllegalFileNameChars -Name $TransportTrafficType )" ;
        } ;
        if($Source){     $pltGMTL.add("Source",$Source)  ;
            #$ofile+=",Source-$($Source )" ;
        } ;
        if($Server){
            $pltGMTL.Server = $Server  ;
            if($Server -ne 'unlimited'){
                #$ofile+="-$($Server)" ;
            } ; 
        } ;
        if($ResultSize){
            $pltGMTL.ResultSize = $ResultSize  ;
            if($ResultSize -ne 'unlimited'){
                #$ofile+=",RSize-$($ResultSize)" ;
            } ; 
        } ;

        #region MSGTRKFILENAME ; #*------v MSGTRKFILENAME v------
        $LogPath = split-path $logfile ; 
        $smsg = "Writing export files to discovered `$LogPath: $($LogPath)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        if (-not (test-path $LogPath )){mkdir $LogPath -verbose  }
        [string[]]$ofile=@() ; 
        write-verbose "Add comma-delimited elements" ; 
        #$ofile+=if($ticket -AND $Tag){@($ticket,$tag) -join '_'}else{$ticket} ;
        $ofile+= (@($ticket,$tag) | ?{$_}) -join '_' ; 
        $ofile+= (@($Ten,$Requestor,'XOPMsgTrk') | ?{$_} ) -join '-' ;
        $ofile+=if($Sender){
            "FROM_$((($Sender | select -first 2) -join ',').replace('*','ANY'))"
        }else{''} ;
        $ofile+=if($Recipients){
            "TO_$(( ($Recipients| select -first 2) -join ',').replace('*','ANY'))"
        }else{''} ;
        $ofile+=if($MessageId){
            if($MessageId -is [array]){
                "MSGID_$($MessageId[0] -replace '[\<\>]','')..."
            } else { 
                "MSGID_$($MessageId -replace '[\<\>]','')"                
            } ; 
        }else{''} ;
        $ofile+=if($MessageSubject){"SUBJ_$($MessageSubject.substring(0,[System.Math]::Min(10,$MessageSubject.Length)))..."}else{''} ;
        $ofile+=if($EventID){
            "EVT_$($EventID -join ',')"
        }else{''} ;
        write-verbose "comma join the non-empty elements" ; 
        [string[]]$ofile=($ofile |  ?{$_} ) -join ',' ; 
        write-verbose "add the dash-delimited elements" ; 
        $ofile+=if($days){"$($days)d"}else{''} ;
        $ofile+=if($Start){"$(get-date $Start -format 'yyyyMMdd-HHmm')"}else{''} ;
        $ofile+=if($End){$ofile+= "$(get-date $End -format 'yyyyMMdd-HHmm')"}else{''} ;
        $ofile+=if($MessageSubject){"Subj_$($MessageSubject.replace("*"," ").replace("\"," "))"}else{''} ;
        $ofile+="run$(get-date -format 'yyyyMMdd-HHmm').csv" ;
        write-verbose "dash-join non-empty elems" ; 
        [string]$ofile=($ofile |  ?{$_} ) -join '-' ; 
        write-verbose "replace filesys illegal chars" ; 
        [string]$ofile=[RegEx]::Replace($ofile, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ;
        if($LogPath){
            write-verbose "add configured `LogPath" ; 
            $ofile = join-path $LogPath $ofile ; 
        } else { 
            write-verbose "add relative path" ; 
            $ofile=".\logs\$($ofile)" ;
        } ; 

        $hReports = [ordered]@{} ; 
        #rx10 ;
        $error.clear() ;

        TRY {

        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Running Get-MessageTrackingLog w`n$(($pltGMTL|out-string).trim())" ; 
        $Srvrs=Get-ExchangeServer ;
        if($Srvrs.name -contains $Site){
            write-verbose "Edge Role detected" ;
            $Srvrs=$Srvrs | where {$_.name -eq $Site -AND $_.IsEdgeServer} | select -expand Name ;
        }else{$Srvrs=($Srvrs | where { $_.isHubTransportServer -eq $true -and $_.Site -match ".*\/$($Site)$"} | select -expand Name)} ;

        #if($tag){$vn =  Remove-InvalidVariableNameChars -name "msgsOP$($ticket)_$($tag)"} else {$vn =  Remove-InvalidVariableNameChars -name "msgsOP$($ticket)" } ;
        #write-host "`$vn: $($vn)" ;
        #write-host -fore gray "collecting to variable: `$$($vn)" ;
        #if(gv $vn -ea 0){sv -name $vn -value $null } else {nv -name $vn -value $null } ;

        if($Server){
            $Msgs=($Server| get-messagetrackinglog @pltGMTL) | sort Timestamp ;
        } else { 
            $Msgs=($Srvrs| get-messagetrackinglog @pltGMTL) | sort Timestamp ;
        } ; 
        $smsg = "Raw matches:$(($msgs|measure).Count) events" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
        if($Connectorid){
            write-host -foregroundcolor gray   "Filtering on Conn:$($Connectorid)" ;
            $Msgs = $Msgs | ?{$_.connectorid -like $Connectorid} ;
            $ofile+="-conn-$($Connectorid.replace("*"," ").replace("\"," "))" ;
            write-host -foregroundcolor gray   "Post Conn filter matches:$(($Msgs|measure).Count)" ;
        } ;
        if($Source){
            write-host -foregroundcolor gray   "Filtering on Source:$($Source)" ;
            $Msgs = $Msgs | ?{$_.Source -like $Source} ;
            write-host -foregroundcolor gray   "Post Src filter matches:$(($Msgs|measure).Count)" ;
            $ofile+="-src-$($Source)" ;
        } ;

        if($MessageSubject){
            $smsg = "Post-Filtering on MessageSubject:$($MessageSubject)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            # detect whether to filter on -match (regex) or -like (asterisk, or default non-regex)
            if(test-IsRegexPattern -string $MessageSubject -verbose:$($VerbosePreference -eq "Continue")){
                $smsg = "(detected -MessageSubject as regex - using -match comparison)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $MsgsFltrd = $Msgs | ?{$_.MessageSubject -match $MessageSubject} ;
                if(-not $MsgsFltrd){
                    $smsg = "MessageSubject: regex -match comparison *FAILED* to return matches`nretrying MessageSubject filter as -Like..." ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $MsgsFltrd = $Msgs | ?{$_.MessageSubject -like $MessageSubject} ;
                } ; 
            } else { 
                $smsg = "(detected -MessageSubject as NON-regex - using -like comparison)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $MsgsFltrd = $Msgs | ?{$_.MessageSubject -like $MessageSubject} ;
                if(-not $MsgsFltrd){
                    $smsg = "MessageSubject: -like comparison *FAILED* to return matches`nretrying MessageSubject filter as -match..." ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $MsgsFltrd = $Msgs | ?{$_.MessageSubject -match $MessageSubject} 
                } ; 
            } ; 
            $smsg = "Post Subj filter matches:$(($MsgsFltrd|measure).Count)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            $msgs = $MsgsFltrd ; 
        } ;

        
        #if(get-variable -name $vn -scope global -ea 0){remove-variable -name $vn -scope global -force -ea 0} ; 
        #set-variable -name $vn -Value ($Msgs) -scope global ;
        
        if($Msgs){
            if($DoExports){
                $smsg = "($(($Msgs|measure).count) events | export-csv $($ofile))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                TRY{
                    $Msgs | SELECT $prpXCsv | EXPORT-CSV -notype -path $ofile ;
                    $smsg = "export-csv'd to:`n$((resolve-path $ofile).path)" ; 
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $smsg = "(adding `$hReports.MTMessages)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    
                # add the csvfilename
                $smsg = "(adding `$hReports.MTMessagesCSVFile:$($ofile))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $hReports.add('MTMessagesCSVFile',$ofile) ; 
            } ; 

            $hReports.add('MTMessages',$msgs) ; 

            $smsg = "`n`n#*------v EventID DISTRIB v------`n`n" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

            if($Msgs){
                $hReports.add('EventIDHisto',($Msgs | select -expand EventID | group | sort count,count -desc | select count,name)) ;

                #$smsg = "$(($Msgs | select -expand EventID | group | sort count,count -desc | select count,name|out-string).trim())" ;
                $smsg = "`n$(($hReports.EventIDHisto|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "`n`n#*------^ EventID DISTRIB ^------`n`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                $smsg = "`n`n#*------v MOST RECENT MATCH v------`n`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                $hReports.add('MsgLast',($msgs[-1]| fl $prpMTLfta)) ;
            

                $smsg = "`n$(($hReports.MsgLast |out-string).trim())";
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "`n`n#*------^ MOST RECENT MATCH ^------`n`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

            } ; 


            <#---
            write-host -foregroundcolor gray   "`n#*------v MOST RECENT MATCH v------" ;
            write-host "$(($msgs[-1]| format-list $prpMTLfta|out-string).trim())";
            write-host -foregroundcolor gray   "`n#*------^ MOST RECENT MATCH ^------`n" ;
            write-host -foregroundcolor gray   "`n#*------v EventID DISTRIB v------" ;
            write-host -foregroundcolor yellow "$(($Msgs | group EventID | sort count -desc | ft -a count,name |out-string).trim())";
            write-host -foregroundcolor gray   "`n#*------^ EventID DISTRIB ^------" ;
            #---
            #>

            if($mFails = $msgs | ?{$_.EventID -eq 'FAIL'}){
                $hReports.add('MsgsFail',$mFails) ; 
                $ofileF = $ofile.replace('-XOPMsgTrk,','FAILMsgs,') ;
                if($DoExports){
                    TRY{
                        $mFails | export-csv -notype -path $ofileF -ea STOP ;
                        $smsg = "export-csv'd to:`n$((resolve-path $ofileF).path)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 
                } ; 
                if($mOOO = $mFails | ?{$_.Subject -match '^Automatic\sreply:\s'}){
                    $smsg = $sBnr3="`n#*~~~~~~v EventID FAIL: Expected Policy Blocked External OutOfOffice v~~~~~~" ;
                    $smsg += "`n$($mOOO| measure | select -expand count) msgs:Expected Out-of-Office Policy:(attempt to send externally)`n$(($mOOO| ft -a $prpGXMTfta |out-string).trim())" ;
                    $smsg += "`n$($sBnr3.replace('~v','~^').replace('v~','^~'))`n"  ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ;
                if($mRecl = $mFails | ?{$_.Subject -match '^Recall:\s'}){
                    $smsg = $sBnr3="`n#*~~~~~~v EventID FAIL: Expected: Recalled message v~~~~~~" ;
                    $smsg += "`n$($mRecl| measure | select -expand count) msgs:Expected Sender Recalled Message `n$(($mRecl| ft -a $prpGXMTfta |out-string).trim())" ;
                    $smsg += "`n$($sBnr3.replace('~v','~^').replace('v~','^~'))`n"  ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ;
                if($mFails = $mFails | ?{$_.Subject -notmatch '^Recall:\s' -AND $_.Subject -notmatch '^Automatic\sreply:\s'}){
                    $smsg = $sBnr3="`n#*~~~~~~v EventID FAIL: Other Failure message v~~~~~~" ;
                    $smsg += "`n$(($mFails | ft -a |out-string).trim())" ;
                    $smsg += "`n$($sBnr3.replace('~v','~^').replace('v~','^~'))`n"  ;
                    write-host -foregroundcolor yellow $smsg ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ;
                if($mFailAddr = $mFails | ?{$_.EventID -eq 'FAIL' -AND $_.Source -eq 'ROUTING'}){
                    $smsg = $sBnr3="`n#*~~~~~~v EventID FAIL: BAD ADDRESS FAILS: (EventID:'FAIL' & Source:'ROUTING') ($(($mFailAddr|measure).count)msgs) v~~~~~~" ;
                    $smsg += "`n$(($mFailAddr | ft -a |out-string).trim())" ;
                    $smsg += "`n$($sBnr3.replace('~v','~^').replace('v~','^~'))`n"  ;
                    write-host -foregroundcolor yellow $smsg ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ;
                if($mFailRej = $mFails | ?{$_.EventID -eq 'FAIL' -AND $_.Source -eq 'SMTP'}){
                    $smsg = $sBnr3="`n#*~~~~~~v EventID FAIL: REJECTED BY RECIPIENT SERVER : (EventID:'FAIL' & Source:'SMTP') ($(($mFailRej|measure).count)msgs) v~~~~~~" ;
                    $smsg += "`n$(($mFailRej | ft -a |out-string).trim())" ;
                    $smsg += "`n$($sBnr3.replace('~v','~^').replace('v~','^~'))`n"  ;
                    write-host -foregroundcolor yellow $smsg ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ;
                if($mFailDom = $mFails | ?{$_.EventID -eq 'FAIL' -AND $_.Source -eq 'DNS'}){
                    $smsg = $sBnr3="`n#*~~~~~~v EventID FAIL: BAD DOMAIN : (EventID:'FAIL' & Source:'DNS') ($(($mFailDom|measure).count)msgs) v~~~~~~" ;
                    $smsg += "`n$(($mFailDom | ft -a |out-string).trim())" ;
                    $smsg += "`n$($sBnr3.replace('~v','~^').replace('v~','^~'))`n"  ;
                    write-host -foregroundcolor yellow $smsg ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ;

            } ;

            if($mDefers = $msgs|?{$_.EventID -eq 'DEFER'}){
                $smsg = "`n`n#*------v DEFER's Distribution v------`n`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                
                $hReports.add('MsgsDefer',$mDefers) ;

                $smsg = "`n$(($mDefers | select -expand RecipientStatus | group | sort count -desc | ft -auto count,name|out-string).trim())";
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "`n`n#*------^ DEFER's Distribution ^------`n`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 


            } else {
                $smsg = "(no DEFERs logged)" 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;

            
            if($mFails){
                $smsg = "`n`n#*------v FAIL's RecipientStatus Distribution v------`n`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                $hReports.add('MsgsFailRcpStatHisto',($fails| select -expand RecipientStatus | group | sort count -desc | select count,name)) ;

                $smsg = "$(($hReports.MsgsFailRcpStat | ft -auto |out-string).trim())";
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                $smsg = "`n`n#*------^ FAIL's RecipientStatus Distribution ^------`n`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                if(($msgs|?{$_.EventID -eq 'FAIL'}).count -lt 20){
                    $smsg = "`n`n#*------v FAIL Details v------`n`n" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = "$(($msgs|?{$_.EventID -eq 'FAIL'} | fl $prpMTFailFL|out-string).trim())";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = "`n`n#*------^ FAIL Details ^------`n`n" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } else {
                    $smsg = "(more than 20 FAIL:not echoing to console)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                };
            } else {
                $smsg = "(no FAILs logged)" 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;

            $msgs = $null ;
            if(test-path -path $ofile){  
                write-host -foregroundcolor green  "(log file confirmed)" ;
                #Resolve-Path -Path $ofile | select -expand Path | out-clipboard ;
                $smsg = "`n(Tracked output file confirmed)" ;
                write-host -fore green $smsg ;
            } else { write-warning "MISSING LOG FILE!" } ;
        } else {
            $smsg = "NO MATCHES FOUND From Qry:" ; 
            $smsg += "`n$(($pltGMTL|out-string).trim())" ; 
            write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
        } ;
    } CATCH {
        $ErrTrapd=$Error[0] ;
        $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
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

        write-verbose "#*------^ END Function SUB MAIN ^------" ;
    } ;  # BEG-E
    END {
        if($SimpleTrack -AND ($hReports.Keys.Count -gt 0)){
            $smsg = "-SimpleTrack specified: Only returning net message tracking set to pipeline" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            $msgs | write-output ; 
        } else { 
            $smsg = "(no -SimpleTrack: returning full summary object to pipeline)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            if($hReports.Keys.Count -gt 0){
                # convert the hashtable to object for output to pipeline
                #$Rpt += New-Object PSObject -Property $hReports ;
                $smsg = "(Returning summary object to pipeline)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                
                TRY{
                    New-Object -TypeName PsObject -Property $hReports | write-output ; 
                    # export the entire return object into xml
                    $smsg = "(exporting `$hReports summary object to xml:$($ofile.replace('.csv','.xml')))" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    New-Object -TypeName PsObject -Property $hReports | export-clixml -path $ofile.replace('.csv','.xml') -ea STOP -verbose
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 

            } else { 
                $smsg = "Unpopulated `$hReports, skipping output to pipeline" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARNING } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                $false | write-output ; 
            } ;  
        } ; 
    } ; 
}

#*------^ Get-MessageTrackingLogTDO.ps1 ^------


#*------v get-UserMailADSummary.ps1 v------
function get-UserMailADSummary {
    <#
    .SYNOPSIS
    get-UserMailADSummary.ps1 - Resolve specified array of -users (displayname, emailaddress, samaccountname) to mail asset and AD details
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-07-30
    FileName    : get-UserMailADSummary.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell,Exchange,ExchangeOnline,ActiveDirectory
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 2:16 PM 6/24/2024: rem'd out #Requires -RunasAdministrator; sec chgs in last x mos wrecked RAA detection
    * 3:27 PM 3/8/2022 CBH update, pointed link at verb-ex2010 module repo
    * 10:57 AM 8/2/2021 add 'Unresolved' (failed to return an object from 
    get-recipient, normally a 'System.Management.Automation.RemoteException', 
    Exception -match "couldn't\sbe\sfound\son") & 'Failed' (catch failure that 
    doesn't match the prior) properties, stuffing in the input designators as a raw 
    array (indicates cloud objects hybrid sync'd to another external directory) 
    * 12:07 PM 7/30/2021 added CustAttribs to the props dumped ; pulled 'verb-Ex2010' from #requires (nesting limit) ; init
    .DESCRIPTION
    get-UserMailADSummary.ps1 - Resolve specified array of -users (displayname, emailaddress, samaccountname) to mail asset and AD details
    The specific goal of this function is to assist in differentiated 'Active' user mailboxes from termainted/disabled/offboarded user mailboxes (or (Shared|Room|Equpment)mailboxes with disabled ADUser accounts.
    
    Fed an array of mailbox descriptors - emailaddress/UPN, alias, displayname, etc) - the function attempts to resolve a local recipient object, and a matching local ADUser object. 

    The ADUser is evaluated for .enabled status, and the distinguishedName is checked to locate 'non-Active/Term' users (by OU name in their DN). 
    
    The resulting lookups are categorized into four properties of the returned SystemObject:
    - Enabled - resolved local recipient & local ADUser. ADUser.Enabled=$true, ADUser.distinguishedName does not include a 'Disabled*' or 'Termed*' OU in it's tree.
    - Disabled - resolved local recipient, & local ADUser. ADUser.Enabled=$false and/or ADUser.distinguishedName includes a 'Disabled*' or 'Termed*' OU in it's tree.
    - UnResolved - get-recipient failed to return a local recipient matching the specified mailbox designator (most frequently reflects EXO cloud mailboxes that are hybrid-federated out of an other external AD as Source-Of-Authority.
    - Failed - any get-recipient try/catch fail that doesn't appear to be Exception.Type 'System.Management.Automation.RemoteException', Exception match "couldn't\sbe\sfound\son"

    The resulting SystemObject from the above is returned to the pipeline.
    .PARAMETER  users
    Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)
    .PARAMETER ADDisabledOnly
    Switch to exclude users solely on ADUser.disabled (not Disabled OU presense), or with that have the ADUser below an OU matching '*OU=(Disabled|TERMedUsers)'  [-ADDisabledOnly]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    Returns a System.Object[] object to pipeline, with subsets of processed users as 'Enabled' (ADUser.enabled),'Disabled', and 'Contacts' properties. 
    .EXAMPLE
    PS> $rpt = get-UserMailADSummary -users 'username1','user2@domain.com','[distinguishedname]' ;
    PS> $rpt | export-csv -nottype ".\pathto\usersummaries-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ; 
    Processes local/remotemailbox & ADUser details on three specified users (alias,email address, DN). Summaries are returned to pipeline, and assigned to the $rpt variable, which is then exported to csv.
    .EXAMPLE
    PS> $rpt = get-UserMailADSummary -users 'username1','user2@domain.com','[distinguishedname]' -ADDisabledOnly ;
    PS> $rpt | export-csv -nottype ".\pathto\usersummaries-ENABLEDUSERS-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ; 
    Processes local/remotemailbox & ADUser details on three specified users (alias,email address, DN). 
    And allocate as 'Disabled', accounts that are *solely* ADUser.disabled 
    (e.g. considers users below OU's with names like 'OU=Disabled*' as 'Enabled' users), 
    and then exports to csv. 
    .EXAMPLE
    $rpt = get-UserMailADSummary -users 'username1','user2@domain.com','[distinguishedname]' ;
    $rpt.enabled | export-csv -nottype ".\pathto\usersummaries-ENABLEDUSERS-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
    Process specified identifiers, and export solely the 'Enabled' users returned to csv. 
    .LINK
    https://github.com/tostka/verb-ex2010
    .LINK
    #>
    ###Requires -Version 5
    #Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS
    ##Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    ##[Alias('ulu')]
    PARAM(
        [Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)")]
        #[ValidateNotNullOrEmpty()]
        #[Alias('ALIAS1', 'ALIAS2')]
        [array]$users,
        [Parameter(HelpMessage="Switch to exclude users solely on ADUser.disabled (not Disabled OU presense), or with that have the ADUser below an OU matching '*OU=(Disabled|TERMedUsers)'[-ADDisabledOnly]")]
        [switch] $ADDisabledOnly,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2
    ) ;
    BEGIN{
        $Verbose = ($VerbosePreference -eq 'Continue') ; 

        $rgxDisabledOUs = '.*OU=(Disabled|TERMedUsers).*' ; 
        if(!$users){
            $users= (get-clipboard).trim().replace("'",'').replace('"','') ; 
            if($users){
                write-verbose "No -users specified, detected value on clipboard:`n$($users)" ; 
            } else { 
                write-warning "No -users specified, nothing matching dname, emailaddress or samaccountname, found on clipboard. EXITING!" ; 
                Break ; 
            } ; 
        } ; 
        $ttl = ($users|measure).count ; 
        if($ttl -lt 10){
            write-verbose "($(($users|measure).count)) user(s) specified:`n'$($users -join "','")'" ;    
        } else { 
            write-verbose "($(($users|measure).count)) user(s) specified (-gt 10, suppressing details)" ;    
        } 

        rx10 -Verbose:$false ; #rxo  -Verbose:$false ; #cmsol  -Verbose:$false ;
        connect-ad -Verbose:$false; 
        
        $propsmbx='Database','UseDatabaseRetentionDefaults','SingleItemRecoveryEnabled','RetentionPolicy','ProhibitSendQuota',
            'ProhibitSendReceiveQuota','SamAccountName','ServerName','UseDatabaseQuotaDefaults','IssueWarningQuota','Office',
            'UserPrincipalName','Alias','OrganizationalUnit  global.ad.toro.com/TERMedUsers','DisplayName','EmailAddresses',
            'HiddenFromAddressListsEnabled','LegacyExchangeDN','PrimarySmtpAddress','RecipientType','RecipientTypeDetails',
            'WindowsEmailAddress','DistinguishedName','CustomAttribute1','CustomAttribute2','CustomAttribute3','CustomAttribute4',
            'CustomAttribute5','CustomAttribute6','CustomAttribute7','CustomAttribute8','CustomAttribute9','CustomAttribute10',
            'CustomAttribute11','CustomAttribute12','CustomAttribute13','CustomAttribute14','CustomAttribute15''EmailAddressPolicyEnabled',
            'WhenChanged','WhenCreated' ;
        $propsadu = "accountExpires","CannotChangePassword","Company","Compound","Country","countryCode","Created","Department",
            "Description","DisplayName","DistinguishedName","Division","EmployeeID","EmployeeNumber","employeeType","Enabled","Fax",
            "GivenName","homeMDB","homeMTA","info","Initials","lastLogoff","lastLogon","LastLogonDate","mail","mailNickname","Manager",
            "mobile","MobilePhone","Modified","Name","Office","OfficePhone","Organization","physicalDeliveryOfficeName","POBox","PostalCode",
            "SamAccountName","sAMAccountType","State","StreetAddress","Surname","Title","UserPrincipalName",'CustomAttribute1',
            'CustomAttribute2','CustomAttribute3','CustomAttribute4','CustomAttribute5','CustomAttribute6','CustomAttribute7',
            'CustomAttribute8','CustomAttribute9','CustomAttribute10','CustomAttribute11','CustomAttribute12','CustomAttribute13',
            'CustomAttribute14','CustomAttribute15','EmailAddressPolicyEnabled',"whenChanged","whenCreated" ;
        $propsMC = 'ExternalEmailAddress','Alias','DisplayName','EmailAddresses','PrimarySmtpAddress','RecipientType',
            'RecipientTypeDetails','WindowsEmailAddress','Name','DistinguishedName','Identity','CustomAttribute1','CustomAttribute2',
            'CustomAttribute3','CustomAttribute4','CustomAttribute5','CustomAttribute6','CustomAttribute7','CustomAttribute8',
            'CustomAttribute9','CustomAttribute10','CustomAttribute11','CustomAttribute12','CustomAttribute13','CustomAttribute14',
            'CustomAttribute15','EmailAddressPolicyEnabled','whenChanged','whenCreated' ;
    } 
    PROCESS{
        $Procd=0 ;$pct = 0 ; 
        $aggreg =@() ; $contacts =@() ; $UnResolved = @() ; $Failed = @() ;
        $pltGRcp=[ordered]@{identity=$null;erroraction='STOP';resultsize=1;} ; 
        $pltGMbx=[ordered]@{identity=$null;erroraction='STOP'} ; 
        $pltGadu=[ordered]@{Identity = $null ; Properties='*' ;errorAction='STOP'} ;
        foreach ($usr in $users){
            $procd++ ; $pct = '{0:p0}' -f ($procd/$ttl) ; 
            $rrcp = $mbx = $mc = $mbxspecs = $adspecs = $summary = $NULL ; 
            #write-verbose "processing:$($usr)" ; 
            $sBnrS="`n#*------v PROCESSING ($($procd)/$($ttl)):$($usr)`t($($pct)) v------" ; 
            if($verbose){
                write-verbose "$((get-date).ToString('HH:mm:ss')):$($sBnrS)" ;
            } else { 
                write-host "." -NoNewLine ; 
            } ; 
            
            TRY {
                $pltGRcp.identity = $usr ; 
                write-verbose "get-recipient  w`n$(($pltGRcp|out-string).trim())" ; 
                $rrcp = get-recipient @pltGRcp ;
                if($rrcp){
                    $pltgmbx.identity = $rrcp.PrimarySmtpAddress ; 
                    switch ($rrcp.recipienttype){
                        'MailUser'{
                            write-verbose "get-remotemailbox  w`n$(($pltgmbx|out-string).trim())" ; 
                            $mbx = get-remotemailbox @pltgmbx 
                        } 
                        'UserMailbox' {
                            write-verbose "get-mailbox w`n$(($pltgmbx|out-string).trim())" ; 
                            $mbx = get-mailbox @pltgmbx ;
                        }
                        'MailContact' {
                            write-verbose "get-mailcontact w`n$(($pltgmbx|out-string).trim())" ; 
                            $mc = get-mailcontact @pltgmbx ;
                        }
                        default {throw "$($rrcp.alias):Unsupported RecipientType:$($rrcp.recipienttype)" }
                    } ; 
                    if(-not($mc)){
                        $mbxspecs =  $mbx| select $propsmbx ;
                        $pltGadu.identity = $mbx.samaccountname ; 
                        write-verbose "Get-ADUser w`n$(($pltGadu|out-string).trim())" ; 
                        Try {
                            $adspecs =Get-ADUser @pltGadu | select $propsadu ;
                        } CATCH [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                            $smsg = "(no matching ADuser found:$($pltGadu.identity))" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Continue ;
                        } catch {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Continue ;
                        } ; 
                        $summary = @{} ;
                        foreach($object_properties in $mbxspecs.PsObject.Properties) {
                            $summary.add($object_properties.Name,$object_properties.Value) ;
                        } ;
                        foreach($object_properties in $adspecs.PsObject.Properties) {
                            $summary.add("AD$($object_properties.Name)",$object_properties.Value) ;
                        } ;
                        $aggreg+= New-Object PSObject -Property $summary ;

                    } else { 
                        $smsg = "Resolved user for $($usr) is RecipientType:$($mc.RecipientType)`nIt is not a local mail object, or AD object, and simply reflects a pointer to an external mail recipient.`nThis object is being added to the 'Contacts' section of the output.." ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $contacts += $mc | select $propsMC ;
                    } ; 
                } else { 
                    # in ISE $Error[0] is empty
                    #$ErrTrapd=$Error[0] ;
                    #if($ErrTrapd.Exception -match "couldn't\sbe\sfound\son"){
                        $UnResolved += $pltGRcp.identity ;
                    #} else { 
                        #$Failed += $pltGRcp.identity ;
                    #} ; 
                } ; 
            } CATCH [System.Management.Automation.RemoteException] {
                # catch error never gets here (at least not in ISE)
                $ErrTrapd=$Error[0] ;
                if($ErrTrapd.Exception -match "couldn't\sbe\sfound\son"){
                    $UnResolved += $pltGRcp.identity ;
                } else { 
                    $Failed += $pltGRcp.identity ;
                } ;                 
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $Failed += $pltGRcp.identity ;
            } ; 

            if($verbose){
                write-verbose "$((get-date).ToString('HH:mm:ss')):$($sBnrS.replace('-v','-^').replace('v-','^-'))" ; ;
            } ; 
            
        } ; 
    }
    END{
        if(-not($ADDisabledOnly)){
            $Report = [ordered]@{
                Enabled = $Aggreg|?{($_.ADEnabled -eq $true ) -AND -not($_.distinguishedname -match $rgxDisabledOUs) } ;#?{$_.adDisabled -ne $true -AND -not($_.distinguishedname -match $rgxDisabledOUs)}
                Disabled = $Aggreg|?{($_.ADEnabled -eq $False) } ; 
                Contacts = $contacts ; 
                Unresolved = $Unresolved ; 
                Failed = $Failed;
            } ; 
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):outputing $(($Report.Enabled|measure).count) Enabled User summaries,`nand $(($Report.Disabled|measure).count) ADUser.Disabled or Disabled/TERM-OU account summaries`nand $(($Report.Contacts|measure).count) users resolved to MailContacts" ; 
        } else { 
            $Report = [ordered]@{
                Enabled = $Aggreg|?{($_.ADEnabled -eq $true) -AND -not($_.distinguishedname -match $rgxDisabledOUs) } ;#?{$_.adDisabled -ne $true -AND -not($_.distinguishedname -match $rgxDisabledOUs)}
                Disabled = $Aggreg|?{($_.ADEnabled -eq $False) -OR ($_.distinguishedname -match $rgxDisabledOUs) } ; 
                Contacts = $contacts ; 
                Unresolved = $Unresolved ; 
                Failed = $Failed;
            } ; 
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):outputing $(($Report.Enabled|measure).count) Enabled User summaries,`nand $(($Report.Disabled|measure).count) ADUser.Disabled`nand $(($Report.Contacts|measure).count) users resolved to MailContacts" ; 
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):-ADDisabledOnly specified: 'Disabled' output are *solely* ADUser.Disabled (no  Disabled/TERM-OU account filtering applied)`nand $(($Report.Contacts|measure).count) users resolved to MailContacts" ; 
        } ; 
        New-Object PSObject -Property $Report | write-output ;
        
     } ; 
 }

#*------^ get-UserMailADSummary.ps1 ^------


#*------v import-EMSLocalModule.ps1 v------
Function import-EMSLocalModule {
  <#
    .SYNOPSIS
    import-EMSLocalModule - Setup local server bin-module-based ExchOnPrem Mgmt Shell connection (contrasts with Ex2007/10 snapin use ; validated Exch2016)
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    Version     : 1.0.0
    CreatedDate : 2021-04-15
    FileName    : import-EMSLocalModule()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell,Exchange,Exchange-2013,Exchange-2016
    REVISIONS   :
    * 9:48 AM 7/27/2021 added verbose to -pstitlebar
    * 11:40 AM 5/14/2021 added -ea 0 to the gv tests (suppresses not-found error when called without logging config)
    * 9:21 AM 4/16.2.31 renamed load-emsmodule -> import-EMSLocalModule, added pretest and post verify
    * 10:14 AM 4/12/2021 init vers
    .DESCRIPTION
    import-EMSLocalModule - Setup local server bin-module-based ExchOnPrem Mgmt Shell connection (contrasts with Ex2007/10 snapin use ; validated Exch2016)
    Wraps the native ex server desktop EMS, .lnk commands:
    . 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1';
    Connect-ExchangeServer -auto -ClientApplication:ManagementShell
    Handy for loading local non-dehydrated support in ISE, regular PS etc, where existing code relies on non-dehydrated objects.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    import-EMSLocalModule ;
    .LINK
    https://github.com/tostka/verb-Ex2010/
    #>
    [CmdletBinding()]
    #[Alias()]
    Param()  ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ;
    } ;  # BEG-E
    PROCESS{
        $error.clear() ;
        TRY {
            if($tMod = get-module ([System.Net.Dns]::gethostentry($(hostname))).hostname -ea 0){
                write-verbose "(local EMS module already loaded:)`n$(($tMod | ft -auto ModuleType,Version,Name,ExportedCommands|out-string).trim())" ;
            } else {
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(importing local Exchange Mgmt Shell binary module)" ;
                if($env:ExchangeInstallPath){
                    $rExps1 = "$($env:ExchangeInstallPath)bin\RemoteExchange.ps1"
                    if(test-path $rExps1){
                        . $rExps1 ;
                        if(Get-Command Connect-ExchangeServer){
                            Connect-ExchangeServer -auto -ClientApplication:ManagementShell ;
                        } else {
                            throw "Unable to gcm Connect-ExchangeServer!" ;
                        } ;
                    } else {
                        throw "Unable to locate: `$(`$env:ExchangeInstallPath)bin\RemoteExchange.ps1" ;
                    } ;
                } else {
                    throw "Unable to locate: `$env:ExchangeInstallPath Environment Variable (Exchange does not appear to be locally installed)" ;
                } ;
            } ;
        } CATCH {
            $ErrTrapd = $_ ;
            write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            #-=-record a STATUSERROR=-=-=-=-=-=-=
            $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(Get-Variable passstatus -scope Script -ea 0 ){$script:PassStatus += $statusdelta } ;
            if(Get-Variable -Name PassStatus_$($tenorg) -scope Script  -ea 0 ){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
            #-=-=-=-=-=-=-=-=
        } ;
        # 7:54 AM 11/1/2017 add titlebar tag
        if(Get-Command Add-PSTitleBar -ea 0 ){Add-PSTitleBar 'EMSL' -verbose:$($VerbosePreference -eq "Continue");} ;
        # tag E10IsDehydrated
        $Global:ExOPIsDehydrated = $false ;
    } ;  # PROC-E
    END {
        $tMod = $null ;
        if($tMod = GET-MODULE ([System.Net.Dns]::gethostentry($(hostname))).HOSTNAME){
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(local EMS module loaded:)`n$(($tMod | ft -auto ModuleType,Version,Name,ExportedCommands|out-string).trim())" ;
        } else {
            throw "Unable to resolve target local EMS module:GET-MODULE $([System.Net.Dns]::gethostentry($(hostname))).HOSTNAME)" ;
        } ;
    }
}

#*------^ import-EMSLocalModule.ps1 ^------


#*------v Invoke-ExchangeCommand.ps1 v------
function Invoke-ExchangeCommand{
    <#
    .SYNOPSIS
    Invoke-ExchangeCommand.ps1 - PowerShell function allows you to run PowerShell commands and script blocks on Exchange servers in different forests without a forest trust.
    .NOTES
    Version     : 0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2020-09-15
    FileName    : IInvoke-ExchangeCommand.ps1
    License     : MIT License
    Copyright   : (c) 2015 Mark Gossa
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell,Exchange,ExchangeForestMigration,CrossForest,ExchangeRemotePowerShell,ExchangePowerShell
    AddedCredit : Mark Gossa
    AddedWebsite: https://gallery.technet.microsoft.com/Exchange-Cross-Forest-e25d48eb
    AddedTwitter:
    REVISIONS
    * 2:40 PM 9/17/2020 cleanup <?> encode damage (emdash's for dashes)
    * 4:28 PM 9/15/2020 cleanedup, added CBH, added to verb-Ex2010
    * 10/26/2015 posted vers
    .DESCRIPTION
    Invoke-ExchangeCommand.ps1 - PowerShell function allows you to run PowerShell commands and script blocks on Exchange servers in different forests without a forest trust.
    This PowerShell function allows you to run PowerShell commands and script blocks on Exchange servers in different forests without a forest trust.
    To run Invoke-ExchangeCommand, you must connect to the Exchange server using a hostname and not an IP address. Invoke-ExchangeCommand works best on Server 2012 R2/Windows 8.1 and later but also works on Server 2008 R2/Windows 7. Tested on Exchange 2010 and later. More information on cross-forest Exchange PowerShell can be found here: http://markgossa.blogspot.com/2015/10/exchange-2010-2013-cross-forest-remote-powershell.html
    Usage:
    1. Enable connections to all PowerShell hosts:
    winrm s winrm/config/client '@{TrustedHosts="*"}'
    # TSK: OR BETTER: _SELECTIVE_ HOSTS:
    Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value 'LYNMS7330.global.ad.toro.com' -Concatenate -Force ; Get-Item -Path WSMan:\localhost\Client\TrustedHosts | fl Name, Value ;
    cd WSMan:\localhost\Client ;
    dir | format-table -auto ; # review existing settings:
    # AllowEncrypted is defined on the client end, via the WSMAN: drive
    set-item .\allowunencrypted $true ;
    # You probably will need to set the AllowUnencrypted config setting in the Service as well, which has to be changed in the remote server using the following:
    set-item -force WSMan:\localhost\Service\AllowUnencrypted $true ;
    # tsk: reverted it back out:
    #-=-=-=-=-=-=-=-=
    [PS] WSMan:\localhost\Service\Auth> set-item -force WSMan:\localhost\Service\AllowUnencrypted $false ;
    cd ..
    [PS] WSMan:\localhost\Service>dir | format-table -auto ;
       WSManConfig: Microsoft.WSMan.Management\WSMan::localhost\Service
    Type          Name                             SourceOfValue Value
    ----          ----                             ------------- -----
    System.String RootSDDL                                       O:NSG:BAD:P(A;;GA;;;BA)(A;;GR;;;IU)S:P(AU;FA;GA;;;WD)(A...
    System.String MaxConcurrentOperations                        4294967295
    System.String MaxConcurrentOperationsPerUser                 1500
    System.String EnumerationTimeoutms                           240000
    System.String MaxConnections                                 300
    System.String MaxPacketRetrievalTimeSeconds                  120
    System.String AllowUnencrypted                               false
    Container     Auth
    Container     DefaultPorts
    System.String IPv4Filter                                     *
    System.String IPv6Filter                                     *
    System.String EnableCompatibilityHttpListener                false
    System.String EnableCompatibilityHttpsListener               false
    System.String CertificateThumbprint
    System.String AllowRemoteAccess                              true
    #-=-=-=-=-=-=-=-=
    TSK: try it *without* AllowUnencrypted before opening it up
    # And don't forget to also enable Digest Authorization:
    set-item -force WSMan:\localhost\Service\Auth\Digest $true ;
    # (to allow the system to digest the new settings)
    TSK: I don't even see the path existing on the lab Ex651
    WSMan:\localhost\Service\Auth\Digest
    TSK: but winrm shows the config enabled with Digest:
    winrm get winrm/config/client
    #-=-=-=-=-=-=-=-=
    Client
      NetworkDelayms = 5000
      URLPrefix = wsman
      AllowUnencrypted = true
      Auth
          Basic = true
          Digest = true
          Kerberos = true
          Negotiate = true
          Certificate = true
          CredSSP = false
      DefaultPorts
          HTTP = 5985
          HTTPS = 5986
      TrustedHosts = LYNMS7330
    #-=-=-=-=-=-=-=-=
    #-=-=L650'S settings-=-=-=-=-=-=
    # SERVICE AUTH
    [PS] C:\scripts>winrm get winrm/config/service/auth
    Auth
        Basic = false
        Kerberos = true
        Negotiate = true
        Certificate = false
        CredSSP = false
        CbtHardeningLevel = Relaxed
    # SERVICE OVERALL
    [PS] C:\scripts>winrm get winrm/config/service
    Service
    RootSDDL = O:NSG:BAD:P(A;;GA;;;BA)S:P(AU;FA;GA;;;WD)(AU;SA;GWGX;;;WD)
    MaxConcurrentOperations = 4294967295
    MaxConcurrentOperationsPerUser = 15
    EnumerationTimeoutms = 60000
    MaxConnections = 25
    MaxPacketRetrievalTimeSeconds = 120
    AllowUnencrypted = false
    Auth
        Basic = false
        Kerberos = true
        Negotiate = true
        Certificate = false
        CredSSP = false
        CbtHardeningLevel = Relaxed
    DefaultPorts
        HTTP = 5985
        HTTPS = 5986
    IPv4Filter = *
    IPv6Filter = *
    EnableCompatibilityHttpListener = false
    EnableCompatibilityHttpsListener = false
    CertificateThumbprint
    #-=-=-=-=-=-=-=-=
    ==3:22 PM 9/17/2020:POST settings on CurlyHoward:
    #-=-=-=-=-=-=-=-=
    [PS] WSMan:\localhost\Client>winrm get winrm/config/client
    Client
        NetworkDelayms = 5000
        URLPrefix = wsman
        AllowUnencrypted = true
        Auth
            Basic = true
            Digest = true
            Kerberos = true
            Negotiate = true
            Certificate = true
            CredSSP = false
        DefaultPorts
            HTTP = 5985
            HTTPS = 5986
        TrustedHosts = LYNMS7330
    [PS] WSMan:\localhost\Client>winrm get winrm/config/client
    Client
        NetworkDelayms = 5000
        URLPrefix = wsman
        AllowUnencrypted = true
        Auth
            Basic = true
            Digest = true
            Kerberos = true
            Negotiate = true
            Certificate = true
            CredSSP = false
        DefaultPorts
            HTTP = 5985
            HTTPS = 5986
        TrustedHosts = LYNMS7330
    [PS] WSMan:\localhost\Client>winrm get winrm/config/service/auth
    Auth
        Basic = false
        Kerberos = true
        Negotiate = true
        Certificate = false
        CredSSP = false
        CbtHardeningLevel = Relaxed
    [PS] WSMan:\localhost\Client>winrm get winrm/config/service
    Service
        RootSDDL = O:NSG:BAD:P(A;;GA;;;BA)(A;;GR;;;IU)S:P(AU;FA;GA;;;WD)(AU;SA;GXGW;;;WD)
        MaxConcurrentOperations = 4294967295
        MaxConcurrentOperationsPerUser = 1500
        EnumerationTimeoutms = 240000
        MaxConnections = 300
        MaxPacketRetrievalTimeSeconds = 120
        AllowUnencrypted = true
        Auth
            Basic = false
            Kerberos = true
            Negotiate = true
            Certificate = false
            CredSSP = false
            CbtHardeningLevel = Relaxed
        DefaultPorts
            HTTP = 5985
            HTTPS = 5986
        IPv4Filter = *
        IPv6Filter = *
        EnableCompatibilityHttpListener = false
        EnableCompatibilityHttpsListener = false
        CertificateThumbprint
        AllowRemoteAccess = true
    #-=-=-=-=-=-=-=-=
    #-=-ABOVE SETTINGS VIA WSMAN: PSDRIVE=-=-=-=-=-=-=
    [PS] WSMan:\localhost\Client>cd WSMan:\localhost\Client ;
    [PS] WSMan:\localhost\Client>dir | format-table -auto ;
       WSManConfig: Microsoft.WSMan.Management\WSMan::localhost\Client
    Type          Name             SourceOfValue Value
    ----          ----             ------------- -----
    System.String NetworkDelayms                 5000
    System.String URLPrefix                      wsman
    System.String AllowUnencrypted               true
    Container     Auth
    Container     DefaultPorts
    System.String TrustedHosts                   LYNMS7330

    [PS] WSMan:\localhost\Client>cd WSMan:\localhost\Service
    [PS] WSMan:\localhost\Service>dir | format-table -auto ;
       WSManConfig: Microsoft.WSMan.Management\WSMan::localhost\Service
    Type          Name                             SourceOfValue Value
    ----          ----                             ------------- -----
    System.String RootSDDL                                       O:NSG:BAD:P(A;;GA;;;BA)(A;;GR;;;IU)S:P(AU;FA;GA;;;WD)(A...
    System.String MaxConcurrentOperations                        4294967295
    System.String MaxConcurrentOperationsPerUser                 1500
    System.String EnumerationTimeoutms                           240000
    System.String MaxConnections                                 300
    System.String MaxPacketRetrievalTimeSeconds                  120
    System.String AllowUnencrypted                               true
    Container     Auth
    Container     DefaultPorts
    System.String IPv4Filter                                     *
    System.String IPv6Filter                                     *
    System.String EnableCompatibilityHttpListener                false
    System.String EnableCompatibilityHttpsListener               false
    System.String CertificateThumbprint
    System.String AllowRemoteAccess                              true

    [PS] WSMan:\localhost\Service>cd WSMan:\localhost\Service\Auth\
    [PS] WSMan:\localhost\Service\Auth>dir | format-table -auto ;
       WSManConfig: Microsoft.WSMan.Management\WSMan::localhost\Service\Auth
    Type          Name              SourceOfValue Value
    ----          ----              ------------- -----
    System.String Basic                           false
    System.String Kerberos                        true
    System.String Negotiate                       true
    System.String Certificate                     false
    System.String CredSSP                         false
    System.String CbtHardeningLevel               Relaxed
    #-=-=-=-=-=-=-=-=
    # ^ clearly digest doesn't even exist in the list on the service\auth
    
    Need to set to permit Basic Auth too?
    cd .\Auth ;
    Set-Item Basic $True ;
    Check if the user you're connecting with has proper authorizations on the remote machine (triggers GUI after the confirm prompt; use -force to suppress).
    Set-PSSessionConfiguration -ShowSecurityDescriptorUI -Name Microsoft.PowerShell ;
    .PARAMETER  ExchangeServer
    Target Exchange Server[-ExchangeServer server.domain.com]
    .PARAMETER  Scriptblock
    Scriptblock/Command to be executed on target server[-ScriptBlock {Get-Mailbox | ft}]
    .PARAMETER  $Credential
    Credential object to be used for connection[-Credential cred]
    .PARAMETER Whatif
    Parameter to run a Test no-change pass [-Whatif switch]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns objects returned to pipeline
    .EXAMPLE
    .\Invoke-ExchangeCommand.ps1
    .EXAMPLE
    .\Invoke-ExchangeCommand.ps1
    .LINK
    https://github.com/tostka/verb-Ex2010
    .LINK
    https://gallery.technet.microsoft.com/Exchange-Cross-Forest-e25d48eb
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, HelpMessage = "Target Exchange Server[-ExchangeServer server.domain.com]")]
        [string] $ExchangeServer,
        [Parameter(Mandatory = $true, HelpMessage = "Scriptblock/Command to be executed[-ScriptBlock {Get-Mailbox | ft}]")]
        [string] $ScriptBlock,
        [Parameter(Mandatory = $true, HelpMessage = "Credentials [-Credential credobj]")]
        [System.Management.Automation.PSCredential] $Credential
    ) ;
    BEGIN {
        #${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        # Get parameters this function was invoked with
        #$PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        # silently stop any running transcripts
        $stopResults = try { Stop-transcript -ErrorAction stop } catch {} ;
        $WarningPreference = "SilentlyContinue" ;
    } ; # BEGIN-E
    PROCESS {
        $Error.Clear() ;
        #Connect to DC and pass through credential variable
        $pltICPS = @{
            ComputerName   = $ExchangeServer ;
            ArgumentList   = $Credential, $ExchangeServer, $ScriptBlock, $WarningPreference ;
            Credential     = $Credential ;
            Authentication = 'Negotiate'
        } ;
        write-verbose "Invoke-Command  w`n$(($pltICPS|out-string).trim())`n`$ScriptBlock:`n$(($ScriptBlock|out-string).trim())" ;
        #Invoke-Command -ComputerName $ExchangeServer -ArgumentList $Credential,$ExchangeServer,$ScriptBlock,$WarningPreference -Credential $Credential -Authentication Negotiate
        Invoke-Command @pltICPS -ScriptBlock {

                #Specify parameters
                param($Credential,$ExchangeServer,$ScriptBlock,$WarningPreference)

                #Create new PS Session
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ `
                -Authentication Kerberos -Credential $Credential

                #Import PS Session
                Import-PSSession $Session | Out-Null

                #Run commands
                foreach($Script in $ScriptBlock){
                    Invoke-Expression $Script
                }

                #Close all open sessions
                Get-PSSession | Remove-PSSession -Confirm:$false
            }
    } ; # PROC-E
    END {    } ; # END-E
}

#*------^ Invoke-ExchangeCommand.ps1 ^------


#*------v load-EMSLatest.ps1 v------
function load-EMSLatest {
  #  #Checks local machine for registred E20[13|10|07] EMS, and then loads the newest one found
  #Returns the string 2013|2010|2007 for reuse for version-specific code

  <#
  .SYNOPSIS
  load-EMSLatest - Checks local machine for registred E20[13|10|07] EMS, and then loads the newest one found.
  Attempts remote Ex2010 connection if no local EMS installed
  Returns the string 2013|2010|2007 for reuse for version-specific code
    .NOTES
  Author: Todd Kadrie
  Website:	http://tinstoys.blogspot.com
  Twitter:	http://twitter.com/tostka
  REVISIONS   :
  * 6:59 PM 1/15/2020 cleanup
  9:39 AM 2/4/2015 updated to remote to a local hub, updated latest TOR
    .INPUTS
  None. Does not accepted piped input.
    .OUTPUTS
  Returns version number connected to: [2013|2010|2007]
    .EXAMPLE
  .\load-EMSLatest
    .LINK
  #>

  # check registred & loaded ;
  $SnapsReg = Get-PSSnapin -Registered ;
  $SnapsLoad = Get-PSSnapin ;
  $Snapin13 = "Microsoft.Exchange.Management.PowerShell.E2013";
  $Snapin10 = "Microsoft.Exchange.Management.PowerShell.E2010";
  $Snapin7 = "Microsoft.Exchange.Management.PowerShell.Admin";
  # check/load E2013, E2010, or E2007, stop at newest (servers wouldn't be running multi-versions)
  if (($SnapsReg | where { $_.Name -eq $Snapin13 })) {
    if (!($SnapsLoad | where { $_.Name -eq $Snapin13 })) {
      Add-PSSnapin $Snapin13 -ErrorAction SilentlyContinue ; return "2013" ;
    }
    else {
      return "2013" ;
    } # if-E
  }
  elseif (($SnapsReg | where { $_.Name -eq $Snapin10 })) {
    if (!($SnapsLoad | where { $_.Name -eq $Snapin10 })) {
      Add-PSSnapin $Snapin10 -ErrorAction SilentlyContinue ; return "2010" ;
    }
    else {
      return "2010" ;
    } # if-E
  }
  elseif (($SnapsReg | where { $_.Name -eq $Snapin7 })) {
    if (!($SnapsLoad | where { $_.Name -eq $Snapin7 })) {
      Add-PSSnapin $Snapin7 -ErrorAction SilentlyContinue ; return "2007" ;
    }
    else {
      return "2007" ;
    } # if-E
  }
  else {
    Write-Verbose "Unable to locate Exchange tools on localhost, attempting to remote to Exchange 2010 server...";
    #Try implicit remoting-only works for Exchange 2010
    Try {
      # connect to a local hub (leverages ADSI function)
      $Ex2010Server = (Get-ExchangeServerInSite | ? { $_.Roles -match "^(36|38)$" })[0].fqdn
      $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Ex2010Server/PowerShell/ -ErrorAction Stop ;
      Import-PSSession $ExchangeSession -ErrorAction Stop;
    }
    Catch {
      Write-Host -ForegroundColor Red "Unable to import Exchange tools from $Exchange2010Server, is it running Exchange 2010?" ;
      Write-Host -ForegroundColor Magenta "Error:  $($Error[0])" ;
      Exit;
    } # try-E
  }# if-E
}

#*------^ load-EMSLatest.ps1 ^------


#*------v Load-EMSSnap.ps1 v------
function Load-EMSSnap {
  <#
    .SYNOPSIS
    Checks local machine for registred Exchange2010 EMS, and loads the component
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter:	http://twitter.com/tostka

    REVISIONS   :
    # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods/add-Pssnapins
    * 6:59 PM 1/15/2020 cleanup
    vers: 9:39 AM 8/12/2015: retool into generic switched version to support both modules & snappins with same basic code ; building a stock EMS version (vs the fancier load-EMSSnapLatest)
    vers: 10:43 AM 1/14/2015 fixed return & syntax expl to true/false
    vers: 10:20 AM 12/10/2014 moved commentblock into function
    vers: 11:40 AM 11/25/2014 adapted to Lync
    ers: 2:05 PM 7/19/2013 typo fix in 2013 code
    vers: 1:46 PM 7/19/2013
    .INPUTS
    None.
    .OUTPUTS
    Outputs $true if successful. $false if failed.
    .EXAMPLE
    $EMSLoaded = Load-EMSSnap ; Write-Debug "`$EMSLoaded: $EMSLoaded" ;
    Stock free-standing Exchange Mgmt Shell load
    .EXAMPLE
    $EMSLoaded = Load-EMSSnap ; Write-Debug "`$EMSLoaded: $EMSLoaded" ; get-exchangeserver | out-null ;
    Example utilizing a workaround for bug in EMS, where loading ADMS causes Powershell/ISE to crash if ADMS is loaded after EMS, before EMS has executed any commands
    .EXAMPLE
    TRY {
        if(($host.version.major -lt 3) -AND (get-service MSExchangeADTopology -ea SilentlyContinue)){
                write-verbose -verbose:$bshowVerbose  "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Using Local Server EMS10 Snapin" ;
                $sName="Microsoft.Exchange.Management.PowerShell.E2010"; if (!(Get-PSSnapin | where {$_.Name -eq $sName})) {Add-PSSnapin $sName -ea Stop} ;
        } else {
             write-verbose -verbose:$bshowVerbose  "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Initiating REMS connection" ;
            $reqMods="connect-Ex2010;Disconnect-Ex2010;".split(";") ;
            $reqMods | % {if( !(test-path function:$_ ) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing $($_) function. EXITING." } } ;
            Reconnect-Ex2010 ;
        } ;
    } CATCH {
        Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
        Exit #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ;
    } ;
    Example demo'ing check for local psv2 & ADtopo svc to defer
    #>

  # check registred v loaded ;
  # style of plugin we want to test/load
  $PlugStyle = "Snapin"; # for Exch EMS
  #"Module" ; # for Lync/ADMS
  $PlugName = "Microsoft.Exchange.Management.PowerShell.E2010" ;

  switch ($PlugStyle) {
    "Module" {
      # module-style (for LMS or ADMS
      $PlugsReg = Get-Module -ListAvailable;
      $PlugsLoad = Get-Module;
    }
    "Snapin" {
      $PlugsReg = Get-PSSnapin -Registered ;
      $PlugsLoad = Get-PSSnapin ;
    }
  } # switch-E

  TRY {
    if ($PlugsReg | where { $_.Name -eq $PlugName }) {
      if (!($PlugsLoad | where { $_.Name -eq $PlugName })) {
        #
        switch ($PlugStyle) {
          "Module" {
            Import-Module $PlugName -ErrorAction Stop -verbose:$($false); write-output $TRUE ;
          }
          "Snapin" {
            Add-PSSnapin $PlugName -ErrorAction Stop -verbose:$($false); write-output $TRUE
          }
        } # switch-E
      }
      else {
        # already loaded
        write-output $TRUE;
      } # if-E
    }
    else {
      Write-Error { "$(Get-TimeStamp):($env:computername) does not have $PlugName installed!"; };
      #return $FALSE ;
      write-output $FALSE ;
    } # if-E ;
  } CATCH {
        Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
        Exit #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ;
  } ;

}

#*------^ Load-EMSSnap.ps1 ^------


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
    # 1:15 PM 9/6.2.33 updated CBH, pulled in expls from 7PSnMbxG/psb-PSnewMbxG.cbp. Works with current cba auth etc. 
    # 10:30 AM 10/13/2021 pulled [int] from $ticket , to permit non-numeric & multi-tix
    * 11:37 AM 9/16.2.31 string
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
    FileName    : new-MailboxShared.ps1
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite:	URL
    AddedTwitter:	URL
    REVISIONS
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
        # [int] # 10:30 AM 10/13/2021 pulled, to permit non-numeric & multi-tix
        $Ticket,
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
            if($lVers){                 $lVers=($lVers | Sort-Object version)[-1];                 try {                     import-module -name $tModName -RequiredVersion $lVers.Version.tostring() -force -DisableNameChecking -verbose:$($false)                }   catch {                      write-warning "*BROKEN INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;import-module -name $tModDFile -force -DisableNameChecking -verbose:$($false)                } ;
            } elseif (test-path $tModFile) {                 write-warning "*NO* INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;                 try {import-module -name $tModDFile -force -DisableNameChecking -verbose:$($false)}                 catch {                     write-error "*FAILED* TO LOAD MODULE*:$($tModName) VIA $(tModFile) !" ;                     $tModFile = "$($tModName).ps1" ;                     $sLoad = (join-path -path $LocalInclDir -childpath $tModFile) ;                     if (Test-Path $sLoad) {                         write-verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                         . $sLoad ;                         if ($showdebug) { write-verbose "Post $sLoad" };                     } else {                         $sLoad = (join-path -path $backInclDir -childpath $tModFile) ;                         if (Test-Path $sLoad) {                             write-verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                             . $sLoad ;                             if ($showdebug) { write-verbose "Post $sLoad" };                         } else {                             Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;                             exit;                         } ;                     } ;                 } ;             } ;
            if(!(test-path function:$tModCmdlet)){                 write-warning "UNABLE TO VALIDATE PRESENCE OF $tModCmdlet`nfailing through to `$backInclDir .ps1 version" ;                 $sLoad = (join-path -path $backInclDir -childpath "$($tModName).ps1") ;                 if (Test-Path $sLoad) {                     write-verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                     . $sLoad ;                     if ($showdebug) { write-verbose "Post $sLoad" };                     if(!(test-path function:$tModCmdlet)){                         write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO CONFIRM `$tModCmdlet:$($tModCmdlet) FOR $($tModName)" ;                     } else {                         write-verbose  "(confirmed $tModName loaded: $tModCmdlet present)"                     }                 } else {                     Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;                     exit;                 } ;
            } else {                 write-verbose  "(confirmed $tModName loaded: $tModCmdlet present)"             } ;
        } ;  # loop-E
        #*------^ END MOD LOADS ^------

        <# rem, shifting preprocessor to module, loses $parentpath function, (resolves to module file in allusers context)
        if($ParentPath){
            $rgxProfilePaths='(\\Documents\\WindowsPowerShell\\scripts|\\Program\sFiles\\windowspowershell\\scripts)' ;
            if($ParentPath -match $rgxProfilePaths){
                if(test-path -Path 'd:\scripts\'){
                    $ParentPath = "$(join-path -path 'd:\scripts\' -ChildPath (split-path $ParentPath -leaf))" ;
                }else{
                    $ParentPath = "$(join-path -path 'c:\scripts\' -ChildPath (split-path $ParentPath -leaf))" ;
                } ; 
            } ;
            $logspec = start-Log -Path ($ParentPath) -showdebug:$($showdebug) -whatif:$($whatif) -tag $DisplayName;
            if($logspec){
                $logging=$logspec.logging ;
                $logfile=$logspec.logfile ;
                $transcript=$logspec.transcript ;
            } else {$smsg = "Unable to configure logging!" ; write-warning "$((get-date).ToString('HH:mm:ss')):$($sMsg)" ; Exit ;} ;
        } else {$smsg = "No functional `$ParentPath found!" ; write-warning "$((get-date).ToString('HH:mm:ss')):$($sMsg)" ;  Exit ;} ;
        #>
        # detect profile installs (installed mod or script), and redir to stock location
        $dPref = 'd','c' ; foreach($budrv in $dpref){ if(test-path -path "$($budrv):\scripts" -ea 0 ){ break ;  } ;  } ;
        [regex]$rgxScriptsModsAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)" ;
        [regex]$rgxScriptsModsCurrUserScope="^$([regex]::escape([environment]::getfolderpath('Mydocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)" ;
        # -Tag "($TenOrg)-LASTPASS" 
        $pltSLog = [ordered]@{ NoTimeStamp=$false ; Tag=$lTag  ; showdebug=$($showdebug) ;whatif=$($whatif) ;} ;
        if($PSCommandPath){
            if(($PSCommandPath -match $rgxScriptsModsAllUsersScope) -OR ($PSCommandPath -match $rgxScriptsModsCurrUserScope) ){
                # AllUsers or CU installed script, divert into [$budrv]:\scripts (don't write logs into allusers context folder)
                if($PSCommandPath -match '\.ps(d|m)1$'){
                    # module function: use the ${CmdletName} for childpath
                    $pltSLog.Path= (join-path -Path "$($budrv):\scripts" -ChildPath "$(${CmdletName}).ps1" )  ;
                } else { 
                    $pltSLog.Path=(join-path -Path "$($budrv):\scripts" -ChildPath (split-path $PSCommandPath -leaf)) ;
                } ; 
            }else {
                $pltSLog.Path=$PSCommandPath ;
            } ;
        } else {
            if( ($MyInvocation.MyCommand.Definition -match $rgxScriptsModsAllUsersScope) -OR ($MyInvocation.MyCommand.Definition -match $rgxScriptsModsCurrUserScope) ){
                $pltSLog.Path=(join-path -Path "$($budrv):\scripts" -ChildPath (split-path $PSCommandPath -leaf)) ;
            } else {
                $pltSLog.Path=$MyInvocation.MyCommand.Definition ;
            } ;
        } ;
        $smsg = "start-Log w`n$(($pltSLog|out-string).trim())" ;
        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        $logspec = start-Log @pltSLog ;
        if($logspec){
            $logging=$logspec.logging ;
            $logfile=$logspec.logfile ;
            $transcript=$logspec.transcript ;

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

            if(Test-TranscriptionSupported){
                $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                start-transcript -Path $transcript ;
            } ;
        } else {throw "Unable to configure logging!" } ;

        

        $xxx="====VERB====";
        $xxx=$xxx.replace("VERB","NewMbx") ;
        $BARS=("="*10);

        $reqMods+="Add-PSTitleBar;Remove-PSTitleBar".split(";") ;
        $reqMods+="Test-TranscriptionSupported;Test-Transcribing;Stop-TranscriptLog;Start-IseTranscript;Start-TranscriptLog;get-ArchivePath;Archive-Log;Start-TranscriptLog".split(";") ;
        $reqMods=$reqMods| Select-Object -Unique ;

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
        if(get-pssession |Where-Object{($_.configurationname -eq 'Microsoft.Exchange') -AND ($_.ComputerName -match $rgxEx10HostName) -AND ($_.IdleTimeout -ne -1)} ){
            write-verbose  "$((get-date).ToString('HH:mm:ss')):LOCAL EMS detected" ;
            $Global:E10IsDehydrated=$false ;
        # REMS detect dleTimeout -eq -1
        } elseif(get-pssession |Where-Object{$_.configurationname -eq 'Microsoft.Exchange' -AND $_.ComputerName -match $rgxEx10HostName -AND ($_.IdleTimeout -eq -1)} ){
            write-verbose  "$((get-date).ToString('HH:mm:ss')):REMOTE EMS detected" ;
            $reqMods+="get-GCFast;Get-ExchangeServerInSite;connect-Ex2010;Reconnect-Ex2010;Disconnect-Ex2010;Disconnect-PssBroken".split(";") ;
            if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
            reconnect-ex2010 ;
            $Global:E10IsDehydrated=$true ;
        } else {
            write-verbose  "$((get-date).ToString('HH:mm:ss')):No existing Ex2010 Connection detected" ;
            # Server snapin defer
            if(($host.version.major -lt 3) -AND (get-service MSExchangeADTopology -ea SilentlyContinue)){
                write-verbose "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Loading Local Server EMS10 Snapin" ;
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
                        if ( $InputSplat.BaseUser=get-mailbox -OrganizationalUnit $($MbxSplat.OrganizationalUnit) -resultsize 50 | Where-Object{($_.distinguishedname -notlike '*demo*') -AND (!$_.CustomAttribute5)} | get-random ) {

                        } elseif ( $InputSplat.BaseUser=get-remotemailbox -OnPremisesOrganizationalUnit  $($MbxSplat.OrganizationalUnit) -resultsize 50 | Where-Object{($_.distinguishedname -notlike '*demo*') -AND (!$_.CustomAttribute5)} | get-random ) {

                        } else {
                            $smsg= "UNABLE TO FIND A BASEUSER - USE -BASEUSER TO SPECIFY A SUITABLE ACCT *SOMEWHERE*" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                            Cleanup ; Exit ;
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
                $smsg= "`n **** NOTE TRUNCATING LASTNAME, -GT 64 CHARS!  ****`N" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
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
                        $smsg = "New-Mailbox  w`n$(($MbxSplat|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        # do loop until up to 4 retries...
                        Do {
                            Try {
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
                        if(!$oNMbx){
                            write-verbose "(using returned New-Mailbox output object...)" ; 
                            do {Write-Host "." -NoNewLine;Start-Sleep -s 1} until ($oMbx = (get-mailbox -identity $Mbxsplat.samaccountname -domaincontroller $Mbxsplat.DomainController -ea silentlycontinue)) ;
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

                    #region UPNFromEmail ; # ------
                    # pull the SMTP: addr and use it to force the UPN
                    if($Whatif){
                        $smsg= "Whatif skipping UPN Update...";if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                    } else {
                        # dynm pull the forestdom from the forest, match @DOMAIN.COM
                        $forestdom=get-adforest -ea stop | Select-Object -expand upnsuffixes |Where-Object{$_ -match $rgxTTCDomainsLegacy}
                        if($forestdom -is [string]){
                            # pull primary SMTP:, verify -is [string]/non-array
                            Do {
                                $smsg= "Waiting for ADUser to return email addresses" ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ;
                                write-host "." -NoNewLine;Start-Sleep -m (1000 * 5)
                                $dirname=(Get-ADUser -identity $oMbx.samaccountname -server $InputSplat.domaincontroller -Properties proxyAddresses -ea 0 | Select-Object -Expand proxyAddresses | Where-Object {$_ -clike "SMTP:*"});
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
                        $ADu = get-ADuser -Identity $($InputSplat.samaccountname) -properties * -server $($InputSplat.domaincontroller)| Select-Object *;
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


#*------v preview-EAPUpdate.ps1 v------
function preview-EAPUpdate {
        <#
        .SYNOPSIS
        preview-EAPUpdate.ps1 - Code to approximate EmailAddressTemplate-generated email addresses
        .NOTES
        Version     : 0.0.
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2021-08-23
        FileName    : 
        License     : MIT License
        Copyright   : (c) 2021 Todd Kadrie
        Github      : https://github.com/tostka/verb-XXX
        Tags        : Powershell
        AddedCredit : REFERENCE
        AddedWebsite: URL
        AddedTwitter: URL
        REVISIONS
        * 2:16 PM 6/24/2024: rem'd out #Requires -RunasAdministrator; sec chgs in last x mos wrecked RAA detection
        * 3:45 PM 8/23/2021 added extended examples, made a function (adding to 
        verb-ex2010); added drop of illegal chars (shows up distinctively as spaces in 
        dname :P); fixed bug in regex ps replace;  
        .DESCRIPTION
        preview-EAPUpdate.ps1 - Code to approximate EmailAddressTemplate-generated email addresses
        Note: This is a quick & dirty *approximation* of the generated email address. 
        Doesn't support multiple %rxy replaces on mult format codes. Just does the first one; plus any non-replace %d|s|i|m|g's. 

        Don't *rely* on this. It's just intended to quickly confirm assigned primarysmtpaddress roughly matches the intended EAP template.
        If it doesn't, put eyes on it and *confirm*, don't use this to drive any revision of the email address!

        Latest ref specs:[Email address policies in Exchange Server | Microsoft Docs - docs.microsoft.com/](https://docs.microsoft.com/en-us/exchange/email-addresses-and-address-books/email-address-policies/email-address-policies?view=exchserver-2019)

        Address types: SMTP| GWISE| NOTES| X400
        Address format variables:
        |Variable |Value|
        |---|---|
        |%d |Display name|
        |%g |Given name (first name)|
        |%i |Middle initial|
        |%m |Exchange alias|
        |%rxy |Replace all occurrences of x with y|
        |%rxx |Remove all occurrences of x|
        |%s |Surname (last name)|
        |%ng |The first n letters of the first name. For example, %2g uses the first two letters of the first name.|
        |%ns |The first n letters of the last name. For example, %2s uses the first two letters of the last name.|
        Use of email-legal ASCII chars:
        |Example |Exchange Management Shell equivalent|
        |---|---|
        |<alias>@contoso.com |%m@contoso.com|
        |elizabeth.brunner@contoso.com |%g.%s@contoso.com|
        |ebrunner@contoso.com |%1g%s@contoso.com|
        |elizabethb@contoso.com |%g%1s@contoso.com|
        |brunner.elizabeth@contoso.com |%s.%g@contoso.com|
        |belizabeth@contoso.com |%1s%g@contoso.com|
        |brunnere@contoso.com |%s%1g@contoso.com|

        %RXY REPLACEMENT EXAMPLES. 
        |Source properties||
        |---|---|
        |user logon name|"jwilson"|
        |Display name|James C. Wilson|
        |Surname|Wilson|
        |Given name|James|
        note: In %rXY, if X = Y - same char TWICE - the character will be DELETED rather than REPLACED.
        |Replacement String|SMTP Address Generated|Comment|
        |---|---|---|
        |%d@domain.com|JamesCWilson@domain.com|"Displayname@domain"|
        |%g.%s@microsoft.com|James.Wilson@microsoft.com|"givenname.surname@domain"|
        |@microsoft.com|JamesW@microsoft.com|"userLogon@domain" (default)|
        |%1g%s@microsoft.com|JWilson@microsoft.com|"[1stcharGivenName][surname]@domain"|
        |%1g%3s@microsoft.com|JWil@microsoft.com|"[1stcharGivenName][3charsSurname]@domain"|
        |@domain.com|<email-alias>@domain.com (this is the one item always a part of the Default policy)|
        |%r._%d@microsoft.com|JamesC_Wilson@microsoft.com|"[replace periods in displayname with underscore]@domain"|
        |%r..%d@microsoft.com|JamesC.Wilson@microsoft.com|"[DELETE periods in displayname]@domain",(avoids double period if name trails with a period)|
        |%g.%r''%s@domain.com|James.Wilson@domain|"[givenname].[surname,delete all APOSTROPHES]@domain"|
        |%r''%g.%r''%s@domain.com|James.Wilson@domain|"[givenname,delete all APOSTROPHES].[surname,delete all APOSTROPHES]@domain"|

        .PARAMETER  EmailAddressPolicy
        EmailAddressPolicy object to be modeled for primarysmtpaddress update
        .PARAMETER  Recipient
        Recipient object to be modeled
        .PARAMETER useEXOv2
        Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
        .INPUTS
        None. Does not accepted piped input.(.NET types, can add description)
        .OUTPUTS
        System.String
        [| get-member the output to see what .NET obj TypeName is returned, to use here]
        .EXAMPLE
        PS> preview-EAPUpdate  -eap $eaps[16] -Recipient $trcp -verbose ;
        Preview specified recipient using the specified EAP (17th in the set in the $eaps variable).
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate '%5s%3g@contoso.com' -Recipient $trcp -verbose ;
        Preview target recipient using a template that takes surname[0-4]givename[0-3]@contoso.com
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate '%5s%1g%1i@contoso.com' -Recipient $trcp -verbose ;
        Preview target recipient using a template that takes surname[0-4]givename[0]mi[0]@contoso.com
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate '%1g%s@contoso.com' -Recipient $trcp -verbose ;
        Preview target recipient using %1g%s@contoso.com|JWilson@contoso.com|"[1stcharGivenName][surname]@domain"|
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate '%r._%d@contoso.com' -Recipient $trcp -verbose ;
        Preview target recipient using %r._%d@contoso.com|JamesC_Wilson@contoso.com|"[replace periods in displayname with underscore]@domain"|
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate "%g.%r''%s@domain.co" -Recipient $trcp -verbose ;
        Preview target recipient using %g.%r''%s@domain.com|James.Wilson@domain|"[givenname].[surname,delete all APOSTROPHES]@domain"|
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate " %r''%g.%r''%s@domain.com" -Recipient $trcp -verbose ;
        Preview target recipient using %r''%g.%r''%s@domain.com|James.Wilson@domain|"[givenname,delete all APOSTROPHES].[surname,delete all APOSTROPHES]@domain"
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate ' %d@contoso.com' -Recipient $trcp -verbose ;
        Preview target recipient using %d@domain.com|JamesCWilson@domain.com|"Displayname@domain"
        .EXAMPLE
        PS> $genEml = preview-EAPUpdate  -addresstemplate ' %d@contoso.com' -Recipient $trcp -verbose ;
            if(($geneml -ne $trcp.primarysmtpaddress)){
                write-warning "Specified recip's PrimarySmtpAddress does *not* appear to match specified template!`nmanualy *review* the template specs`nand validate that the desired scheme is being applied!"
            }else {
                "PrimarysmtpAddr $($trcp.primarysmtpaddress) roughly conforms to specified template primary addr..."
            } ;
        Example testing output against $trcp primarySmtpAddress.
        .LINK
        https://github.com/tostka/verb-ex2010
        .LINK
        https://docs.microsoft.com/en-us/exchange/email-addresses-and-address-books/email-address-policies/email-address-policies?view=exchserver-2019
        #>
        ###Requires -Version 5
        ###Requires -Modules verb-Ex2010 - disabled, moving into the module
        ##Requires -RunasAdministrator
        # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
        ## [OutputType('bool')] # optional specified output type
        [CmdletBinding(DefaultParameterSetName='EAP')]
        PARAM(
            [Parameter(ParameterSetName='EAP',Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Exchange EmailAddressPolicy object[-eap `$eaps[16]]")]
            [ValidateNotNullOrEmpty()]
            #[Alias('ALIAS1', 'ALIAS2')]
            $EAP,
            [Parameter(ParameterSetName='Template',Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Specify raw AddressTemplate string, for output modeling[-AddressTemplate '%3g%5s@microsoft.com']")]
            [ValidateNotNullOrEmpty()]
            #[Alias('ALIAS1', 'ALIAS2')]
            $AddressTemplate,
            [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of recipient descriptors: displayname, emailaddress, UPN, samaccountname[-recip some.user@domain.com]")]
            #[ValidateNotNullOrEmpty()]
            #[Alias('ALIAS1', 'ALIAS2')]
            $Recipient,
            [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
            [switch] $useEXOv2
        ) ;
        BEGIN { 
            # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
            ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
            $rgxEmailDirLegalChars = "[0-9a-zA-Z-._+&']" ; 
            reconnect-ex2010 -verbose:$false ; 
    
        } ;  # BEGIN-E
        PROCESS {
            $Error.Clear() ; 
            if($EAP){
                $ptmpl= $eap.EnabledPrimarySMTPAddressTemplate ;
            } elseif($AddressTemplate){
                $ptmpl= $AddressTemplate ;
            } ; 
            $error.clear() ;
            TRY {
                if($Recipient.alias){
                    $Recipient = get-recipient $Recipient.alias ;
                } else { throw "-recipient invalid, has no Alias property!`nMust be a valid Exchange Recipient object" } ; 
                $usr = get-user -id $Recipient.alias -ErrorAction STOP ; 
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
            $genAddr = $null ; 
            # [0-9a-zA-Z-._+&]{1,64} # alias rgx legal
            # [0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+
            # dirname
            # ^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+ 
            # @domain: @([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,6}$

            #if($ptmpl -match '^@.*$'){$genAddr = $ptmpl = $ptmpl.replace('@',"$($Recipient.alias)@")} ;
            if($ptmpl -match '^@.*$'){
                write-verbose "(matched Alias@domain)" ;
                $string = ($Recipient.alias.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl.replace('@',"$($string)@")
            } ;
            # do replace first, as the %d etc are simpler matches and don't handle the leading %r properly.
            if($ptmpl -match "%r([a-zA-Z_0-9-._+&])([a-zA-Z_0-9-._+&])%(d|g|i|m|s)@.*$"){
                $x = $matches[1] ; 
                $y = $matches[2] ; 
                $vari = $matches[3] ; 
                write-verbose "Parsed:`nx:$($x)`ny:$($y)`nvari:$($vari)" ;
                switch ($vari){
                    'd' {
                        write-verbose "(matched replace $($x) w $($y) on displayname)" ;
                        #$ptmpl = $ptmpl.replace("%d",$usr.displayname.replace($x,$y)) ;
                        $string = ($usr.displayname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                        #$genAddr = $ptmpl = $ptmpl = $ptmpl.replace("%d",$string.replace($x,$y))
                        # subout %rxy first, then the trailing %d w name
                        $genAddr = $ptmpl = $ptmpl = $ptmpl -replace "%r([a-zA-Z_0-9-._+&])([a-zA-Z_0-9-._+&])",'' -replace "%d",$string.replace($x,$y)
                    }
                    'g'  {
                        write-verbose "(matched replace $($x) w $($y) on givenname)" ;
                        #$ptmpl = $ptmpl.replace("%g",$usr.firstname.replace($x,$y)) ;
                        $string = ($usr.firstname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                        #$genAddr = $ptmpl = $ptmpl = $ptmpl.replace("%g",$string.replace($x,$y))
                        $genAddr = $ptmpl = $ptmpl = $ptmpl -replace "%r([a-zA-Z_0-9-._+&])([a-zA-Z_0-9-._+&])",'' -replace "%g",$string.replace($x,$y)
                    }
                    'i' {
                        write-verbose "(matched replace $($x) w $($y) on initials)" ;
                        #$ptmpl = $ptmpl.replace("%i",$usr.Initials.replace($x,$y)) ;
                        $string = ($usr.Initials.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                        #$genAddr = $ptmpl = $ptmpl = $ptmpl.replace("%i",$string.replace($x,$y))
                        $genAddr = $ptmpl = $ptmpl = $ptmpl -replace "%r([a-zA-Z_0-9-._+&])([a-zA-Z_0-9-._+&])",'' -replace "%i",$string.replace($x,$y)
                    }
                    'm' {
                        write-verbose "(matched replace $($x) w $($y) on alias)" ;
                        #$ptmpl = $ptmpl.replace("%m",$Recipient.alias.replace($x,$y)) ;
                        $string = ($Recipient.alias.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                        #$genAddr = $ptmpl = $ptmpl = $ptmpl.replace("%m",$string.replace($x,$y))
                        $genAddr = $ptmpl = $ptmpl = $ptmpl -replace "%r([a-zA-Z_0-9-._+&])([a-zA-Z_0-9-._+&])",'' -replace "%m",$string.replace($x,$y)
                    }
                    's' {
                        write-verbose "(matched replace $($x) w $($y) on surname)" ;
                        #$ptmpl = $ptmpl.replace("%s",$usr.lastname.replace($x,$y)) ;
                        $string = ($usr.lastname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                        #$genAddr = $ptmpl = $ptmpl = $ptmpl.replace("%s",$string.replace($x,$y))
                        $genAddr = $ptmpl = $ptmpl = $ptmpl -replace "%r([a-zA-Z_0-9-._+&])([a-zA-Z_0-9-._+&])",'' -replace "%s",$string.replace($x,$y)
                    }
                    default {
                        throw "unrecognized template: replace (%r) character with no targeted variable (%(d|g|i|m|s))" ;
                    }
                } ;
            } ; 
            if($ptmpl.contains('%g')){
                write-verbose "(matched %g:displayname)" ;
                #$genAddr = $ptmpl = $ptmpl.replace('%g',$usr.firstname)
                $string = ($usr.firstname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl.replace('%g',$string)
            } ;
            if($ptmpl.contains('%s')){
                write-verbose "(matched %s:surname)" ;
                #$genAddr = $ptmpl = $ptmpl.replace('%s',$usr.lastname)} ;
                $string = ($usr.lastname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl.replace('%s',$string)
            };
            if($ptmpl.contains('%d')){
                write-verbose "(matched %d:displayname)" ;
                #$genAddr = $ptmpl = $ptmpl.replace('%d',$usr.displayname)
                $string = ($usr.displayname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl.replace('%d',$string)
            } ;
            if($ptmpl.contains('%i')){
                write-verbose "(matched %i:initials)" ;
                #$genAddr = $ptmpl = $ptmpl.replace('%i',$usr.Initials)} ;
                $string = ($usr.Initials.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl.replace('%i',$string)
            } ;
            if($ptmpl.contains('%m')){
                write-verbose "(matched %m:alias)" ;
                #$genAddr = $ptmpl = $ptmpl.replace('%m',$Recipient.alias)} ;
                $string = ($Recipient.alias.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl.replace('%m',$string)
            } ; 
            if($ptmpl -match '(%(\d)g)'){
                $ltrs = $matches[2] ; 
                write-verbose "(matched %g:displayname, first $($ltrs) chars)" ;
                #$genAddr = $ptmpl = $ptmpl -replace "(%\dg)","$1,$($usr.firstname.substring(0,$ltrs))" ;
                #$genAddr = $ptmpl = $ptmpl -replace "(%\dg)","$($usr.firstname.substring(0,$ltrs))"
                $string = ($usr.firstname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl -replace "(%\dg)","$($string.substring(0,$ltrs))"
            } ; 
            if($ptmpl -match "(%(\d)s)"){
                $ltrs = $matches[2] ; 
                write-verbose "(matched %s:surname, first $($ltrs) chars)" ;
                #$genAddr = $ptmpl = $ptmpl -replace "(%\ds)","$1,$($usr.lastname.substring(0,$ltrs))"
                #$genAddr = $ptmpl =$ptmpl -replace "(%\ds)","$($usr.lastname.substring(0,$ltrs))" ;
                $string = ($usr.lastname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl -replace "(%\ds)","$($string.substring(0,$ltrs))"
            } ; 
            if($ptmpl -match "(%(\d)i)"){
                $ltrs = $matches[2] ; 
                write-verbose "(matched %i:initials, first $($ltrs) chars)" ;
                #$genAddr = $ptmpl = $ptmpl -replace "(%\ds)","$1,$($usr.lastname.substring(0,$ltrs))"
                #$genAddr = $ptmpl =$ptmpl -replace "(%\di)","$($usr.initials.substring(0,$ltrs))" ;
                $string = ($usr.initials.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl -replace "(%\di)","$($string.substring(0,$ltrs))"
            } ; 

        } ;  # PROC-E
        END {
            if($genAddr){
                write-verbose "returning generated address:$($genAddr)" ; 
                $genAddr| write-output 
            } else {
                write-warning "Unable to generate a PrimarySmtpAddress model for user" ; 
                $false | write-output l
            };
        } ;  
    }

#*------^ preview-EAPUpdate.ps1 ^------


#*------v Reconnect-Ex2010.ps1 v------
Function Reconnect-Ex2010 {
  <#
    .SYNOPSIS
    Reconnect-Ex2010 - Reconnect Remote ExchOnPrem Mgmt Shell connection (validated functional Exch2010 - Exch2016)
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    AddedCredit : Inspired by concept code by ExactMike Perficient, Global Knowl... (Partner)
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Version     : 1.1.0
    CreatedDate : 2020-02-24
    FileName    : Reonnect-Ex2010()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell
    REVISIONS   :
    * 3:11 PM 7/15/2024 needed to change CHKPREREQ to check for presence of prop, not that it had a value (which fails as $false); hadn't cleared $MetaProps = ...,'DOESNTEXIST' ; confirmed cxo working non-based
    * 10:47 AM 7/11/2024 cleared debugging NoSuch etc meta tests
    * 1:34 PM 6/21/2024 ren $Global:E10Sess -> $Global:EXOPSess ;updated $rgxRemsPSSName = "^(Session\d|Exchange\d{4}|Exchange\d{2}((\.\d+)*))$" ;
    * 11:02 AM 10/25/2021 dbl/triple-connecting, fliped $E10Sess -> $global:E10Sess (must not be detecting the preexisting session), added post test of session to E10Sess values, to suppres redund dxo/rxo.
    * 1:17 PM 8/17/2021 added -silent param
    * 4:31 PM 5/18/2l lost $global:credOpTORSID, sub in $global:credTORSID
    * 10:52 AM 4/2/2021 updated cbh
    * 1:56 PM 3/31/2021 rewrote to dyn detect pss, rather than reading out of date vari
    * 10:14 AM 3/23/2021 fix default $Cred spec, pointed at an OP cred
    * 8:29 AM 11/17/2020 added missing $Credential param
    * 9:33 AM 5/28/2020 actually added the alias:rx10
    * 12:20 PM 5/27/2020 updated cbh, moved alias: rx10 win func
    * 6:59 PM 1/15/2020 cleanup
    * 8:09 AM 11/1/2017 updated example to pretest for reqMods
    * 1:26 PM 12/9/2016 split no-session and reopen code, to suppress notfound errors, add pshelpported to local EMSRemote
    * 2/10/14 posted version
    .DESCRIPTION
    Reconnect-Ex2010 - Reconnect Remote Exch2010 Mgmt Shell connection
    .PARAMETER  Credential
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    $reqMods="connect-Ex2010;Disconnect-Ex2010;".split(";") ;
    $reqMods | % {if( !(test-path function:$_ ) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing $($_) function. EXITING." } } ;
    Reconnect-Ex2010 ;
    .LINK
    #>
    [CmdletBinding()]
    [Alias('rx10','rxOP','reconnect-ExOP')]
    Param(
        [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")][System.Management.Automation.PSCredential]
        $Credential = $global:credTORSID,
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
        [switch] $silent
    )
    BEGIN{
        # checking stat on canned copy of hist sess, says nothing about current, possibly timed out, check them manually
        $rgxRemsPSSName = "^(Session\d|Exchange\d{4}|Exchange\d{2}((\.\d+)*))$" ;

        #region CHKPREREQ ; #*------v CHKPREREQ v------
        # critical dependancy Meta variables
        $MetaNames = ,'TOR','CMW','TOL' #,'NOSUCH' ;
        # critical dependancy Meta variable properties
        $MetaProps = 'Ex10Server','Ex10WebPoolVariant','ExRevision','ExViewForest','ExOPAccessFromToro','legacyDomain' #,'DOESNTEXIST' ;
        # critical dependancy parameters
        $gvNames = 'Credential'
        $isBased = $true ; $gvMiss = @() ; $ppMiss = @() ;
        foreach($met in $metanames){
            write-verbose "chk:`$$($met)Meta" ;
            if(-not (gv -name "$($met)Meta" -ea 0)){$isBased = $false; $gvMiss += "$($met)Meta" } ;
            if($MetaProps){
                foreach($mp in $MetaProps){
                    write-verbose "chk:`$$($met)Meta.$($mp)" ;
                    #if(-not (gv -name "$($met)Meta" -ea 0).value[$mp]){ # testing has a value, not is present as a spec!
                    if(-not (gv -name "$($met)Meta" -ea 0).value.keys -contains $mp){
                        $isBased = $false; $ppMiss += "$($met)Meta.$($mp)"
                    } ;
                } ;
            } ;
        } ;
        if($gvNames){
            foreach($gvN in $gvNames){
                write-verbose "chk:`$$($gvN)" ;
                if(-not (gv -name "$($gvN)" -ea 0)){$isBased = $false; $gvMiss += "$($gvN)" } ;
            } ;
        } ;
        if($gvmiss){write-warning "Missing Dependant Meta variables:`n$(($gvMiss |%{"`$$($_)" }) -join ',')" } ;
        if($ppMiss){write-warning "Missing Dependant Meta vari properties:`n$(($ppMiss |%{"`$$($_)" }) -join ',')" } ;
        if(-not $isBased){ write-warning  "missing critical dependancy profile config!" } ;
        #endregion CHKPREREQ ; #*------^ END CHKPREREQ ^------

        if($isBased){
            # back the TenOrg out of the Credential        
            $TenOrg = get-TenantTag -Credential $Credential ;
        } ; 
    }  # BEG-E
    PROCESS{
        if($isBased){
            $Rems2Good = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') } ;
            $Rems2WrongOrg = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND (
                $_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND (
                ( -not($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) ) -AND (
                -not($_.ComputerName -match $rgxExoPsHostName)) ) -AND ($_.Availability -eq 'Available')
            } ;
        }else {
            $Rems2Good = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.Availability -eq 'Available') } ;
        } ; 

        $Rems2Broken = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND (
            $_.Name -match $rgxRemsPSSName) -AND ($_.State -like "*Broken*") } ;
        $Rems2Closed = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND (
            $_.Name -match $rgxRemsPSSName) -AND ($_.State -like "*Closed*") } ;
        if ($Rems2Broken.count -gt 0){ for ($index = 0 ;$index -lt $Rems2Broken.count ;$index++){Remove-PSSession -session $Rems2Broken[$index]}  };
        if ($Rems2Closed.count -gt 0){for ($index = 0 ;$index -lt $Rems2Closed.count ; $index++){Remove-PSSession -session $Rems2Closed[$index] } } ;
        if ($Rems2WrongOrg.count -gt 0){for ($index = 0 ;$index -lt $Rems2WrongOrg.count ; $index++){Remove-PSSession -session $Rems2WrongOrg[$index] } } ;
        #if( -not ($Global:EXOPSess ) -AND -not ($Rems2Good)){
        if(-not $Rems2Good){
            if (-not $Credential) {
                Connect-Ex2010 # sets $Global:EXOPSess on connect
            } else {
                Connect-Ex2010 -Credential:$($Credential) ; # sets $Global:EXOPSess on connect
            } ;
            if($Global:EXOPSess -AND ($tSess = get-pssession -id $Global:EXOPSess.id -ea 0 |?{$_.computername -eq $Global:EXOPSess.computername -ANd $_.name -eq $Global:EXOPSess.name})){
                # matches historical session
                if( $tSess | where-object { ($_.State -eq "Opened") -AND ($_.Availability -eq 'Available') } ){
                    $bExistingREms= $true ;
                } else {
                    $bExistingREms= $false ;
                } ;
            }elseif($tSess = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) }){ 
                if( $tSess | where-object { ($_.State -eq "Opened") -AND ($_.Availability -eq 'Available') } ){
                    $Global:EXOPSess = $tSess ;
                    $bExistingREms= $true ;
                } else {
                    $bExistingREms= $false ;
                    $Global:EXOPSess = $null ; 
                } ;
            } ; 
        }elseif($tSess = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) }){
            # matches generic session
            if( $tSess | where-object { ($_.State -eq "Opened") -AND ($_.Availability -eq 'Available') } ){
                if(-not $Global:EXOPSess){$Global:E10Sess = $tSess } ; 
                $bExistingREms= $true ;
            } else {
                $bExistingREms= $false ;
            } ;
        } else {
            # doesn't match histo
            $bExistingREms= $false ;
        } ;
        $propsPss =  'Id','Name','ComputerName','ComputerType','State','ConfigurationName','Availability' ;
    
        if($bExistingREms){
            if($silent){} else { 
                $smsg = "existing connection Open/Available:`n$(($tSess| ft -auto $propsPss |out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
        } else {
            $smsg = "(resetting any existing EX10 connection and re-establishing)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Disconnect-Ex2010 ; Start-Sleep -S 3;
            if (-not $Credential) {
                Connect-Ex2010 ;
            } else {
                Connect-Ex2010 -Credential:$($Credential) ;
            } ;
        } ;
    }  # PROC-E
}

#*------^ Reconnect-Ex2010.ps1 ^------


#*------v Reconnect-Ex2010XO.ps1 v------
Function Reconnect-Ex2010XO {
   <#
    .SYNOPSIS
    Reconnect-Ex2010XO - Reconnect Remote Exch2010 Mgmt Shell connection Cross-Org (XO)
    .NOTES
    Updated By: : Todd Kadrie
    Website:	http://toddomation.com
    Twitter:	@tostka http://twitter.com/tostka
    Based on original function Author: ExactMike Perficient, Global Knowl... (Partner)
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    REVISIONS   :
    # 3:18 PM 5/18/2021 somehow lost $credOpTORSID, so flipped lost default $credOPTor -> $credTORSID
    * 1:57 PM 3/31/2021 wrapped long lines for vis
    * 8:30 AM 10/22/2020 ren'd $TentantTag -> $TenOrg, swapped looping meta resolve with 1-liner approach ; added AcceptedDom caching to the middle status test (suppress one more get-exoaccepteddomain call if possible), replaced all $Meta.value with the $TenOrg version
    * 1:19 PM 10/15/2020 converted connect-exo to Ex2010, adding onprem validation
    .DESCRIPTION
    Reconnect-Ex2010XO - Reconnect Remote Exch2010 Mgmt Shell connection Cross-Org (XO)
    .PARAMETER  Credential
    Credential to use for this connection [-credential 'account@domain.com']
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Reconnect-Ex2010XO;
    Reconnect EXO connection
    .EXAMPLE
    $batchsize = 10 ;
    $RecordCount=$mr.count #this is the array of whatever you are processing ;
    $b=0 #this is the initialization of a variable used in the do until loop below ;
    $mrs = @() ;
    do {
        Write-Progress -Activity "Getting move request statistics for all $wave move requests." -Status "Processing Records $b through $($b+$batchsize) of $RecordCount." -PercentComplete ($b/$RecordCount*100) ;
        $mrs += $mr | Select-Object -skip $b -first $batchsize | foreach-object {Reconnect-Ex2010XO; $_ | Get-OLMoveRequestStatistics} ;
        $b=$b+$batchsize ;
        } ;
    until ($b -gt $RecordCount) ;

    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    #>
    [CmdletBinding()]
    [Alias('rx10xo')]
    <#
    Param(
      [Parameter(HelpMessage="Use Proxy-Aware SessionOption settings [-ProxyEnabled]")]
      [boolean]$ProxyEnabled = $False,
      [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")]
      [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
      [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
      [switch] $showDebug
    ) ;
    #>
     Param(
        [Parameter(HelpMessage = 'Credential object')][System.Management.Automation.PSCredential]$Credential = $credTORSID,
        [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
        [switch] $showDebug
    )  ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ;
        #if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        # $rgxEx10HostName : ^(lyn|bcc|adl|spb)ms6[4,5][0,1].global.ad.toro.com$
        # we'd need to define all possible hostnames to cover potential span. Should probably build dynamically from $XXXMeta vari
        # can build from $TorMeta.OP_ExADRoot:global.ad.toro.com
        <# on curly, from Ps into EMS:
        get-pssession | fl computername,computertype,state,configurationname,availability,name
        ComputerName      : curlyhoward.cmw.internal
        ComputerType      : RemoteMachine
        State             : Opened
        ConfigurationName : Microsoft.Exchange
        Availability      : Available
        Name              : Session1

        ComputerName      : lynms650.global.ad.toro.com
        ComputerType      : RemoteMachine
        State             : Broken
        ConfigurationName : Microsoft.Exchange
        Availability      : None
        Name              : Exchange2010

        "^\w*\.$($CMWMeta.OP_ExADRoot)$"
        => ^\w*\.cmw.internal$
        #>

        $sTitleBarTag = "EMS" ;
        $CommandPrefix = $null ;

        $TenOrg=get-TenantTag -Credential $Credential ;
        if($TenOrg -ne 'TOR'){
            # explicitly leave this tenant (default) untagged
            $sTitleBarTag += $TenOrg ;
        } ;
        <#
        $credDom = ($Credential.username.split("\"))[0] ;
        $Metas=(get-variable *meta|Where-Object{$_.name -match '^\w{3}Meta$'}) ;
        foreach ($Meta in $Metas){
            if( ($credDom -eq $Meta.value.legacyDomain) -OR ($credDom -eq $Meta.value.o365_TenantDomain) -OR ($credDom -eq $Meta.value.o365_OPDomain)){
                if($Meta.value.OP_ExADRoot){
                    if(!$Meta.value.OP_rgxEMSComputerName){
                        write-verbose "(adding XXXMeta.OP_rgxEMSComputerName value)"
                        # build vari that will match curlyhoward.cmw.internal|lynms650.global.ad.toro.com etc
                        set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'OP_rgxEMSComputerName' = "^\w*\.$([Regex]::Escape($Meta.value.OP_ExADRoot))$"} ) ;
                    } ;
                } else {
                    throw "Missing `$$($Meta.value.o365_Prefix).OP_ExADRoot value.`nProfile hasn't loaded proper tor-incl-infrastrings file)!"
                } ;
            } ; # if-E $credDom
        } ; # loop-E
        #>
        # non-looping vers:
        #$TenOrg = get-TenantTag -Credential $Credential ;
        #.OP_ExADRoot
        if( (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName ){

        } else {
            #.OP_rgxEMSComputerName
            if((Get-Variable  -name "$($TenOrg)Meta").value.OP_ExADRoot){
                set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'OP_rgxEMSComputerName' = "^\w*\.$([Regex]::Escape((Get-Variable  -name "$($TenOrg)Meta").value.OP_ExADRoot))$"} )
            } else {
                $smsg = "Missing `$$((Get-Variable  -name "$($TenOrg)Meta").value.o365_Prefix).OP_ExADRoot value.`nProfile hasn't loaded proper tor-incl-infrastrings file)!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;
        } ;
    } ;  # BEG-E

    PROCESS{
        $verbose = ($VerbosePreference -eq "Continue") ;
        # if we're using ems-style BasicAuth, clear incompatible existing Rems PSS's
        # ComputerName      : curlyhoward.cmw.internal ;  ComputerType      : RemoteMachine ;  State             : Opened ;  ConfigurationName : Microsoft.Exchange ;  Availability      : Available ;  Name              : Session1 ;   ;
        $rgxRemsPSSName = "^(Session\d|Exchange\d{4})$" ;
        $Rems2Good = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') } ;
        # Computername wrong fqdn suffix
        #$Rems2WrongOrg = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND (-not($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName)) -AND ($_.Availability -eq 'Available') } ;
        # above is seeing outlook EXO conns as wrong org, exempt them too: .ComputerName -match $rgxExoPsHostName
        $Rems2WrongOrg = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND (
            $_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND (
            ( -not($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) ) -AND (
            -not($_.ComputerName -match $rgxExoPsHostName)) ) -AND ($_.Availability -eq 'Available')
        } ;
        $Rems2Broken = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND (
                $_.Name -match $rgxRemsPSSName) -AND ($_.State -like "*Broken*") } ;
        $Rems2Closed = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND (
                $_.Name -match $rgxRemsPSSName) -AND ($_.State -like "*Closed*") } ;

        write-verbose "(Removing $($Rems2Broken.count) Broken sessions)" ;
        if ($Rems2Broken.count -gt 0){ for ($index = 0 ;$index -lt $Rems2Broken.count ;$index++){Remove-PSSession -session $Rems2Broken[$index]}  };
        write-verbose "(Removing $($Rems2Closed.count) Closed sessions)" ;
        if ($Rems2Closed.count -gt 0){for ($index = 0 ;$index -lt $Rems2Closed.count ; $index++){Remove-PSSession -session $Rems2Closed[$index] } } ;
        write-verbose "(Removing $($Rems2WrongOrg.count) sessions connected to the WRONG ORG)" ;
        if ($Rems2WrongOrg.count -gt 0){for ($index = 0 ;$index -lt $Rems2WrongOrg.count ; $index++){Remove-PSSession -session $Rems2WrongOrg[$index] } } ;
        # preclear until proven *up*
        $bExistingREms = $false ;

        if( Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') } ){

            $bExistingREms= $true ;
            write-verbose "(Authenticated to Ex20XX:$($Credential.username.split('\')[0].tostring()))" ;

        } else {
            write-verbose "(NOT Authenticated to Credentialed Ex20XX Org:$($Credential.username.split('\')[0].tostring()))" ;
            $tryNo=0 ; $1F=$false ;
            Do {
                if($1F){Sleep -s 5} ;
                $tryNo++ ;
                write-host "." -NoNewLine; if($tryNo -gt 1){Start-Sleep -m (1000 * 5)} ;
                write-verbose "$((get-date).ToString('HH:mm:ss')):Reconnecting:No existing PSSESSION matching`n (ConfigurationName -eq 'Microsoft.Exchange') -AND (Name -match $($rgxRemsPSSName)) -AND ($_.ComputerName -match $((Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName))`nwith valid Open/Availability:$((Get-PSSession | where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ($_.Name -match $rgxRemsPSSName)} |ft -a Id,Name,ComputerName,ComputerType,State,ConfigurationName,Availability|out-string).trim())" ;
                Disconnect-Ex2010 ; Disconnect-PssBroken ;Start-Sleep -Seconds 3;

                $bExistingREms = $false ;

                Connect-Ex2010xo -credential:$($Credential) ;

                $1F=$true ;
                if($tryNo -gt $DoRetries ){throw "RETRIED EX20XX CONNECT $($tryNo) TIMES, ABORTING!" } ;
            } Until ( Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') } ) ;

        } ;

    } ;  # PROC-E
    END {
        if($bExistingREms -eq $false){
            if( Get-PSSession | where-object {$_.ConfigurationName -eq "Microsoft.Exchange" -AND $_.Name -match $rgxRemsPSSName -AND $_.State -eq "Opened" -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') }  ){
                $bExistingREms= $true ;
            } else {
                write-error "(Credential mismatch:disconnecting from existing Ex201X:$($eEXO.Identity) tenant)" ;
                Disconnect-Ex2010 ;
                $bExistingREms = $false ;
            } ;
        } ;
    } ; # END-E
}

#*------^ Reconnect-Ex2010XO.ps1 ^------


#*------v remove-EMSLocalModule.ps1 v------
Function remove-EMSLocalModule {
  <#
    .SYNOPSIS
    remove-EMSLocalModule - remove/unload local server bin-module-based ExchOnPrem Mgmt Shell connection ; validated Exch2016)
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    Version     : 1.0.0
    CreatedDate : 2021-04-15
    FileName    : remove-EMSLocalModule()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell,Exchange,Exchange-2013,Exchange-2016
    REVISIONS   :
    * 9:42 AM 7/27/2021 add verbose to *-PsTitleBar calls
    * 10:03 AM 4/16.2.31 init vers
    .DESCRIPTION
    remove-EMSLocalModule - remove/unload local server bin-module-based ExchOnPrem Mgmt Shell connection ; validated Exch2016)
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    remove-EMSLocalModule ;
    .LINK
    https://github.com/tostka/verb-Ex2010/
    #>
    [CmdletBinding()]
    #[Alias()]
    Param()  ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ;
    } ;  # BEG-E
    PROCESS{
        $error.clear() ;
        TRY {
            if($tMod = get-module ([System.Net.Dns]::gethostentry($(hostname))).hostname -ea 0){
                write-verbose "(Removing matched EMS module already loaded:)`n$(($tMod | ft -auto ModuleType,Version,Name,ExportedCommands|out-string).trim())" ;
                $tMod | Remove-Module ;
            } else {
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(No matching loaded local Exchange Mgmt Shell binary module found)" ;
            } ;
        } CATCH {
            $ErrTrapd = $_ ;
            write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            #-=-record a STATUSERROR=-=-=-=-=-=-=
            $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(gv passstatus -scope Script -ea 0 ){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script  -ea 0 ){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
            #-=-=-=-=-=-=-=-=
        } ;
        # 7:54 AM 11/1/2017 add titlebar tag
        if(gcm Remove-PSTitleBar-PSTitleBar -ea 0 ){Remove-PSTitleBar-PSTitleBar 'EMSL' -verbose:$($VerbosePreference -eq "Continue") ;} ;
        # tag E10IsDehydrated
        $Global:ExOPIsDehydrated = $null ;
    } ;  # PROC-E
    END {
        <#
        $tMod = $null ;
        if($tMod = GET-MODULE ([System.Net.Dns]::gethostentry($(hostname))).HOSTNAME){
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(local EMS module loaded:)`n$(($tMod | ft -auto ModuleType,Version,Name,ExportedCommands|out-string).trim())" ;
        } else {
            throw "Unable to resolve target local EMS module:GET-MODULE $([System.Net.Dns]::gethostentry($(hostname))).HOSTNAME)" ;
        } ;
        #>
    }
}

#*------^ remove-EMSLocalModule.ps1 ^------


#*------v resolve-ExchangeServerVersionTDO.ps1 v------
function resolve-ExchangeServerVersionTDO {
    <#
    .SYNOPSIS
    resolve-ExchangeServerVersionTDO - Resolves the ExchangeVersion details from a returned get-ExchangeServer, whether local undehydrated ('Microsoft.Exchange.Data.Directory.Management.ExchangeServer') or remote EMS ('System.Management.Automation.PSObject')
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2024-05-22
    FileName    : resolve-ExchangeServerVersionTDO
    License     : (none asserted)
    Copyright   : (none asserted)
    Github      : https://github.com/tostka/verb-Ex2010
    Tags        : Powershell,ExchangeServer,Version
    AddedCredit : Bruno Lopes (brunokktro )
    AddedWebsite: https://www.linkedin.com/in/blopesinfo
    AddedTwitter: @brunokktro / https://twitter.com/brunokktro
    REVISIONS
    * 1:47 PM 7/9/2024 CBA github field correction
    * 1:22 PM 5/22/2024init
    .DESCRIPTION
    resolve-ExchangeServerVersionTDO - Resolves the ExchangeVersion details from a returned get-ExchangeServer, whether local undehydrated ('Microsoft.Exchange.Data.Directory.Management.ExchangeServer') or remote EMS ('System.Management.Automation.PSObject')
    Returns a  PSCustomObject to the pipleine with the following properties:

        isEx2019             : [boolean]
        isEx2016             : [boolean]
        isEx2007             : [boolean]
        isEx2003             : [boolean]
        isEx2000             : [boolean]
        ExVers               : [string] 'Ex2010'
        ExchangeMajorVersion : [string] '14.3'
        isEx2013             : [boolean]
        isEx2010             : [boolean]

    Extends on sample code by brunokktro's Get-ExchangeEnvironmentReport.ps1

    .PARAMETER ExchangeServer
    Object returned by a get-ExchangeServer command
    .OUTPUT
    PSCustomObject version summary.
    .EXAMPLE
    PS> write-verbose 'Resolve the local ExchangeServer object to version description, and assign to `$returned' ;     
    PS> $returned = resolve-ExchangeServerVersionTDO -ExchangeServer (get-exchangeserver $env:computername) 
    PS> write-verbose "Expand returned populated properties into local variables" ; 
    PS> $returned.psobject.properties | ?{$_.value} | %{ set-variable -Name $_.name -Value $_.value -verbose } ; 
        
        VERBOSE: Performing the operation "Set variable" on target "Name: ExVers Value: Ex2010".
        VERBOSE: Performing the operation "Set variable" on target "Name: ExchangeMajorVersion Value: 14.3".
        VERBOSE: Performing the operation "Set variable" on target "Name: isEx2010 Value: True".

    Demo retrieving get-exchangeserver, assigning to output, processing it for version info, and expanding the populated returned values to local variables. 
    .LINK
    https://github.com/brunokktro/ExchangeServer/blob/master/Get-ExchangeEnvironmentReport.ps1
    .LINK
    https://github.com/tostka/verb-Ex2010
    #>
    [CmdletBinding()]
    #[Alias('rvExVers')]
    PARAM(
        [Parameter(Mandatory = $true,Position=0,HelpMessage="Object returned by a get-ExchangeServer command[-ExchangeServer `$ExObject]")]
            [array]$ExchangeServer
    ) ;
    BEGIN {
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        $verbose = $($VerbosePreference -eq "Continue")
        $sBnr="#*======v $($CmdletName): v======" ;
        write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
    }
    PROCESS {
        foreach ($item in $ExchangeServer){
            
            if($host.version.major -ge 3){$oReport=[ordered]@{Dummy = $null ;} }
            else {$oReport = New-Object Collections.Specialized.OrderedDictionary} ;
            If($oReport.Contains("Dummy")){$oReport.remove("Dummy")} ;
            #$oReport.add('sessionid',$sid) ; # add a static pre-stocked value
            # then loop out the commonly-valued/typed entries
            
            $fieldsnull = 'isEx2019','isEx2016','isEx2007','isEx2003','isEx2000','ExVers','ExchangeMajorVersion' ; $fieldsnull | % { $oReport.add($_,$null) } ;
            #$fieldsArray = 'key4','key5','key6' ; $fieldsArray | % { $oReport.add($_,@() ) } ;
            #$oReport.add('',) ; # explicit variable
            #$oReport.add('','') ; # explicit value
            #$oReport.key1 = 'value' ;  # assign value 

            # this may be undehydrated $RemoteExchangePath, or REMS dehydrated, w fundementally diff properties
            switch($item.admindisplayversion.gettype().fullname){

                'Microsoft.Exchange.Data.ServerVersion'{
                    #    '6.0'  = @{Long = 'Exchange 2000'; Short = 'E2000' }
                         #  '6.5'  = @{Long = 'Exchange 2003'; Short = 'E2003' }
            #               '8'    = @{Long = 'Exchange 2007'; Short = 'E2007' }
            #               '14'   = @{Long = 'Exchange 2010'; Short = 'E2010' } # Ex2010 version.Minor == SP#
            #               '15'   = @{Long = 'Exchange 2013'; Short = 'E2013' } # Ex2010 version.Minor == SP#
            #               '15.1' = @{Long = 'Exchange 2016'; Short = 'E2016' } 
            #               '15.2' = @{Long = 'Exchange 2019'; Short = 'E2019' } #2019-05-17 TST Exchange Server 2019 added
                    #
                    if ($item.AdminDisplayVersion.Major -eq 6) {
                        # 6(.0) == Ex2000  ; 6.5 == Ex2003 
                        $oReport.ExchangeMajorVersion = [double]('{0}.{1}' -f $item.AdminDisplayVersion.Major, $item.AdminDisplayVersion.Minor)
                        $ExchangeSPLevel = $item.AdminDisplayVersion.FilePatchLevelDescription.Replace('Service Pack ', '')
                    } elseif ($item.AdminDisplayVersion.Major -eq 15 -and $item.AdminDisplayVersion.Minor -ge 1) {
                        # 15.1 == Ex2016 ; 15.2 == Ex2019
                        $oReport.ExchangeMajorVersion = [double]('{0}.{1}' -f $item.AdminDisplayVersion.Major, $item.AdminDisplayVersion.Minor)
                        $ExchangeSPLevel = 0
                    } else {
                        # 8(.0) == Ex2007 ; 14(.0) == Ex2010 ; 15(.0) == Ex2013 
                        $oReport.ExchangeMajorVersion = $item.AdminDisplayVersion.Major ; 
                        $ExchangeSPLevel = $item.AdminDisplayVersion.Minor ; 
                    } ; 

                    $oReport.isEx2000 = $oReport.isEx2003 = $oReport.isEx2007 = $oReport.isEx2010 = $oReport.isEx2013 = $oReport.isEx2016 = $oReport.isEx2019 = $false ; 
                    $oReport.ExVers = $null ; 
                    switch ([string]$oReport.ExchangeMajorVersion) {
                        '15.2' { $oReport.isEx2019 = $true ; $oReport.ExVers = 'Ex2019' }
                        '15.1' { $oReport.isEx2016 = $true ; $oReport.ExVers = 'Ex2016'}
                        '15' { $oReport.isEx2013 = $true ; $oReport.ExVers = 'Ex2013'}
                        '14' { $oReport.isEx2010 = $true ; $oReport.ExVers = 'Ex2010'}
                        '8' { $oReport.isEx2007 = $true ; $oReport.ExVers = 'Ex2007'}  
                        '6.5' { $oReport.isEx2003 = $true ; $oReport.ExVers = 'Ex2003'} 
                        '6' {$oReport.isEx2000 = $true ; $oReport.ExVers = 'Ex2000'} ;
                        default { 
                            $smsg = "UNRECOGNIZED ExchangeServer.AdminDisplayVersion.Major.Minor string:$($oReport.ExchangeMajorVersion)! ABORTING!" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            THROW $SMSG ; 
                            BREAK ; 
                        }
                    } ; 
                }
                'System.String'{
                    $oReport.ExVers = $oReport.isEx2000 = $oReport.isEx2003 = $oReport.isEx2007 = $oReport.isEx2010 = $oReport.isEx2013 = $oReport.isEx2016 = $oReport.isEx2019 = $false ; 
                    if([double]$ExVersNum = [regex]::match($item.AdminDisplayVersion,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
                        switch -regex ([string]$ExVersNum) {
                            '15.2' { $oReport.isEx2019 = $true ; $oReport.ExVers = 'Ex2019' }
                            '15.1' { $oReport.isEx2016 = $true ; $oReport.ExVers = 'Ex2016'}
                            '15.0' { $oReport.isEx2013 = $true ; $oReport.ExVers = 'Ex2013'}
                            '14.*' { $oReport.isEx2010 = $true ; $oReport.ExVers = 'Ex2010'}
                            '8.*' { $oReport.isEx2007 = $true ; $oReport.ExVers = 'Ex2007'}
                            '6.5' { $oReport.isEx2003 = $true ; $oReport.ExVers = 'Ex2003'}
                            '6' {$oReport.isEx2000 = $true ; $oReport.ExVers = 'Ex2000'} ;
                            default {
                                $smsg = "UNRECOGNIZED ExchangeServer.AdminDisplayVersion string:$($item.AdminDisplayVersion)! ABORTING!" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                THROW $SMSG ;
                                BREAK ;
                            }
                        } ; 
                        $smsg = "Need `$oReport.ExchangeMajorVersion as well (emulating output of non-dehydrated)" 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $oReport.ExchangeMajorVersion = $ExVersNum
                    }else {
                        $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$item.version:$($item.version)!" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        throw $smsg ; 
                        break ; 
                    } ;
                } ;
                default {
                    # $item.admindisplayversion.gettype().fullname
                    $smsg = "Unable to detect `$item.admindisplayversion.gettype():$($item.admindisplayversion.gettype().fullname)!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    throw $smsg ; 
                    break ; 
                };  
            }
            $smsg = "(returning  results for $($item.name) to pipeline)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

            New-Object -TypeName PsObject -Property $oReport | write-output ; 
        } ; 
    } # PROC-E
    END{
        write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
    }
}

#*------^ resolve-ExchangeServerVersionTDO.ps1 ^------


#*------v resolve-RecipientEAP.ps1 v------
function resolve-RecipientEAP {
    <#
    .SYNOPSIS
    resolve-RecipientEAP.ps1 - Resolve a recipient against the onprem local EmailAddressPolicies, and return the matching/applicable EAP Policy object
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-08-18
    FileName    : resolve-RecipientEAP.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 2:16 PM 6/24/2024: rem'd out #Requires -RunasAdministrator; sec chgs in last x mos wrecked RAA detection 
    * 11:21 AM 9/16.2.31 string clean
    * 3:27 PM 8/23/2021 revised patched in new preview-EAPUpdate() support; added 
    default EAP cheatsheet output dump to console; suppress get-EAP warning ; 
    revised recipientfilter support to simple ($(existingRcpFltr) -AND (alias -eq $rcp.alias)).
    Much less complicated, should work on any eap with a recip fltr. 
    * 3:00 PM 8/19/2021 tested, fixed perrotpl issue (overly complicated rcpfltr), 
    pulls a single recipient back on a match on any domain. Considered running a 
    blanket 'get all matches' on each, and then post-filtering for target user(s) 
    but: filtered to a single recip in the rcptfilter, takes 8s for @toro.com; for 
    all targeted is's 1m+. And, just running for broad matches, wo considering 
    priority isn't valid: higher priority matches shut down laters, so you *need* 
    to run them in order, one at a time, and quit on first match. You can't try to 
    datacolect & postfilter wo considering priority, given user may match mult EAPs.
    * 11:11 AM 8/18/2021 init
    .DESCRIPTION
    resolve-RecipientEAP.ps1 - Resolve an array of recipients against the onprem local EmailAddressPolicies, and return the matching/applicable EAP Policy object
    Runs a single recipient (rather than an array) because you really can't pre-collect full populations and stop. Need to run the EAPs in priority order, filter population returned, and quit on first match.
    .PARAMETER  Recipient
    Array of recipient descriptors: displayname, emailaddress, UPN, samaccountname[-recip some.user@domain.com]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER useAltFilter
    Switch to attempt broad append '(existing) -AND (Alias -eq '`$(alias)' to eap.recipientfilter, rather than fancy search/replc on clauses (defaulted TRUE) [-useAltFilter]
    .INPUTS
    None. Does not accepted piped input
    .OUTPUTS
    System.Management.Automation.PSCustomObject of matching EAP
    .EXAMPLE
    PS> $matchedEAP = resolve-RecipientEAP -rec SOMEACCT@DOMAIN.COM -verbose ;
    PS> if($matchedEAP){"User matches $($matchedEAP.name"} else { "user matches *NO* existing EAP! (re-run with -verbose for further details)" } ; 
    .EXAMPLE
    "user1@domain.com","user2@domain.com"|%{resolve-RecipientEAP -rec $_ -verbose} ; 
    Foreach-object loop an array of descriptors 
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    ###Requires -Version 5
    ###Requires -Modules verb-Ex2010 - disabled, moving into the module
    ##Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    PARAM(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of recipient descriptors: displayname, emailaddress, UPN, samaccountname[-recip some.user@domain.com]")]
        #[ValidateNotNullOrEmpty()]
        #[Alias('ALIAS1', 'ALIAS2')]
        $Recipient,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2,
        [Parameter(HelpMessage="Display EmailAddressPolicy format strings 'cheatsheet' (defaults true) [-showCheatsheet]")]
        [switch] $showCheatsheet=$true
    ) ;
    BEGIN{
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        $rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ; 
        $rgxDName = "^([a-zA-Z]{2,}\s[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)" ; 
        $rgxSamAcctNameTOR = "^\w{2,20}$" ; # up to 20k, the limit prior to win2k
        #$rgxSamAcctName = "^[^\/\\\[\]:;|=,+?<>@?]+$" # no char limit ;
        $propsEAPFiltering = 'EmailAddressPolicyEnabled','CustomAttribute5','primarysmtpaddress','Office','distinguishedname','Recipienttype','RecipientTypeDetails' ; 
        $rgxEmailDirLegalChars = "[0-9a-zA-Z-._+&']" ; 
        $hCheatSheet = @"

Email Address Policy AddressTemplate format variables:
|Vari |Value
|-----|-------------------------------------|
|%d   |Display name                       
|%g   |Given name                         
|%i   |Middle initial                      
|%m   |Exchange alias                      
|%rxy |Replace all occurrences of x with y 
|%rxx |Remove all occurrences of x         
|%s   |Surname                 
|%ng  |The first n letters of the givenname.
|%ns  |The first n letters of the surname. 

All smtpaddr-illegal chars are dropped from source string. 
Commonly-permitted SmtpAddrChars:
$($rgxEmailDirLegalChars)
(RFC 5322 technically permits broader set, but frequently blocked as risks)

"@ ; 
        
        rx10 -Verbose:$false ; 
        #rxo  -Verbose:$false ; cmsol  -Verbose:$false ;

        # move the properties out to a separate vari
        [array]$eapprops = 'name','RecipientFilter','RecipientContainer','EnabledPrimarySMTPAddressTemplate','EnabledEmailAddressTemplates',
            'DisabledEmailAddressTemplates','Enabled' ; 
        # append an expression that if/then's Priority text value: coercing IsNumeric()'s to [int], else - only non-Numeric is 'Default' - replacing that Priority with [int](EAPs.count)+1
        $eapprops += @{Name="Priority";Expression={ 
            if($_.priority.trim() -match "^[-+]?([0-9]*\.[0-9]+|[0-9]+\.?)$"){
                [int]$_.priority 
            } else { 
                [int]($eaps.count+1) 
            }
            } } ; 
       
        # pull EAP's and sub sortable integer values for Priority (Default becomes EAPs.count+1)
        $smsg = "(polling:Get-EmailAddressPolicy...)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        $sw = [Diagnostics.Stopwatch]::StartNew();
        # use -warningaction silentlycontinue to suppress the 'WARNING: Recipient policy objects that don't contain e-mail address won't be shown unless you include the IncludeMailboxSettingOnlyPolicy'
        $eaps = Get-EmailAddressPolicy -WarningAction 0 ;
        $sw.Stop() ;
        $eaps = $eaps | select $eapprops | sort Priority  ; 
        
        $smsg = ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                
        
    } 
    PROCESS{
       
        $hSum = [ordered]@{
            OPRcp = $OPRcp;
            xoRcp = $xoRcp;
        } ;
                    
        $sBnr="===vInput: '$($Recipient)' v===" ;
        $smsg = $sBnr ;        
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        $xMProps="samaccountname","windowsemailaddress","DistinguishedName","Office","RecipientTypeDetails" ;
        
        $pltgM=[ordered]@{} ; 
        $smsg = "processing:'identity':$($Recipient)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

        $pltgM.add('identity',$Recipient) ;
            
        $smsg = "get-recipient w`n$(($pltgM|out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

        rx10 -Verbose:$false -silent ;
        $error.clear() ;

        $sw = [Diagnostics.Stopwatch]::StartNew();
        $hSum.OPRcp=get-recipient @pltgM -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'}
        $sw.Stop() ;
        $smsg = ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 


        if($hSum.OPRcp){
            $smsg = "`Matched On-Premesis Recipient:`n$(($hSum.OPRcp|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

            $hMsg=@"
Recipient $($hSum.OpRcp.primarysmtpaddress) has the following EmailAddressPolicy-related settings:

$(($hSum.OPRcp | fl $propsEAPFiltering|out-string).trim())

The above settings need to exactly match one or more of the EAP's to generate the desired match...

"@ ;
            
            $smsg = $hMsg ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            if($hSum.OPRcp.EmailAddressPolicyEnabled -eq $false){
                $smsg = "Recipient $($hSum.OpRcp.primarysmtpaddress) is DISABLED for EAP use:`n" ; 
                $smsg += "$(($hSum.OPRcp | fl EmailAddressPolicy|out-string).trim())`n`n" ; 
                $smsg += "This user will *NOT* be governed by any EAP until this value is reset to `$true!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $smsg = "Recipient $($hSum.OpRcp.primarysmtpaddress) properly has:`n$(($hSum.OPRcp | fl EmailAddressPolicyEnabled|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            }  ;

            $bBadRecipientType =$false ;
            switch -regex ($hSum.OPRcp.recipienttype){
                "UserMailbox" {
                    $smsg = "'UserMailbox'"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } 
                "MailUser" {
                    $smsg = "'MailUser'" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ;
                default {
                    $smsg = "Unsupported RecipientType:($hSum.OPRcp.recipienttype). EXITING!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $bBadRecipientType = $true ;
                    Break ; 
                }
            }

            if(!$bBadRecipientType ){
                $error.clear() ;
                TRY {
                   
                    $matchedEAP = $null ; 
                    $propsEAP = 'name','RecipientFilter','RecipientContainer','Priority','EnabledPrimarySMTPAddressTemplate',
                        'EnabledEmailAddressTemplates','DisabledEmailAddressTemplates','Enabled' ; 
                    $aliasmatch = $hSum.OPRcp.alias ;

                    write-host "`n(Comparing to $(($Eaps|measure).count) EmailAddressPolicies for filter-match...)" ;
                    foreach($eap in $eaps){
                        if(!$verbose){
                            write-host "." -NoNewLine ;
                        } ; 
                        $smsg = "`n`n==$($eap.name):$($eap.RecipientFilter)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        
                        # try a simple (existing) -AND "Alias -eq '$($aliasmatch)'" filter mod
                        $tmpfilter = "($($eap.recipientfilter)) -and (Alias -eq '$($aliasmatch)')" ; 
                        
                        $smsg = "using `$tmpfilter recipientFilter:`n$($tmpfilter)" ;  
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        $pltGRcpV=[ordered]@{
                            RecipientPreviewFilter=$tmpfilter ;
                            OrganizationalUnit=$eap.RecipientContainer ;
                            resultsize='unlimited';
                            ErrorAction='STOP';
                        } ;
                        $smsg = "get-recipient w`n$(($pltGRcpV|out-string).trim())`n$(($pltGRcpV.RecipientPreviewFilter|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        
                        $sw = [Diagnostics.Stopwatch]::StartNew();
                        if($rcp =get-recipient @pltGRcpV| ?{$_.alias -eq $aliasmatch} ){
                            $sw.Stop() ;
                            write-host "MATCHED!:$($Eap.name)`n" ;
                            $matchedEAP = $eap ;
                            $smsg = ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
                            $smsg = "Matched OnPremRecipient $($Sum.OPRcp.alias) to EAP Preview grp:$($rcp.primarysmtpaddress)`n" ; 
                            $smsg += "filtered under EmailAddressPolicy:`n$(($eap | fl ($propsEAP |?{$_ -ne 'EnabledEmailAddressTemplates'}) |out-string).trim())`n" ; 
                            $smsg += "EnabledEmailAddressTemplates:`n$(($eap | select -expand EnabledEmailAddressTemplates |out-string).trim())`n" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            $genEml = preview-EAPUpdate  -eap $matchedEAP -Recipient $hSum.OPRcp -Verbose:($VerbosePreference -eq 'Continue')
                            if($geneml -ne $hSum.OPRcp.PrimarySmtpAddress){
                                $smsg = "`n===Specified recip's PrimarySmtpAddress ($hSum.OPRcp.PrimarySmtpAddress))`n"
                                $smsg += "does *not* appear to match specified template!`n" ; 
                                $smsg += "*manualy review* the template specs *validate*`n"
                                $smsg += "that the desired scheme is being applied!`n==="
                                #write-warning $smsg ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            }else {
                                $smsg = "`n===PrimarysmtpAddr $($hSum.OPRcp.PrimarySmtpAddress))`n"
                                $smsg += "roughly conforms to specified template primary addr`n" ;
                                $smsg += "$($matchedEAP.EnabledPrimarySMTPAddressTemplate)...===`n" ;
                                #write-host -foregroundcolor Green $smsg ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            } ;

                            break ;
                        } else {
                            $sw.Stop() ;
                            $smsg = ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        } ;
                    }; # E-loop
                    
                    if($showCheatsheet){
                        write-host $hCheatSheet
                    } ; 

                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ; 
                } ; 

            } else { 
                 $smsg = "-Recipient:$($Recipient) is of an UNSUPPORTED type by this script! (only Mailbox|MailUser are supported)"   ; 
                 if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
            
        } else { 
            $smsg = "(no matching EXOP recipient object:$($Recipient))"   
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
    } #  PROC-E
    END{
        if( $matchedEAP){
            $matchedEAP | write-output ; 
        } else { 
            $smsg = "Failed to resolve specified recipient $($user) to a matching EmailAddressPolicy" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $false | write-output ;
        } ; 
        $smsg = "$($sBnr.replace('=v','=^').replace('v=','^='))" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
     }

     #*======^ END SUB MAIN ^======
 }

#*------^ resolve-RecipientEAP.ps1 ^------


#*------v rx10cmw.ps1 v------
function rx10cmw {
    <#
    .SYNOPSIS
    rx10cmw - Reonnect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Reconnect-EX2010 - Reonnect to specified on-prem Exchange
    .EXAMPLE
    rx10cmw
    #>
    [CmdletBinding()] 
        [Alias('rxOPcmw')]
    Param()
    $Verbose = ($VerbosePreference -eq 'Continue') ;
    $pltGHOpCred=@{TenOrg="CMW" ;userrole=@('ESVC','LSVC','SID') ;verbose=$($verbose)} ;
    if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
        ReConnect-EX2010 -cred $OPCred -Verbose:($VerbosePreference -eq 'Continue') ; 
    } else {
        $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole $($UserRole -join '|') value!"
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
        exit ;
    } ;
}

#*------^ rx10cmw.ps1 ^------


#*------v rx10tol.ps1 v------
function rx10tol {
    <#
    .SYNOPSIS
    rx10tol - Reonnect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Reconnect-EX2010 - Reonnect to specified on-prem Exchange
    .EXAMPLE
    rx10tol
    #>
    [CmdletBinding()] 
    [Alias('rxOPtol')]
    Param()
    $Verbose = ($VerbosePreference -eq 'Continue') ;
    $pltGHOpCred=@{TenOrg="TOL" ;userrole=@('ESVC','LSVC','SID') ;verbose=$($verbose)} ;
    if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
        ReConnect-EX2010 -cred $OPCred -Verbose:($VerbosePreference -eq 'Continue') ; 
    } else {
        $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole $($UserRole -join '|') value!"
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
        exit ;
    } ;
}

#*------^ rx10tol.ps1 ^------


#*------v rx10tor.ps1 v------
function rx10tor {
    <#
    .SYNOPSIS
    rx10tor - Reonnect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Reconnect-EX2010 - Reonnect to specified on-prem Exchange
    .EXAMPLE
    rx10tor
    #>
    [CmdletBinding()] 
    [Alias('rxOPtor')]
    Param()
    $Verbose = ($VerbosePreference -eq 'Continue') ;
    $pltGHOpCred=@{TenOrg="TOR" ;userrole=@('ESVC','LSVC','SID') ;verbose=$($verbose)} ;
    if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
        ReConnect-EX2010 -cred $OPCred -Verbose:($VerbosePreference -eq 'Continue') ; 
    } else {
        $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole $($UserRole -join '|') value!"
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
        exit ;
    } ;
}

#*------^ rx10tor.ps1 ^------


#*------v test-ExOPPSession.ps1 v------
Function test-ExOPPSession {
  <#
    .SYNOPSIS
    test-ExOPPSession - Does a *simple* - NO-ORG REVIEW - validation of functional PSSession with: ConfigurationName:'Microsoft.Exchange' -AND Name match  '^(Exchange2010|Session\sfor\simplicit\sremoting\smodule\sat\s.*)' -AND State:'Opened' -AND Availability:'Available' -AND can gcm -name 'Add-ADPermission'
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    Version     : 1.0.0
    CreatedDate : 2021-04-15
    FileName    : test-ExOPPSession()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell,Exchange,Exchange-2013,Exchange-2016
    REVISIONS   :
    * 12:30 PM 5/3/2021 init vers ; revised rgxRemsPSSName
    .DESCRIPTION
    test-ExOPPSession - Does a *simple* - NO-ORG REVIEW - validation of functional PSSession with: ConfigurationName:'Microsoft.Exchange' -AND Name match  '^(Exchange2010|Session\sfor\simplicit\sremoting\smodule\sat\s.*)' -AND State:'Opened' -AND Availability:'Available' -AND can gcm -name 'Add-ADPermission'.
    This does *NO* validation that any specific EXOnPrem org is attached! It just validates that an existing PSSession *exists* that *generically* matches a Remote Exchange Mgmt Shell connection in a usable state. Use case is scripts/functions that *assume* you've already pre-established a suitable connection, and just need to pre-test that *any* PSS is already open, before attempting commands. 
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    System.Management.Automation.Runspaces.PSSession. Returns the functional PSSession object(s)
    .EXAMPLE
    PS> if(test-ExOPPSession){'OK'} else { 'NOGO!'}  ;
    .LINK
    https://github.com/tostka/verb-Ex2010/
    #>
    [CmdletBinding()]
    #[Alias()]
    Param()  ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ;
        $rgxRemsPSSName = "^(Exchange\d{4})$" ; 
        $testCommand = 'Add-ADPermission' ; 
        $propsREMS = 'Id','Name','ComputerName','ComputerType','State','ConfigurationName','Availability' ; 
    } ;  # BEG-E
    PROCESS{
        $error.clear() ;
        TRY {
            if($RemsGood = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.Availability -eq 'Available') }){
                $smsg = "valid EMS PSSession found:`n$(($RemsGood|ft -a $propsREMS |out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-VERBOSE "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                
                if($tmod = (get-command -name $testCommand).source){
                    $smsg = "(confirmed PSSession open/available, with $($testCommand) available)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $RemsGood | write-output ; ;
                } else { 
                    throw "NO FUNCTIONAL PSSESSION FOUND!" ; 
                } ; 
            } else {
                throw "No existing open/available Remote Exchange Management Shell found!"
            } ;
        } CATCH {
            $ErrTrapd = $_ ;
            write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            #-=-record a STATUSERROR=-=-=-=-=-=-=
            $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(gv passstatus -scope Script -ea 0 ){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script  -ea 0 ){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
            #-=-=-=-=-=-=-=-=
        } ;
        
    } ;  # PROC-E
    END {}
}

#*------^ test-ExOPPSession.ps1 ^------


#*------v test-EXOPStatus.ps1 v------
function test-EXOPConnection {
    <#
    .SYNOPSIS
    test-EXOPConnection.ps1 - Validate EXOP connection, and that the proper Tenant is connected (as per provided Credential)
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2022-06-24
    FileName    : test-EXOPConnection.ps1
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-EXO
    Tags        : Powershell
    REVISIONS
    * 8:40 AM 6/23/2023 rmvd req's: ps3, just clutters, most mods aren't using a rev req drop it.
    * 2:23 PM 4/17/2023 pulled MinNoWinRMVersion refs (spurious)
    *11:44 AM 9/12/2022 init ; port Test-EXO2Connection to EXOP support
    .DESCRIPTION
    test-EXOPConnection.ps1 - Validate EXOP connection, and that the proper Tenant is connected (as per provided Credential)
    .PARAMETER Credential
    Credential to be used for connection
    .OUTPUT
    System.Boolean
    .EXAMPLE
    PS> $oRet = test-EXOPConnection -verbose ; 
    PS> if($oRet.Valid){
    PS>     $pssEXOP = $oRet.PsSession ; 
    PS>     write-host 'Validated EXOv2 Connected to Tenant aligned with specified Credential'
    PS> } else { 
    PS>     write-warning 'NO EXO USERMAILBOX TYPE LICENSE!'
    PS> } ; 
    Evaluate EXOP connection status with verbose output
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    [CmdletBinding()]
     Param(
        [Parameter(Mandatory=$False,HelpMessage="Credentials [-Credentials [credential object]]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credOpTORSID
    )
    BEGIN {
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        
        
        #*------v PSS & GMO VARIS v------
        # get-pssession session varis
        # select key differentiating properties:
        $pssprops = 'Id','ComputerName','ComputerType','State','ConfigurationName','Availability', 
            'Description','Guid','Name','Path','PrivateData','RootModuleModule', 
            @{name='runspace.ConnectionInfo.ConnectionUri';Expression={$_.runspace.ConnectionInfo.ConnectionUri} },  
            @{name='runspace.ConnectionInfo.ComputerName';Expression={$_.runspace.ConnectionInfo.ComputerName} },  
            @{name='runspace.ConnectionInfo.Port';Expression={$_.runspace.ConnectionInfo.Port} },  
            @{name='runspace.ConnectionInfo.AppName';Expression={$_.runspace.ConnectionInfo.AppName} },  
            @{name='runspace.ConnectionInfo.Credentialusername';Expression={$_.runspace.ConnectionInfo.Credential.username} },  
            @{name='runspace.ConnectionInfo.AuthenticationMechanism';Expression={$_.runspace.ConnectionInfo.AuthenticationMechanism } },  
            @{name='runspace.ExpiresOn';Expression={$_.runspace.ExpiresOn} } ; 
        
        if(-not $EXoPConfigurationName){$EXoPConfigurationName = "Microsoft.Exchange" };
        if(-not $rgxEXoPrunspaceConnectionInfoAppName){$rgxEXoPrunspaceConnectionInfoAppName = '^/(exadmin|powershell)$'}; 
        if(-not $EXoPrunspaceConnectionInfoPort){$EXoPrunspaceConnectionInfoPort = '80' } ; 
                
        # gmo varis
        # EXOP
        if(-not $rgxExoPsessionstatemoduleDescription){$rgxExoPsessionstatemoduleDescription = '/(exadmin|powershell)$' };
        if(-not $PSSStateOK){$PSSStateOK = 'Opened' };
        if(-not $PSSAvailabilityOK){$PSSAvailabilityOK = 'Available' };
        if(-not $EXOPGmoFilter){$EXOPGmoFilter = 'tmp_*' } ; 
        if(-not $EXOPGmoTestCmdlet){$EXOPGmoTestCmdlet = 'Add-ADPermission' } ; 
        #*------^ END PSS & GMO VARIS ^------

        # exop is dyn remotemod, not installed
        <#Try { Get-Module $modname -ErrorAction Stop | out-null } Catch {
            $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            Import-Module @pltIMod ;
        } ; # IsImported
        [boolean]$IsNoWinRM = [boolean]([version](get-module $modname).version -ge $MinNoWinRMVersion) ; 
        #>

    } ;  # if-E BEGIN    
    PROCESS {
        $oReturn = [ordered]@{
            PSSession = $null ; 
            Valid = $false ; 
        } ; 
        $isEXOPValid = $false ;
        
        if($pssEXOP = Get-PSSession |?{ (
            $_.runspace.connectioninfo.appname -match $rgxEXoPrunspaceConnectionInfoAppName) -AND (
            $_.runspace.connectioninfo.port -eq $EXoPrunspaceConnectionInfoPort) -AND (
            $_.ConfigurationName -eq $EXoPConfigurationName)}){
                    <# rem'd state/avail tests, run separately below: -AND (
                    $_.State -eq $PSSStateOK)  -AND (
                    $_.Availability -eq $PSSAvailabilityOK)
                    #>
            $smsg = "`n`nEXOP PSSessions:`n$(($pssEXOP | fl $pssprops|out-string).trim())" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

            if($pssEXOPGood = $pssEXOP | ?{ ($_.State -eq $PSSStateOK)  -AND (
                    $_.Availability -eq $PSSAvailabilityOK)}){

                # verify the exop cmdlets actually imported as a tmp_ module w specifid prefix & 1st cmdlet
                # tmp_prpfxxlb.ozy
                if ( (get-module -name $EXOPGmoFilter | ForEach-Object { 
                    Get-Command -module $_.name -name $EXOPGmoTestCmdlet -ea 0 
                })) {

                    $smsg = "(EXOPGmo Basic-Auth PSSession module detected)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $isEXOPValid = $true ; 
                } else { $isEXOPValid = $false ; }
            } else{
                # pss but disconnected state
                rxo2 ; 
            } ; 
            
        } else { 
            $smsg = "Unable to detect EXOP PSSession!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
            else{ write-host -ForegroundColor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            #throw $smsg ;
            #Break ; 
            $isEXOPValid = $false ; 
        } ; 

        if($isEXOPValid){
            $oReturn.PSSession = $pssEXOPGood ;
            $oReturn.Valid = $isEXOPValid ; 
        } else { 
            $smsg = "(invalid session `$isEXOPValid:$($isEXOPValid))" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            Disconnect-ex2010 ;
            $oReturn.PSSession = $pssEXOPGood ; 
            $oReturn.Valid = $isEXOPValid ; 
        } ; 

    }  # PROC-E
    END{
        $smsg = "Returning `$oReturn:`n$(($oReturn|out-string).trim())" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        New-Object PSObject -Property $oReturn | write-output ; 
    } ;
}

#*------^ test-EXOPStatus.ps1 ^------


#*------v toggle-ForestView.ps1 v------
Function toggle-ForestView {
<#
.SYNOPSIS
toggle-ForestView.ps1 - Toggle Exchange onprem AD ViewEntireForest setting (permits org-wide object access, wo use of proper explicit -domaincontroller sub.domain.com)
.NOTES
Version     : 1.0.2
Author      : Todd Kadrie
Website     : http://www.toddomation.com
Twitter     : @tostka / http://twitter.com/tostka
CreatedDate : 2020-10-26
FileName    : 
License     : MIT License
Copyright   : (c) 2020 Todd Kadrie
Github      : https://github.com/tostka/verb-XXX
Tags        : Powershell
REVISIONS
* 10:53 AM 4/2/2021 typo fix
* 10:07 AM 10/26.2.30 added CBH
.DESCRIPTION
toggle-ForestView.ps1 - Toggle Exchange onprem AD ViewEntireForest setting (permits org-wide object access, wo use of proper explicit -domaincontroller sub.domain.com)
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output
.EXAMPLE
toggle-ForestView
.LINK
https://github.com/tostka/verb-ex2010
.LINK
#>
[CmdletBinding()]
PARAM() ;
    # toggle forest view
    if (get-command -name set-AdServerSettings){ 
        if (!(get-AdServerSettings).ViewEntireForest ) {
              write-warning "Enabling WholeForest"
              write-host "`a"
              if (get-command -name set-AdServerSettings -ea 0) { set-AdServerSettings -ViewEntireForest $TRUE } ;
        } else {
          write-warning "Disabling WholeForest"
          write-host "`a"
          set-AdServerSettings -ViewEntireForest $FALSE ;
        } ;
    } else {
        THROW "MISSING:set-AdServerSettings`nOPEN an Exchange OnPrem connection FIRST!"
    } ; 
}

#*------^ toggle-ForestView.ps1 ^------


#*======^ END FUNCTIONS ^======

Export-ModuleMember -Function add-MailboxAccessGrant,add-MbxAccessGrant,_cleanup,Connect-Ex2010,Connect-ExchangeServerTDO,_connect-ExOP,get-ADExchangeServerTDO,Connect-Ex2010XO,Connect-ExchangeServerTDO,_connect-ExOP,cx10cmw,cx10tol,cx10tor,disable-ForestView,Disconnect-Ex2010,enable-ForestView,get-ADExchangeServerTDO,get-DAGDatabaseCopyStatus,Get-ExchServerFromExServersGroup,get-ExRootSiteOUs,get-MailboxDatabaseQuotas,Get-MessageTrackingLogTDO,Remove-InvalidVariableNameChars,Connect-ExchangeServerTDO,_connect-ExOP,get-ADExchangeServerTDO,out-Clipboard,remove-SmtpPlusAddress,get-UserMailADSummary,import-EMSLocalModule,Invoke-ExchangeCommand,load-EMSLatest,Load-EMSSnap,new-MailboxGenericTOR,_cleanup,new-MailboxShared,preview-EAPUpdate,Reconnect-Ex2010,Reconnect-Ex2010XO,remove-EMSLocalModule,resolve-ExchangeServerVersionTDO,resolve-RecipientEAP,rx10cmw,rx10tol,rx10tor,test-ExOPPSession,test-EXOPConnection,toggle-ForestView -Alias *




# SIG # Begin signature block
# MIIELgYJKoZIhvcNAQcCoIIEHzCCBBsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUAHFKcNz6lYI1k1yzxX5WDq/1
# mCCgggI4MIICNDCCAaGgAwIBAgIQWsnStFUuSIVNR8uhNSlE6TAJBgUrDgMCHQUA
# MCwxKjAoBgNVBAMTIVBvd2VyU2hlbGwgTG9jYWwgQ2VydGlmaWNhdGUgUm9vdDAe
# Fw0xNDEyMjkxNzA3MzNaFw0zOTEyMzEyMzU5NTlaMBUxEzARBgNVBAMTClRvZGRT
# ZWxmSUkwgZ8wDQYJKoZIhvcNAQEBBQADgY0AMIGJAoGBALqRVt7uNweTkZZ+16QG
# a+NnFYNRPPa8Bnm071ohGe27jNWKPVUbDfd0OY2sqCBQCEFVb5pqcIECRRnlhN5H
# +EEJmm2x9AU0uS7IHxHeUo8fkW4vm49adkat5gAoOZOwbuNntBOAJy9LCyNs4F1I
# KKphP3TyDwe8XqsEVwB2m9FPAgMBAAGjdjB0MBMGA1UdJQQMMAoGCCsGAQUFBwMD
# MF0GA1UdAQRWMFSAEL95r+Rh65kgqZl+tgchMuKhLjAsMSowKAYDVQQDEyFQb3dl
# clNoZWxsIExvY2FsIENlcnRpZmljYXRlIFJvb3SCEGwiXbeZNci7Rxiz/r43gVsw
# CQYFKw4DAh0FAAOBgQB6ECSnXHUs7/bCr6Z556K6IDJNWsccjcV89fHA/zKMX0w0
# 6NefCtxas/QHUA9mS87HRHLzKjFqweA3BnQ5lr5mPDlho8U90Nvtpj58G9I5SPUg
# CspNr5jEHOL5EdJFBIv3zI2jQ8TPbFGC0Cz72+4oYzSxWpftNX41MmEsZkMaADGC
# AWAwggFcAgEBMEAwLDEqMCgGA1UEAxMhUG93ZXJTaGVsbCBMb2NhbCBDZXJ0aWZp
# Y2F0ZSBSb290AhBaydK0VS5IhU1Hy6E1KUTpMAkGBSsOAwIaBQCgeDAYBgorBgEE
# AYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwG
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBSn3atc
# JcGfpqHuhPLEMHm2ExLkDDANBgkqhkiG9w0BAQEFAASBgGPam/EPJXc5QL3ZQmcY
# Y1vmyEqiRq+WCK0uXw/cInKAiLeNmug3wLLIc7yJNqAOkhT+y2TX/YrlBpZIwITz
# 5GO4KAy4UL3ibogKTfr6n6igi644zao/iIbDy6+2SL7Ss+AuR8nn+ySH+xrDbYq6
# Zlp5eoYADwqv8wiYcAcKuc2F
# SIG # End signature block
