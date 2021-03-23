﻿# verb-ex2010.psm1


<#
.SYNOPSIS
VERB-Ex2010 - Exchange 2010 PS Module-related generic functions
.NOTES
Version     : 1.1.48.0
Author      : Todd Kadrie
Website     :	https://www.toddomation.com
Twitter     :	@tostka
CreatedDate : 1/16/2020
FileName    : VERB-Ex2010.psm1
License     : MIT
Copyright   : (c) 1/16/2020 Todd Kadrie
Github      : https://github.com/tostka
REVISIONS
* 11:22 AM 3/13/2020 Get-ExchangeServerInSite added a ping-test, to only return matches that are pingable, added -NoPing param, to permit (faster) untested bypass
* 6:25 PM 1/21/2020 - 1.0.0.1, rebuild, see if I can get a functional module out
* 1/16/2020 - 1.0.0.0
# 7:31 PM 1/15/2020 major revise - subbed out all identifying constants, rplcd regex hardcodes with builds sourced in tor-incl-infrastrings.ps1. Tests functional.
# 11:34 AM 12/30/2019 ran vsc alias-expansion
# 7:51 AM 12/5/2019 Connect-Ex2010:retooled $ExAdmin variant webpool support - now has detect in the server-pick logic, and on failure, it retries to the stock pool.
# 10:19 AM 11/1/2019 trimmed some whitespace
# 10:05 AM 10/31/2019 added sample load/call info
# 12:02 PM 5/6/2019 added cx10,rx10,dx10 aliases
# 11:29 AM 5/6/2019 load-EMSLatest: spliced in from tsksid-incl-ServerApp.ps1, purging ; alias Add-EMSRemote-> Connect-Ex2010 ; toggle-ForestView():moved from tsksid-incl-ServerApp.ps1
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
    .\add-MailboxAccessGrant -ticket 123456 -SiteOverride LYN -TargetID lynctest13 -Owner kadrits -PermsDays 999 -members "lynctest16,lynctest18" -showDebug -whatIf ;
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
        [int]$Ticket,
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
            if($lVers){                 $lVers=($lVers | sort version)[-1];                 try {                     import-module -name $tModName -RequiredVersion $lVers.Version.tostring() -force -DisableNameChecking                 }   catch {                      write-warning "*BROKEN INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;import-module -name $tModDFile -force -DisableNameChecking                 } ;
            } elseif (test-path $tModFile) {                 write-warning "*NO* INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;                 try {import-module -name $tModDFile -force -DisableNameChecking}                 catch {                     write-error "*FAILED* TO LOAD MODULE*:$($tModName) VIA $(tModFile) !" ;                     $tModFile = "$($tModName).ps1" ;                     $sLoad = (join-path -path $LocalInclDir -childpath $tModFile) ;                     if (Test-Path $sLoad) {                         Write-Verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                         . $sLoad ;                         if ($showdebug) { Write-Verbose "Post $sLoad" };                     } else {                         $sLoad = (join-path -path $backInclDir -childpath $tModFile) ;                         if (Test-Path $sLoad) {                             write-verbose  ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                             . $sLoad ;                             if ($showdebug) { write-verbose  "Post $sLoad" };                         } else {                             Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;                             exit;                         } ;                     } ;                 } ;             } ;
            if(!(test-path function:$tModCmdlet)){                 write-warning -verbose:$true  "UNABLE TO VALIDATE PRESENCE OF $tModCmdlet`nfailing through to `$backInclDir .ps1 version" ;                 $sLoad = (join-path -path $backInclDir -childpath "$($tModName).ps1") ;                 if (Test-Path $sLoad) {                     write-verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                     . $sLoad ;                     if ($showdebug) { Write-Verbose -verbose "Post $sLoad" };                     if(!(test-path function:$tModCmdlet)){                         write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO CONFIRM `$tModCmdlet:$($tModCmdlet) FOR $($tModName)" ;                     } else {                         write-verbose  "(confirmed $tModName loaded: $tModCmdlet present)"                     }                 } else {                     Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;                     exit;                 } ;
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
        $reqMods = $reqMods | select -Unique ;

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
            # add for AD SendAs perms grant
            #pulling id, pipeline it in
            $ADMbxGrantSplat = [ordered]@{
                User           = "" ;
                ExtendedRights = "Send As" ;
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


                if ($whatif) {
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

            if ($whatif) {
                $smsg = "-Whatif pass, skipping exec." ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
            } else {
                foreach ($Mbr in $SGMembers) {
                    If ($ExistMbrs -notcontains $Mbr.sAMAccountName) {
                        $smsg = "Test ADD:$($mbr.samaccountname)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                        Add-ADGroupMember @SGUpdtSplat -members $($mbr.samaccountname) -ea stop -whatif ;
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


            # AD SendAs too

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

            # add retry :
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
                    Continue ;
                } # try-E
            } Until ($Exit -eq $Retries) # loop-E

            write-host -foregroundcolor green "$((get-date).ToString("HH:mm:ss")):Exec Permissions Grant Update";
            if ($whatif) {
                # 11:17 AM 6/22/2015 whatif-only pass
                write-verbose "SKIPPING EXEC: Whatif-only pass";
            } else {
                write-host -foregroundcolor red "$((get-date).ToString("HH:mm:ss")):EXEC Add-MailboxPermission...";

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
            write-verbose "$(Get-Date -Format 'HH:mm:ss'):Waiting 5secs to refresh";
            Start-Sleep -s 5 ;

            # secgrp membership seldom comes through clean, add a refresh loop
            do {
                $smsg = "===REVIEW SETTINGS:===`n----Updated Permissions:`n`nChecking Mailbox/AD Permission on $($Tmbx.samaccountname) mailbox `n to accessing user:`n $($oSG.SamAccountName)`n---" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

                $smsg = "`n$((get-mailboxpermission -identity $($TMbx.Identity) -user $(($oSG).Name) -domaincontroller $($InputSplat.domaincontroller) | ?{$_.user -match ".*-(SEC|Data)-Email-.*$"} | format-list user,AccessRights,IsInhertied,Deny |out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
                #$smsg = "`n$((get-mailboxpermission -identity $($TMbx.Identity) -user $(($oSG).Name) -domaincontroller $($InputSplat.domaincontroller) | ?{$_.user -match ".*-(SEC|Data)-Email-.*$"} | format-list user,AccessRights,IsInhertied,Deny|out-string).trim())" ;
                #if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

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

#*------v Connect-Ex2010.ps1 v------
Function Connect-Ex2010 {
  <#
    .SYNOPSIS
    Connect-Ex2010 - Setup Remote Exch2010 Mgmt Shell connection
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
    [CmdletBinding()]
    [Alias('Add-EMSRemote','cx10')]
    Param(
        [Parameter(Position = 0, HelpMessage = "Exch server to Remote to")][string]$ExchangeServer,
        [Parameter(HelpMessage = 'Use exadmin IIS WebPool for remote EMS[-ExAdmin]')]$ExAdmin,
        [Parameter(HelpMessage = 'Credential object')][System.Management.Automation.PSCredential]$Credential = $credOpTORSID
    )  ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ; 
        $sWebPoolVariant = "exadmin" ;
        $CommandPrefix = $null ;
        # use credential domain to determine target org
        $rgxLegacyLogon = '\w*\\\w*' ; 
        $TenOrg = get-TenantTag -Credential $Credential ;
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
            if( $ExchangeServer = (Get-ExchangeServerInSite | ? { ($_.roles -eq 36) } | Get-Random ).FQDN){}
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
        #$EMSsplat = @{ConnectionURI = "http://$ExchangeServer/powershell"; ConfigurationName = 'Microsoft.Exchange' ; name = 'Exchange2010' } ;
        $EMSsplat = @{ConnectionURI = "http://$ExchangeServer/powershell"; ConfigurationName = 'Microsoft.Exchange' ; name = "Exchange$($ExVers)" } ;
        if($env:USERDOMAIN -ne (Get-Variable  -name "$($TenOrg)Meta").value.legacyDomain){
            # if not in the $TenOrg legacy domain - running cross-org -  add auth:Kerberos
            <#suppresses: The WinRM client cannot process the request. It cannot determine the content type of the HTTP response f rom the destination computer. The content type is absent or invalid
            #>
            $EMSsplat.add('Authentication','Kerberos') ;
        } ; 
        if ($ExAdmin) {
          # use variant IIS Webpool
          $EMSsplat.ConnectionURI = $EMSsplat.ConnectionURI.replace("/powershell", "/$($sWebPoolVariant)") ;
        }
        if ($Credential) {
             $EMSsplat.Add("Credential", $Credential) 
             write-verbose "(using cred:$($credential.username))" ; 
        } ;
        
        # -Authentication Basic only if specif needed: for Ex configured to connect via IP vs hostname)
        # try catch against and retry into stock if fails
        $error.clear() ;
        TRY {
          $Global:E10Sess = New-PSSession @EMSSplat -ea STOP  ;
        } CATCH {
          $ErrTrapd = $_ ; 
          write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
          if ($ExAdmin) {
            # switch to stock pool and retry
            $EMSsplat.ConnectionURI = $EMSsplat.ConnectionURI.replace("/$($sWebPoolVariant)", "/powershell") ;
            write-warning -verbose:$true "$((get-date).ToString('HH:mm:ss')):FAILED TARGETING EXADMIN POOL`nRETRY W STOCK POOL: New-PSSession w`n$(($EMSSplat|out-string).trim())" ;
            $Global:E10Sess = New-PSSession @EMSSplat -ea STOP  ;
          }
          else {
            STOP ;
          } ;
        } ;

        write-verbose "$((get-date).ToString('HH:mm:ss')):Importing Exchange 2010 Module" ;

        if ($CommandPrefix) {
          write-host -foregroundcolor white "$((get-date).ToString("HH:mm:ss")):Note: Prefixing this Mod's Cmdlets as [verb]-$($CommandPrefix)[noun]" ;
          $Global:E10Mod = Import-Module (Import-PSSession $Global:E10Sess -DisableNameChecking -Prefix $CommandPrefix -AllowClobber) -Global -Prefix $CommandPrefix -PassThru -DisableNameChecking   ;
        } else {
          $Global:E10Mod = Import-Module (Import-PSSession $Global:E10Sess -DisableNameChecking -AllowClobber) -Global -PassThru -DisableNameChecking   ;
        } ;
        if($ExVwForest){
            write-host "Setting EMS Session: Set-AdServerSettings -ViewEntireForest `$True" ; 
            Set-AdServerSettings -ViewEntireForest $True ; 
        } ; 
        # 7:54 AM 11/1/2017 add titlebar tag
        Add-PSTitleBar 'EMS' ;
        # tag E10IsDehydrated 
        $Global:E10IsDehydrated = $true ;
        write-host -foregroundcolor darkgray "`n$(($Global:E10Sess | select ComputerName,Availability,State,ConfigurationName | format-table -auto |out-string).trim())" ;
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
        # if we're using ems-style BasicAuth, clear incompatible existing Rems PSS's
        # ComputerName      : curlyhoward.cmw.internal ;  ComputerType      : RemoteMachine ;  State             : Opened ;  ConfigurationName : Microsoft.Exchange ;  Availability      : Available ;  Name              : Session1 ;   ;
        $rgxRemsPSSName = "^(Session\d|Exchange\d{4})$" ;
        $Rems2Good = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName) -AND ($_.Availability -eq 'Available') } ;
        # Computername wrong fqdn suffix
        $Rems2WrongOrg = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND (-not($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName)) -AND ($_.Availability -eq 'Available') } ;
        $Rems2Broken = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -like "*Broken*") } ;
        $Rems2Closed = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -like "*Closed*") } ;

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

            $EMSsplat = @{
                ConnectionURI = "http://$((Get-Variable  -name "$($TenOrg)Meta").value.Ex10ServerXO)/powershell";
                ConfigurationName = 'Microsoft.Exchange' ;
                name = 'Exchange2010' ;
            } ;
            if ((Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant) {
              # use variant IIS Webpool
              $EMSsplat.ConnectionURI = $EMSsplat.ConnectionURI.replace("/powershell", "/$((Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant)") ;
            }
            $EMSsplat.Add("Credential", $Credential); # just use the passed $Credential vari
            $cMsg = "Connecting to OP Ex20XX ($($credDom))";
            Write-Host $cMsg ;
            write-verbose "`n$((get-date).ToString('HH:mm:ss')):New-PSSession w`n$(($EMSsplat|out-string).trim())" ;

            $error.clear() ;
            TRY { $global:E10Sess = New-PSSession @EMSSplat -ea STOP
            } CATCH {
                $ErrTrapd = $_ ;
                write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ((Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant) {
                  # switch to stock pool and retry
                  $EMSsplat.ConnectionURI = $EMSsplat.ConnectionURI.replace("/$((Get-Variable  -name "$($TenOrg)Meta").value.Ex10WebPoolVariant)", "/powershell") ;
                  write-warning -verbose:$true "$((get-date).ToString('HH:mm:ss')):FAILED TARGETING VARIANT POOL`nRETRY W STOCK POOL: New-PSSession w`n$(($EMSSplat|out-string).trim())" ;
                  $global:E10Sess = New-PSSession @EMSSplat -ea STOP  ;
                } else {
                    STOP ;
                } ;
            } ; # try-E

            if(!$global:E10Sess){
                write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO RETURN PSSESSION!`nAUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                Break ;
            } ;

            $pltIMod=@{Global = $true ;PassThru = $true;DisableNameChecking = $true ; } ;
            if ($CommandPrefix) {
                write-host -foregroundcolor white "$((get-date).ToString("HH:mm:ss")):Note: Prefixing this Mod's Cmdlets as [verb]-$($CommandPrefix)[noun]" ;
                $pltIMod.add('Prefix',$CommandPrefix) ;
            } ;
            $pltPSS = [ordered]@{
                Session             = $global:E10Sess ;
                DisableNameChecking = $true  ;
                AllowClobber        = $true ;
                ErrorAction         = 'Stop' ;
            } ;
             #-Global -Prefix $CommandPrefix -PassThru -DisableNameChecking   ;
            # $Global:E10Mod = Import-Module (Import-PSSession $Global:E10Sess -DisableNameChecking -Prefix $CommandPrefix -AllowClobber) -Global -Prefix $CommandPrefix -PassThru -DisableNameChecking   ;

            write-verbose "`n$((get-date).ToString('HH:mm:ss')):Import-PSSession w`n$(($pltPSS|out-string).trim())`nImport-Module w`n$(($pltIMod|out-string).trim())" ;

            # Verbose:Continue is VERY noisey for module loads. Bracketed suppress:
            if($VerbosePreference = "Continue"){
                $VerbosePrefPrior = $VerbosePreference ;
                $VerbosePreference = "SilentlyContinue" ;
                $verbose = ($VerbosePreference -eq "Continue") ;
            } ;
            Try {
                $Global:E10Mod = Import-Module (Import-PSSession @pltPSS) @pltIMod  ;
                #$Global:EOLModule = Import-Module (Import-PSSession @pltPSS) -Global -Prefix $CommandPrefix -PassThru -DisableNameChecking   ;
            } catch {
                Write-Warning -Message "Tried but failed to import the EXO PS module.`n`nError message:" ;
                throw $_ ;
            } ;
            # reenable VerbosePreference:Continue, if set, during mod loads
            if($VerbosePrefPrior -eq "Continue"){
                $VerbosePreference = $VerbosePrefPrior ;
                $verbose = ($VerbosePreference -eq "Continue") ;
            } ;
            Add-PSTitleBar $sTitleBarTag ;

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
    .EXAMPLE
    cx10tol
    #>
    [CmdletBinding()] 
    Param()
    $Verbose = ($VerbosePreference -eq 'Continue') ;
    $pltGHOpCred=@{TenOrg="TOL" ;userrole=@('ESVC','LSVC','SID') ;verbose=$($verbose)} ;
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

#*------^ cx10tol.ps1 ^------

#*------v cx10tor.ps1 v------
function cx10tor {
    <#
    .SYNOPSIS
    cx10tor - Connect-EX2010 to specified on-prem Exchange
    .DESCRIPTION
    Connect-EX2010 - Connect-EX2010 to specified on-prem Exchange
    .EXAMPLE
    cx10tor
    #>
    Connect-EX2010 -cred $credTorSID -Verbose:($VerbosePreference -eq 'Continue') ; 
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
              write-warning "Disabling WholeForest"
              write-host "`a"
              if (get-command -name set-AdServerSettings -ea 0) { set-AdServerSettings -ViewEntireForest $False } ;
        } ;
    } else {
        THROW "MISSING:set-AdServerSettings`nOPEN an Exchange OnPrem connection FIRST!"
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
    if($Global:E10Mod){$Global:E10Mod | Remove-Module -Force } ; 
    if($Global:E10Sess){$Global:E10Sess | Remove-PSSession } ;
    # 7:56 AM 11/1/2017 remove titlebar tag
    Remove-PSTitlebar 'EMS' ;
    # kill any other sessions using distinctive name; add verbose, to ensure they're echo'd that they were missed
    Get-PSSession | ? { $_.name -eq 'Exchange2010' } | Remove-PSSession -verbose ;
    # kill any broken PSS, self regen's even for L13 leave the original borked and create a new 'Session for implicit remoting module at C:\Users\', toast them, they don't reopen. Same for Ex2010 REMS, identical new PSS, indistinguishable from the L13 regen, except the random tmp_xxxx.psm1 module name. Toast them, it's just a growing stack of broken's
    Disconnect-PssBroken ;
    [console]::ResetColor()  # reset console colorscheme
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
              write-warning "Enabling WholeForest"
              write-host "`a"
              if (get-command -name set-AdServerSettings -ea 0) { set-AdServerSettings -ViewEntireForest $TRUE } ;
        } ;
    } else {
        THROW "MISSING:set-AdServerSettings`nOPEN an Exchange OnPrem connection FIRST!"
    } ; 
}

#*------^ enable-ForestView.ps1 ^------

#*------v Get-ExchangeServerInSite.ps1 v------
Function Get-ExchangeServerInSite {
    <#
    .SYNOPSIS
    Get-ExchangeServerInSite - Returns the name of an Exchange server in the local AD site.
    .NOTES
    Author: Mike Pfeiffer
    Website:	http://mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
    REVISIONS   :
    * 12:08 PM 5/15/2020 fixed vpn issue: Try/Catch'd around recently failing $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName qry
    * 11:22 AM 3/13/2020 Get-ExchangeServerInSite added a ping-test, to only return matches that are pingable, added -NoPing param, to permit (faster) untested bypass
    * 6:59 PM 1/15/2020 cleanup
    # 10:03 AM 11/16/2018 Get-ExchangeServerInSite:can't do AD-related functions when not AD authentictaed (home, pre-vpn connect). Added if/then test on status and abort balance when false.
    # 12:10 PM 8/1/2017 updated example code at bottom, to accommodate lyn & adl|spb
    # 11:28 AM 3/31/2016 validated that latest round of updates are still functional
    #1:58 PM 9/3/2015 - added pshelp and some docs
    #April 12, 2010 - web version
    .PARAMETER  Site
    Optional: Ex Servers from which Site (defaults to AD lookup against local computer's Site)
    .DESCRIPTION
    Get-ExchangeServerInSite - Returns the name of an Exchange server in the local AD site.
    Uses an ADSI DirectorySearcher to search the current Active Directory site for Exchange 2010 servers.
    Returned object includes the post-filterable Role property which reflects the following
    installed-roles on the discovered server
	    Mailbox Role - 2
        Client Access Role - 4
        Unified Messaging Role - 16
        Hub Transport Role - 32
        Edge Transport Role - 64
        Add the above up to combine roles:
        HubCAS = 32 + 4 = 36
        HubCASMbx = 32+4+2 = 38
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns the name of an Exchange server in the local AD site.
    .EXAMPLE
    .\Get-ExchangeServerInSite
    .EXAMPLE
    get-exchangeserverinsite |?{$_.roles -match "(4|32|36)"}
    Return Hub,CAS,or Hub+CAS servers
    .EXAMPLE
    If(!($ExchangeServer)){$ExchangeServer=(Get-ExchangeServerInSite |?{($_.roles -eq 36) -AND ($_.FQDN -match "SITECODE.*")} | Get-Random ).FQDN }
    Return a random HubCas Role server with a name beginning LYN
    .EXAMPLE
    $regex = '(' + [regex]($ADSiteCodeUK,$ADSiteCodeAU -join "|") + ')'
    switch -regex ($($env:computername).substring(0,3)){
       "$($ADSiteCodeUS)" {$tExRole=36 } ;
       "$($regex)" {$tExRole= 38 }  default { write-error "$((get-date).ToString('HH:mm:ss')):UNRECOGNIZED SERVER PREFIX!."; } ;
    } ;
    $exhubcas = (Get-ExchangeServerInSite |?{($_.roles -eq $tExRole) -AND ($_.FQDN -match "$($env:computername.substring(0,3)).*")} | Get-Random ).FQDN ;
    Return a random HubCas Role server with a name matching the $ENV:COMPUTERNAME
    .LINK
    http://mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
    #>
    [CmdletBinding()]
    PARAM(
        [Parameter(HelpMessage="Switch to suppress default 'pingable' test (e.g. returns all matches, no testing)[-NoPing]")]
        [switch] $NoPing
    ) ;
    $Verbose = ($VerbosePreference -eq 'Continue') ; 
    # 9:53 AM 11/16/2018 from vpn/home, $ADSite doesn't populate prior to domain logon (via vpn)
    # 9:41 AM 5/15/2020 issue: vpn/home, $siteDN suddenly doesn't populate, no longer dyn locs an Ex box, implemented try/catch workaround
    if ($ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]) {
        TRY {$siteDN = $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName}
        CATCH {
            $siteDN =$Ex10siteDN # [infra] returns DN to : cn=[SITENAME],cn=sites,cn=configuration,dc=ad,dc=[DOMAIN],dc=com
            write-warning "$((get-date).ToString('HH:mm:ss')):`$siteDN lookup FAILED, deferring to hardcoded `$Ex10siteDN string in infra file!" ;
        } ; 
        TRY {$configNC = ([ADSI]"LDAP://RootDse").configurationNamingContext}
        CATCH {
            $configNC =$Ex10configNC #  [infra] returns: "CN=Configuration,DC=ad,DC=[DOMAIN],DC=com"
            write-warning "$((get-date).ToString('HH:mm:ss')):`$configNC lookup FAILED, deferring to hardcoded `$Ex10configNC string in infra file!" ;
        } ; 
        if($siteDN -AND $configNC){
            $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC") ;
            $objectClass = "objectClass=msExchExchangeServer" ;
            $version = "versionNumber>=1937801568" ;
            $site = "msExchServerSite=$siteDN" ;
            $search.Filter = "(&($objectClass)($version)($site))" ;
            $search.PageSize = 1000 ;
            [void] $search.PropertiesToLoad.Add("name") ;
            [void] $search.PropertiesToLoad.Add("msexchcurrentserverroles") ;
            [void] $search.PropertiesToLoad.Add("networkaddress") ;
            $search.FindAll() | % {
                $matched = New-Object PSObject -Property @{
                    Name  = $_.Properties.name[0] ;
                    FQDN  = $_.Properties.networkaddress |
                        % { if ($_ -match "ncacn_ip_tcp") { $_.split(":")[1] } } ;
                    Roles = $_.Properties.msexchcurrentserverroles[0] ;
                } ;
                if($NoPing){
                    $matched | write-output ; 
                } else { 
                    $matched | %{If(test-connection $_.FQDN -count 1 -ea 0) {$_} else {} } | 
                        write-output ; 
                } ; 
            } ;
        }else {
            write-warning  "$((get-date).ToString('HH:mm:ss')):MISSING `$siteDN:($($siteDN)) `nOR `$configNC:($($configNC)) values`nABORTING!" ;
            $false | write-output ;
        } ;
    }else {
        write-warning -verbose:$true  "$((get-date).ToString('HH:mm:ss')):`$ADSite blank, not authenticated to a domain! ABORTING!" ;
        $false | write-output ;
    } ;
}

#*------^ Get-ExchangeServerInSite.ps1 ^------

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
        ? { $_.distinguishedname -match $ServerRegex }).name | 
            get-random | write-output ;
}

#*------^ Get-ExchServerFromExServersGroup.ps1 ^------

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
            Import-Module $PlugName -ErrorAction Stop ; write-output $TRUE;
          }
          "Snapin" {
            Add-PSSnapin $PlugName -ErrorAction Stop ; write-output $TRUE
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
            if($lVers){                 $lVers=($lVers | sort version)[-1];                 try {                     import-module -name $tModName -RequiredVersion $lVers.Version.tostring() -force -DisableNameChecking                 }   catch {                      write-warning "*BROKEN INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;import-module -name $tModDFile -force -DisableNameChecking                 } ;
            } elseif (test-path $tModFile) {                 write-warning "*NO* INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;                 try {import-module -name $tModDFile -force -DisableNameChecking}                 catch {                     write-error "*FAILED* TO LOAD MODULE*:$($tModName) VIA $(tModFile) !" ;                     $tModFile = "$($tModName).ps1" ;                     $sLoad = (join-path -path $LocalInclDir -childpath $tModFile) ;                     if (Test-Path $sLoad) {                         write-verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                         . $sLoad ;                         if ($showdebug) { write-verbose "Post $sLoad" };                     } else {                         $sLoad = (join-path -path $backInclDir -childpath $tModFile) ;                         if (Test-Path $sLoad) {                             write-verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;                             . $sLoad ;                             if ($showdebug) { write-verbose "Post $sLoad" };                         } else {                             Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;                             exit;                         } ;                     } ;                 } ;             } ;
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

#*------v Reconnect-Ex2010.ps1 v------
Function Reconnect-Ex2010 {
  <#
    .SYNOPSIS
    Reconnect-Ex2010 - Reconnect Remote Exch2010 Mgmt Shell connection
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
    [Alias('rx10')]
    Param(
        [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")][System.Management.Automation.PSCredential]
        $Credential = $global:credOpTORSID
    )
    if (!$E10Sess) {
      if (!$Credential) {
        Connect-Ex2010
      }
      else {
        Connect-Ex2010 -Credential:$($Credential) ;
      } ;
    }
  elseif ($E10Sess.state -ne 'Opened' -OR $E10Sess.Availability -ne 'Available' ) {
    Disconnect-Ex2010 ; Start-Sleep -S 3;
    if (!$Credential) {
      Connect-Ex2010
    }
    else {
      Connect-Ex2010 -Credential:$($Credential) ;
    } ;
  } ;
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
        $Rems2WrongOrg = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND (-not($_.ComputerName -match (Get-Variable  -name "$($TenOrg)Meta").value.OP_rgxEMSComputerName)) -AND ($_.Availability -eq 'Available') } ;
        $Rems2Broken = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -like "*Broken*") } ;
        $Rems2Closed = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -like "*Closed*") } ;

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
* 10:07 AM 10/26/2020 added CBH
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
          write-warning "Disableing WholeForest"
          write-host "`a"
          set-AdServerSettings -ViewEntireForest $FALSE ;
        } ;
    } else {
        THROW "MISSING:set-AdServerSettings`nOPEN an Exchange OnPrem connection FIRST!"
    } ; 
}

#*------^ toggle-ForestView.ps1 ^------

#*======^ END FUNCTIONS ^======

Export-ModuleMember -Function add-MailboxAccessGrant,Connect-Ex2010,Connect-Ex2010XO,cx10cmw,cx10tol,cx10tor,disable-ForestView,Disconnect-Ex2010,enable-ForestView,Get-ExchangeServerInSite,Get-ExchServerFromExServersGroup,Invoke-ExchangeCommand,load-EMSLatest,Load-EMSSnap,new-MailboxShared,Reconnect-Ex2010,Reconnect-Ex2010XO,rx10cmw,rx10tol,rx10tor,toggle-ForestView -Alias *


# SIG # Begin signature block
# MIIELgYJKoZIhvcNAQcCoIIEHzCCBBsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUQkwr4FO1eHPgqkcEB4e2YNVf
# ygmgggI4MIICNDCCAaGgAwIBAgIQWsnStFUuSIVNR8uhNSlE6TAJBgUrDgMCHQUA
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
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBSK3fzh
# koEL8ksdetLrJn7PZnZILTANBgkqhkiG9w0BAQEFAASBgFA4TbhhisbqiUADLTPN
# qd8XRU2x0PnzHNz/kiT+2uEEWv6chDu4btN3jxSfcDtXcOjuKnui9WoPtxsIx3Jv
# Ya+jM85duo0heEH6rNkkfbobSuq6/W019moewTiwvlH3ibiDXvqOU3NKLlZSE4dx
# wj8dlfSZvLQEjmjTlctlLRRc
# SIG # End signature block
