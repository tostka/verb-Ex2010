# add-MailboxAccessGrant.ps1

#region ADD_MAILBOXACCESSGRANT ; #*------v add-MailboxAccessGrant v------
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
    * 3:06 PM 4/15/2026 shifted to Name/Cn & samaccountname as strict stripped char versions
    * 4:15 PM 4/14/2026 captured outputs of add- cmds (blowing pipeline); 
    * 5:15 PM 4/3/2026 substantial recode to support non-standard OU names (LF) in migrations OUs. 
    * 4:18 PM 3/30/2026 fixed borked $FallBackBaseUserOU typo; subbed in full begin block up to splat defs from new-mailboxshared()
    # 5:12 PM 10/13/2021 fixed long standing random add-adgroupmember bug (failed to see target sg/dg), by swapping in ADGM ex cmd;  pulled [int] from $ticket , to permit non-numeric & multi-tix
    # 11:36 AM 9/16/2021 string
    # 10:27 AM 9/14/2021 beefed up echos w 7pswhsplat's pre
    # 3:21 PM 8/17/2021 recoded grabbing outputs, on object creations in EMS, tho' AD object creations generate no proper output. functoinal on current ticket.
    # 4:48 PM 8/16/2021 still wrestling grant fails, switched the *permission -user targets to the dg object.primarysmtpaddr (was adg.samaccountname), if the adg.sama didn't match the alias, that woulda caused failures. Seemd to work better in debugging
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
    BEGIN {
        $dbgDate = '4/22/2026'; # debugging ipmo force loads variants not in modules
        $Verbose = [boolean]($VerbosePreference -eq 'Continue') ; 
        # Name/Ldap CN banned chars
        $rgxBanCNName = '[\,\+\"\\\<\>\;\=\/]' ; 
        # samAccountName banned chars
        $rgxBanSamA = '[\"\/\\\[\]\:\;\|\=\,\+\*\?\<\>\s]' ; 

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

        $Retries = 4 ; # number of re-attempts
        $RetrySleep = 5 ; # seconds to wait between retries
        # $rgxCU5 = [infra file]
        # OU that's used when can't find any baseuser for the owner's OU, default to a random shared from ($ADSiteCodeUS) (avoid crapping out):
        $FallBackBaseUserOU = "$($DomTORfqdn)/$($ADSiteCodeUS)/Generic Email Accounts" ;
        # 3:46 PM 4/3/2026 add Migrations OU variant support
        $rgxOUMigrations = ',OU=_MIGRATIONS,DC=global,DC=ad,DC=toro,DC=com$' ;
        $rgxMigationsSite = ',OU=(\w+),OU=_MIGRATIONS,DC=global,DC=ad,DC=toro,DC=com$' ;
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
    PROCESS {

        # 1:19 PM 2/13/2019 email trigger vari, it will be semi-delimd list of mail-triggering events
        # TenOrg or other looped-specific PassStatus (auto supported by 7pswlt)
        if(-not(get-Variable -Name PassStatus_$($tenorg) -scope Script -ea 0)){
            New-Variable -Name PassStatus_$($tenorg) -scope Script -Value $null ;
        } ; 
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
            if($ticket){$pltSL.Tag = $ticket} ;
            #$pltSL.Tag = $env:COMPUTERNAME ; 
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
        #$InputSplat.Site = ($Tmbx.identity.tostring().split('/')[1]) ;
        #$rgxOUMigrations = ',OU=_MIGRATIONS,DC=global,DC=ad,DC=toro,DC=com$' ;
        #$rgxMigationsSite = ',OU=(\w+),OU=_MIGRATIONS,DC=global,DC=ad,DC=toro,DC=com$' ;
        if($Tmbx.DistinguishedName -match $rgxOUMigrations){
            $InputSplat.Site = [regex]::Match($Tmbx.DistinguishedName,$rgxMigationsSite).groups[1].value ;
            if($InputSplat.Site){write-host -foregroundcolor green  "Resolved _MIGRATIONS tree OSiteCode:$($InputSplat.Site)" }else{ throw "Unable to resolve $($pltNmbx.Owner.DistinguishedName) into a SiteCode" ;BREAK }
        }else{
            $InputSplat.Site = ($Tmbx.identity.tostring().split('/')[1]) ;
        } ;

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
            #$SiteCode = $InputSplat.OwnerMbx.identity.tostring().split("/")[1]  ;
            if($InputSplat.OwnerMbx.DistinguishedName -match $rgxOUMigrations){
                $SiteCode = [regex]::Match($InputSplat.OwnerMbx.DistinguishedName,$rgxMigationsSite).groups[1].value ;
                if($SiteCode ){write-host -foregroundcolor green  "Resolved _MIGRATIONS tree OSiteCode:$($SiteCode )" }else{ throw "Unable to resolve $($pltNmbx.Owner.DistinguishedName) into a SiteCode" ;BREAK }
            }else{
                $SiteCode = ($InputSplat.OwnerMbx.identity.tostring().split('/')[1]) ;
            } ;
        } ;
        <# 4:16 PM 4/3/2026 PRIOR
        if ($env:USERDOMAIN -eq $TORMeta['legacyDomain']) {
            $FindOU = "^OU=Email\sAccess,OU=SEC\sGroups,OU=Managed\sGroups,";
        } ELSEif ($env:USERDOMAIN -eq $TOLMeta['legacyDomain']) {
            # CN=Lab-SEC-Email-Thomas Jefferson,OU=Email Access,OU=SEC Groups,OU=Managed Groups,OU=LYN,DC=SUBDOM,DC=DOMAIN,DC=DOMAIN,DC=com
            $FindOU = "^OU=Email\sAccess,OU=SEC\sGroups,OU=Managed\sGroups,"; ;
        } else {
            throw "UNRECOGNIZED USERDOMAIN:$($env:USERDOMAIN)" ;
        } ;
        #>
        # 4:16 PM 4/3/2026 REVISE TO TRY TO ACCOMODATE F-ERY IN MIGRATIONS TREE (THX LF)
        if ($env:USERDOMAIN -eq $TORMeta['legacyDomain']) {
            if($Tmbx.DistinguishedName -match $rgxOUMigrations){
                $FindOU = "^OU=Email\sAccess," ; 
                # MIGHT HAVE TO TEST A SERIES, IF LF REALLY MANGLED THE PATHS ACROSS DIVISIONS!
            }ELSE{
                $FindOU = "^OU=Email\sAccess,OU=SEC\sGroups,OU=Managed\sGroups,";
            }
        } ELSEif ($env:USERDOMAIN -eq $TOLMeta['legacyDomain']) {
            # CN=Lab-SEC-Email-Thomas Jefferson,OU=Email Access,OU=SEC Groups,OU=Managed Groups,OU=LYN,DC=SUBDOM,DC=DOMAIN,DC=DOMAIN,DC=com
            $FindOU = "^OU=Email\sAccess,OU=SEC\sGroups,OU=Managed\sGroups,"; ;
        } else {
            throw "UNRECOGNIZED USERDOMAIN:$($env:USERDOMAIN)" ;
        } ;
        $SGSplat.DisplayName = "$($SiteCode)-SEC-Email-$($Tmbx.DisplayName)-G";

        <#
        TRY {
            if($Tmbx.DistinguishedName -match $rgxOUMigrations){
                #$OU = (Get-ADObject -filter { ObjectClass -eq 'organizationalunit' } -server $($DomainController) | Where-Object { $_.distinguishedname -match "^$($FindOU).*OU=$($SiteCode),.*,OU=_MIGRATIONS,DC=global,DC=ad,DC=toro,DC=com$" } | Select-Object distinguishedname).distinguishedname.tostring() ;
                #OU=Email\sAccess,.*OU=_TTC_Sync_CMW_NoSync,OU=DIT,OU=_MIGRATIONS,
                #$OU = (Get-ADObject -filter { ObjectClass -eq 'organizationalunit' } -server $($DomainController) | Where-Object { $_.distinguishedname -match "^$($FindOU).*OU=_TTC_Sync_CMW_NoSync,OU=$($SiteCode),OU=_MIGRATIONS," } | Select-Object distinguishedname).distinguishedname.tostring() ;
                # leverage verb-adms\get-SiteMbxOU
            }else{
                #$OU = (Get-ADObject -filter { ObjectClass -eq 'organizationalunit' } -server $($DomainController) | Where-Object { $_.distinguishedname -match "^$($FindOU).*OU=$($SiteCode),.*,DC=ad,DC=toro((lab)*),DC=com$" } | Select-Object distinguishedname).distinguishedname.tostring() ;
                # leverage verb-adms\get-SiteMbxOU

            }
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
        #>
        # port in code from new-mailboxshared
        #-=-=-=-=-=-=-=-=
        $pltGSmbx = [ordered]@{
            Sitecode = $SiteCode ;
            Type = 'PermissionGroup' ; 
        }
        if ($Tmbx.DistinguishedName -match $rgxOUMigrations) {
            $pltGSmbx.add('modelDistinguishedName',$Tmbx.DistinguishedName) ; 
        }else{
            $pltGSmbx.add('modelDistinguishedName',$Tmbx.DistinguishedName) ; 
        }
        <#
        If($InputSplat.NonGeneric) {
            if($pltGSmbx.keys -contains 'generic'){$pltGSmbx.remove('Generic')}
        } elseIf($Room -OR $Equipement) {
            $pltGSmbx.add('Resource',$true) ;
        } else {
            $pltGSmbx.add('Generic',$true ) ;
        } ;
        #>
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
        if ( $OU  = (Get-SiteMbxOU @pltGSmbx)   ) {
        } else { Cleanup ; BREAK ;}
        #-=-=-=-=-=-=-=-=

        $smsg = "$SiteCode" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug
        $smsg = "$OU" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn|Debug

        $SGSplat.Path = $OU ;
        $smsg = "Checking specified SecGrp Members..." ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
        # 3:50 PM 10/13/2021 flip ea stop to continue, we want it to get through, even if it throws error, and continue will complain
        $SGMembers = ($InputSplat.members.split(",") | ForEach-Object {
             get-recipient $_ -ea continue | select -expand primarysmtpaddress  | select -unique
            }
        )
        $smsg = "Checking for existing $($SGSplat.DisplayName)..." ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

        if ($bDebug) {
            $smsg = "`$SGSrchName:$($SGSrchName)`n`$SGSplat.DisplayName: $($SGSplat.DisplayName)"; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ;
        } ;

        # 4:38 PM 4/14/2026 need samacctname generated 1st
        # THIS IS ALPHANUM STRIP - NOT LEGIBLE
        #$SGSplat.SamAccountName =  -join (($SGSplat.DisplayName -replace '[^a-zA-Z0-9]', '').ToCharArray()[0..19]) ;
        # SHIFT TO STRICT BAN CHAR STRIP
        #$rgxBanSamA = '[\"\/\\\[\]\:\;\|\=\,\+\*\?\<\>\s]' ; 
        $SGSplat.SamAccountName =  -join (($SGSplat.DisplayName -replace $rgxBanSamA, '').ToCharArray()[0..19]) ;
        try {$conflict = $null ;$conflict = get-adgroup -id $SGSplat.SamAccountName -server $InputSplat.DomainController ;} CATCH {} ;
        $incr = 1 ;
        Do{
            if($conflict){
                $incr++ ;
                $SGSplat.SamAccountName = "$( -join (($SGSplat.DisplayName -replace '[^a-zA-Z0-9]', '').ToCharArray()[0..18]) )$($incr)" ;
            }
            try {$conflict = $null ;$conflict = get-adgroup -id $SGSplat.SamAccountName -server $InputSplat.DomainController ;} CATCH {} ;
        }while($conflict) ; 

        $SGSrchName = $($SGSplat.samaccountname);
        # 4:31 PM 4/14/2026 samacct isn't defined yet, search on DisplayName, and get-adgroup needs try/catch to suppress hit fails
        # and can't search on dname, it's not in default properties, have to preconstruct accurate samacct, and search it
        TRY{
            $oSG = Get-ADGroup -Filter { SamAccountName -eq $SGSrchName } -server $($InputSplat.DomainController) -ErrorAction stop;
            #$oSG = Get-ADGroup -Filter { Displayname -eq $SGSplat.DisplayName } -server $($InputSplat.DomainController) -ErrorAction stop;            
        }CATCH{}

        if ($oSG) {
            if ($bDebug) {
                $smsg = "`$oSG:$($oSG.SamAccountname)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; #Error|Warn
                $smsg = "`$oSG.DN:$($oSG.DistinguishedName)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Debug } ; #Error|Warn
            } ;

            # 4:16 PM 8/16/2021 nope, it's not dg-enabled yet, it's a secgrp, can't pull it.
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
            #$SGSplat.Name = $($SGSplat.DisplayName); 
            # LDAP cn/nAME name restrictcions 64char max, ban: , + " \ < > ; = / (must be escaped, strip NON-NUMS SIMPLER)
            #$SGSplat.Name = -join (($SGSplat.DisplayName -replace '[,\+"\\<>;=/]', '').ToCharArray()[0..63]) ;
            # ILLEG ALPHANUM STRIP
            #$SGSplat.Name =  -join (( $SGSplat.DisplayName -replace '[^a-zA-Z0-9\s]','').ToCharArray()[0..63]) ;
            # do legible banned char strip instead
            #$rgxBanCNName = '[\,\+\"\\\<\>\;\=\/]' ; 
            $SGSplat.Name = -join (($SGSplat.DisplayName -replace $rgxBanCNName, '').ToCharArray()[0..63]) ; 
            try {$conflict = $null ;$conflict = get-adgroup -id $SGSplat.Name -server $InputSplat.DomainController ;} CATCH {} ;
            $incr = 1 ;
            Do{
                if($conflict){
                    $incr++ ;
                    $SGSplat.Name = "$( -join (( $SGSplat.DisplayName -replace '[^a-zA-Z0-9\s]','').ToCharArray()[0..62]) )$($incr)" ;
                }
                try {$conflict = $null ;$conflict = get-adgroup -id $SGSplat.Name -server $InputSplat.DomainController ;} CATCH {} ;
            }while($conflict) ; 

            # 8:45 AM 4/13/2026 with Migrations oddities, we need to strip the dname of non-alphanums for the samacctname & identity
            # sAMAccountName: 20chars max, ban:/ \ [ ] : ; | = , + * ? < >
            #$SGSplat.SamAccountName = $($SGSplat.DisplayName -replace '[\W]',''); # strip all non-alphnanum from the dname, doesn't cover 2-20 length req
            #$SGSplat.SamAccountName = [regex]::match($SGSplat.DisplayName,'[-A-Za-z0-9]{2,20}').value ; # this accomds the perm of dash as well
            <# 4:37 PM 4/14/2026 had to move the samaactname generation up before existing search
            $SGSplat.SamAccountName =  -join (($SGSplat.DisplayName -replace '[^a-zA-Z0-9]', '').ToCharArray()[0..19]) ;
            try {$conflict = $null ;$conflict = get-adgroup -id $SGSplat.SamAccountName -server $InputSplat.DomainController ;} CATCH {} ;
            $incr = 1 ;
            Do{
                if($conflict){
                    $incr++ ;
                    $SGSplat.SamAccountName = "$( -join (($SGSplat.DisplayName -replace '[^a-zA-Z0-9]', '').ToCharArray()[0..18]) )$($incr)" ;
                }
                try {$conflict = $null ;$conflict = get-adgroup -id $SGSplat.SamAccountName -server $InputSplat.DomainController ;} CATCH {} ;
            }while($conflict) ; 
            #>

            # 4:52 PM 4/3/2026 owner is the alias, not samaccountname, use get-user to resolve it to the DN hard link regarldess of tree
            #$SGSplat.ManagedBy = $($InputSplat.Owner);
            if($SGSplat.ManagedBy = (get-user -id $inputsplat.owner -ea STOP).distinguishedname){} else { 
                throw "unable to resolve: get-user -id $($inputsplat.owner) " ; 
            } ;
            $SGSplat.Description = "Email - access to $($Tmbx.displayname)'s mailbox";
            $SGSplat.Server = $($InputSplat.DomainController) ;
            # build the Notes/Info field as a hashcode: OtherAttributes=@{    info="TargetMbx:SOMERECIP`r`nPermsExpire:6/19/2015"  } ;
            $SGSplat.OtherAttributes = @{info = $($Infostr) } ;


            $smsg = "`$SGSplat:`n---";
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
            $nADGRes = New-AdGroup @SGSplat -whatif ;
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
                    $nADGRes = New-AdGroup @SGSplat  ;
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
                        $adgRes = Add-DistributionGroupMember @pltAddDGM -member $mbr ; 
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

            #$mbxadp = $Tmbx | Get-ADPermission -domaincontroller $($InputSplat.domaincontroller) -ea Silentlycontinue |
            #     Where-Object { ($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and ($_.user -match ".*-(SEC|Data)-Email-.*$") };
            $mbxadp = Get-ADPermission -Identity $Tmbx.identity -domaincontroller $($InputSplat.domaincontroller) -ea Silentlycontinue |
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
                        $admpRes = add-mailboxpermission @GrantSplat ;
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
                        $addedadmbxp = add-adpermission -identity $($TMbx.Identity) @ADMbxGrantSplat ;
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
                                $dgRet = add-DistributionGroupMember @pltAddDGM ;
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

                $smsg = "`nUpdated $($oSG.Displayname) Membership...`n" ;
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
#endregion ADD_MAILBOXACCESSGRANT ; #*------^ END add-MailboxAccessGrant ^------
