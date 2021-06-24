# new-distributiongroupTOR.ps1

# dbg: cls ; new-distributiongroupTOR -displayname 'ENT-DL-IT-Perry Team Mgrs' -ManagedBy 'aabraham@charlesmachine.works' -members "aabraham@charlesmachine.works,Andy.Marquis@toro.com,David.Traub@toro.com,klooper@charlesmachine.works,scude@charlesmachine.works,dccoldiron@charlesmachine.works,Nik.Sharp@toro.com,Mike.Hillen@toro.com,Michael.Kellen@toro.com,Jeremy.Liebherr@toro.com,Justin.Misunas@toro.com,kredus@charlesmachine.works,Jake.Schultz@toro.com,Darren.Redetzke@toro.com" -ticket 625288 -verbose -whatif

<#
.SYNOPSIS
new-distributiongroupTOR.ps1 - Create new on-prem Exchange distributiongroup, using company standards
.NOTES
Version     : 0.0.
Author      : Todd Kadrie
Website     : http://www.toddomation.com
Twitter     : @tostka / http://twitter.com/tostka
CreatedDate : 2020-
FileName    : new-distributiongroupTOR.ps1
License     : MIT License
Copyright   : (c) 2020 Todd Kadrie
Github      : https://github.com/tostka/verb-XXX
Tags        : Powershell,Exchange,ExchangeOnline,ResourceMailbox,RoomMailbox,Capacity,Maintenance
REVISIONS
* 1:57 PM 6/23/2021 init
.DESCRIPTION
new-distributiongroupTOR.ps1 - Create new on-prem Exchange distributiongroup, using company standards
.PARAMETER DisplayName
Name for new DistributionGroup [[SIT]-DL-[descriptor]
.PARAMETER ManagedBy
Specify the recipient to be responsible for member-modification-approvals[name,emailaddr,alias]
.PARAMETER SiteOverride
Specify a 3-letter Site Code. Used to force DL placement to vary from ManagedBy's current site[-SiteOverride XXX]
.PARAMETER Members
Comma-delimited string of potential users to be granted access[name,emailaddr,alias]
.PARAMETER Ticket
Ticket number[-Ticket nnnnnn]
.PARAMETER NoPrompt
Suppress YYY confirmation prompts [-NoPrompt]
.PARAMETER domaincontroller
Option to hardcode a specific DC [-domaincontroller xxxx]
.PARAMETER TenOrg
TenantTag indicating hosting Tenant[-TenOrg XXX]
.PARAMETER USEexoV2
Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection (not yet implemented feature) [-useEXOv2]
.PARAMETER showDebug
Debugging output switch
.PARAMETER Whatif
Parameter to run a Test no-change pass (defaults true, use -whatif:$false to override) [-Whatif switch]
.INPUTS
None. Does not accepted piped input.(.NET types, can add description)
.OUTPUTS
None. Returns no objects or output (.NET types)
.EXAMPLE
PS> cls ; .\new-distributiongroupTOR.ps1 -groups 'ENT-DL-TestDDG' -MembersAdd 'SENDER1','SENDER2' -tickets 999999 -verbose -whatif ;
Update a single DDG, added 2 new approved senders to the new DistributionGroup, specifies a ticketnumber (for log entires), verbose, explicit whatif, (though default is whatif)
.EXAMPLE
PS> cls ; .\new-distributiongroupTOR.ps1 -groups 'ENT-DL-TestDDG','ENT-DL-TestDDG2' -MembersAdd 'SENDER1','SENDER2' -tickets '999999','999998' -verbose -whatif ;
Update an array of DDGs, with settings above, the common -membersAdd list are added to all groups in the array.
.EXAMPLE
PS> cls ; .\new-distributiongroupTOR.ps1 -groups 'SOMEDDG' -tickets 123456 -verbose -whatif:$false ;
Typical pass, no -membershadd, override default whatif to false (exec pass)
.LINK
https://github.com/tostka/powershell
#>
#Requires -Version 3
#requires -PSEdition Desktop
##Requires -Modules ActiveDirectory, AzureAD, MSOnline, ExchangeOnlineManagement, MicrosoftTeams, SkypeOnlineConnector, Lync,  verb-AAD, verb-ADMS, verb-Auth, verb-Azure, VERB-CCMS, verb-Desktop, verb-dev, verb-Ex2010, verb-EXO, verb-IO, verb-logging, verb-Mods, verb-Network, verb-L13, verb-SOL, verb-Teams, verb-Text, verb-transcript
#Requires -Modules verb-Ex2010, ExchangeOnlineManagement, verb-ADMS, verb-Auth, verb-EXO, verb-IO, verb-logging, verb-Text, verb-Network
##Requires -Modules verb-ADMS, verb-Auth, verb-Ex2010, verb-IO, verb-logging, verb-Text
#Requires -RunasAdministrator
# VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.toro\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
## [OutputType('bool')] # optional specified output type
[CmdletBinding()]
###[Alias('Alias','Alias2')]
PARAM(
    [Parameter(Mandatory = $true, HelpMessage = "Name for new DistributionGroup [[SIT]-DL-[descriptor]")]
    [string]$DisplayName,
    [Parameter(Mandatory = $true, HelpMessage = "Specify the recipient to be responsible for member-modification-approvals[name,emailaddr,alias]")]
    [array]$ManagedBy,
    [Parameter(Mandatory = $false,HelpMessage = "Specify a 3-letter Site Code. Used to force DL placement to vary from ManagedBy's current site[-SiteOverride XXX]")]
    [ValidatePattern("[a-zA-Z0-9]{3,6}")]
    [string]$SiteOverride,
    [Parameter(HelpMessage = "Comma-delimited string of potential users to be granted access[name,emailaddr,alias]")]
    [string]$Members,
    [Parameter(Mandatory = $true,HelpMessage = "Ticket number[-Ticket nnnnnn]")]
    [int]$Ticket,
    [Parameter(HelpMessage = "Suppress YYY confirmation prompts [-NoPrompt]")]
    [switch] $NoPrompt,
    [Parameter(HelpMessage = "Option to hardcode a specific DC [-domaincontroller xxxx]")]
    [string]$domaincontroller,
    [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag indicating hosting Tenant[-TenOrg XXX]")]
    [ValidateNotNullOrEmpty()]
    $TenOrg = 'TOR',
    [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection (not yet implemented feature) [-useEXOv2]")]
    [switch] $useEXOv2,
    [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
    [switch] $showDebug,
    [Parameter(HelpMessage="Whatif Flag (defaults true, override -whatif:`$false) [-whatIf]")]
    [switch] $whatIf=$true
) ;

BEGIN {
    # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
    ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
    # Get parameters this function was invoked with
    $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
    $Verbose = ($VerbosePreference -eq 'Continue') ;
    if ($PSScriptRoot -eq "") {
        if ($psISE) { $ScriptName = $psISE.CurrentFile.FullPath }
        elseif ($context = $psEditor.GetEditorContext()) {$ScriptName = $context.CurrentFile.Path }
        elseif ($host.version.major -lt 3) {
            $ScriptName = $MyInvocation.MyCommand.Path ;
            $PSScriptRoot = Split-Path $ScriptName -Parent ;
            $PSCommandPath = $ScriptName ;
        } else {
            if ($MyInvocation.MyCommand.Path) {
                $ScriptName = $MyInvocation.MyCommand.Path ;
                $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
            } else {throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$MyInvocation IS BLANK!" } ;
        };
        $ScriptDir = Split-Path -Parent $ScriptName ;
        $ScriptBaseName = split-path -leaf $ScriptName ;
        $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($ScriptName) ;
    } else {
        $ScriptDir = $PSScriptRoot ;
        if ($PSCommandPath) {$ScriptName = $PSCommandPath }
        else {
            $ScriptName = $myInvocation.ScriptName
            $PSCommandPath = $ScriptName ;
        } ;
        $ScriptBaseName = (Split-Path -Leaf ((& { $myInvocation }).ScriptName))  ;
        $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
    } ;
    if ($showDebug) { write-debug -verbose:$true "`$ScriptDir:$($ScriptDir)`n`$ScriptBaseName:$($ScriptBaseName)`n`$ScriptNameNoExt:$($ScriptNameNoExt)`n`$PSScriptRoot:$($PSScriptRoot)`n`$PSCommandPath:$($PSCommandPath)" ; } ;
    $ComputerName = $env:COMPUTERNAME ;
    $NoProf = [bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
    # silently stop any running transcripts
    $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
    #*======v SCRIPT/DOMAIN/MACHINE/INITIALIZATION-DECLARE-BOILERPLATE v======

    <#configure EXO EMS aliases to cover useEXOv2 requirements
    switch ($script:useEXOv2){
        $true {
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):using ExoV2 cmdlets" ;
            #reconnect-eXO2 @pltRXO ;
            set-alias ps1GetXRcp get-xorecipient ;
            set-alias ps1GetXMbx get-xomailbox ;
            set-alias ps1SetXMbx Set-xoMailbox ;
            set-alias ps1GetxUser get-xoUser ;
            set-alias ps1GetXCalProc get-xoCalendarprocessing ;
            set-alias ps1GetXMbxFldrPerm get-xoMailboxfolderpermission ;
            set-alias ps1GetXAccDom Get-xoAcceptedDomain ;
            set-alias ps1GGetXRetPol Get-xoRetentionPolicy ;
            set-alias ps1GetXDistGrp get-xoDistributionGroup ;
            set-alias ps1GetXDistGrpMbr get-xoDistributionGroupmember ;
            set-alias ps1GetMsgTrc get-xoMessageTrace ;
            set-alias ps1GetMsgTrcDtl get-xoMessageTraceDetail ;
        }
        $false {
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):using EXO cmdlets" ;
            #reconnect-exo @pltRXO
            set-alias ps1GetXRcp get-exorecipient ;
            set-alias ps1GetXMbx get-exomailbox ;
            set-alias ps1SetXMbx Set-exoMailbox ;
            set-alias ps1GetxUser get-exoUser ;
            set-alias ps1GetXCalProc get-exoCalendarprocessing  ;
            set-alias ps1GetXMbxFldrPerm get-exoMailboxfolderpermission  ;
            set-alias ps1GetXAccDom Get-exoAcceptedDomain ;
            set-alias ps1GGetXRetPol Get-exoRetentionPolicy
            set-alias ps1GetXDistGrp get-exoDistributionGroup  ;
            set-alias ps1GetXDistGrpMbr get-exoDistributionGroupmember ;
            set-alias ps1GetMsgTrc get-exoMessageTrace ;
            set-alias ps1GetMsgTrcDtl get-exoMessageTraceDetail ;
        } ;
    } ;  # SWTCH-E useEXOv2
    #>

    #$rgxBadName = '\w\d{1,2}$' ; # rooms with [letter][1-2digit#] at the end of their name

    #$tcsv = 'C:\usr\work\incid\619835-BLM-RoomResource-CapReductions-20210525-1133AM.csv' ;

    if(!$rgxDeadUserDN){$rgxDeadUserDN = '(,OU=|/)(Disabled|TERMedUsers)(,(DC|OU)=|/)'} ;

    if(!$global:Retries){$Retries = 4 ; } ;# number of re-attempts
    if(!$global:$RetrySleep){$RetrySleep = 5 ;} # seconds to wait between retries
    if(!$global:DomTORfqdn){throw "UNDEFINED `$DomTORfqdn!" ; BREAK ; }

    $rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$"




    $sBnr="#*======v $(${CmdletName}): v======" ;
    $smsg = $sBnr ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    #
    $UseOP=$false ;
    if((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro -AND (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server){
        $UseOP = $true ;
        $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`ENABLING use of OnPrem Ex system this pass." ;
        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
    } else {
        $UseOP = $false ;
        $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`nDISABLING use of OnPrem Ex system this pass." ;
        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
    } ;
    #

    $useEXO = $false ; # non-dyn setting, drives variant EXO reconnect & query code
    if($useEXO){
        #*------v GENERIC EXO CREDS & SVC CONN BP v------
        # o365/EXO creds
        <### Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile*
        $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
        Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
        .EXAMPLE
        $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
        Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
        .EXAMPLE
        $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
        Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
        ###>
        $o365Cred=$null ;
        <# $TenOrg is a mandetory param in this script, skip dyn resolution
        switch -regex ($env:USERDOMAIN){
            "(TORO|CMW)" {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
            "TORO-LAB" {$TenOrg = 'TOL' }
            default {
                throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ;
                BREAK ;
            } ;
        } ;
        #>
        if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -UserRole 'CSVC','SID' -verbose:$($verbose))){
            # make it script scope, so we don't have to predetect & purge before using new-variable
            New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
            $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } else {
            #-=-record a STATUSERROR=-=-=-=-=-=-=
            $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(gv passstatus -scope Script){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
            #-=-=-=-=-=-=-=-=
            $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
            BREAK ;
        } ;
        <### CALLS ARE IN FORM: (cred$($tenorg))
        $pltRXO = @{Credential = $Credential ; verbose = $($verbose) ; }
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            verbose = $($verbose) ; } ;
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
        else { reconnect-EXO @pltRXO } ;
        # or with Tenant-specific cred($Tenorg) lookup
        #$pltRXO creds & .username can also be used for AzureAD connections
        Connect-AAD @pltRXO ;
        ###>
        # configure splat for connections: (see above useage)
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            verbose = $($verbose) ; } ;
        #*------^ END GENERIC EXO CREDS & SVC CONN BP ^------
    } # if-E $useEXO

    if($UseOP){
        #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
        # do the OP creds too
        $OPCred=$null ;
        # default to the onprem svc acct
        $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC' ;verbose=$($verbose)} ;
        if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
            # make it script scope, so we don't have to predetect & purge before using new-variable
            New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
            $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } else {
            #-=-record a STATUSERROR=-=-=-=-=-=-=
            $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(gv passstatus -scope Script){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
            #-=-=-=-=-=-=-=-=
            $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
            BREAK ;
        } ;
        $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        <# CALLS ARE IN FORM: (cred$($tenorg))
        $pltRX10 = @{
            Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
            verbose = $($verbose) ; }
        ReConnect-Ex2010XO @pltRX10 ; # cross-prem conns
        Reconnect-Ex2010 @pltRX10 ; # local org conns
        #$pltRx10 creds & .username can also be used for local ADMS connections
        #>
        $pltRX10 = @{
            Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
            verbose = $($verbose) ; } ;
        # TEST
        #if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; BREAK ;}  ;
        # defer cx10/rx10, until just before get-recipients qry
        #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
        # connect to ExOP X10
        <#
        if($pltRX10){
            ReConnect-Ex2010 @pltRX10 ;
        } else { Reconnect-Ex2010 ; } ;
        #>
    } ;  # if-E $useEXOP
    # TEST
    #if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; BREAK ;}  ;


    # defer cx10/rx10, until just before get-recipients qry

    <# load ADMS
    $reqMods+="load-ADMS".split(";") ;
    if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; BREAK ;}  ;
    $smsg= "(loading ADMS...)" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    load-ADMS ;
    #>

    if($UseOP){
        # resolve $domaincontroller dynamic, cross-org
        # setup ADMS PSDrives per tenant
        if(!$global:ADPsDriveNames){
            $smsg = "(connecting X-Org AD PSDrives)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $global:ADPsDriveNames = mount-ADForestDrives -verbose:$($verbose) ;
        } ;
        if(($global:ADPsDriveNames|measure).count){
            $useEXOforGroups = $false ;
            $smsg = "Confirming ADMS PSDrives:`n$(($global:ADPsDriveNames.Name|%{get-psdrive -Name $_ -PSProvider ActiveDirectory} | ft -auto Name,Root,Provider|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # returned object
            #         $ADPsDriveNames
            #         UserName                Status Name
            #         --------                ------ ----
            #         DOM\Samacctname   True  [forestname wo punc]
            #         DOM\Samacctname   True  [forestname wo punc]
            #         DOM\Samacctname   True  [forestname wo punc]

        } else {
            #-=-record a STATUSERROR=-=-=-=-=-=-=
            $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(gv passstatus -scope Script){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
            #-=-=-=-=-=-=-=-=
            $smsg = "Unable to detect POPULATED `$global:ADPsDriveNames!`n(should have multiple values, resolved to $()"
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
            BREAK ;
        } ;
    }  else {
        # have to defer to get-azuread, or use EXO's native cmds to poll grp members
        # TODO 1/15/2021
        $useEXOforGroups = $true ;
        $smsg = "$($TenOrg):HAS NO ON-PREM ACTIVEDIRECTORY, DEFERRING ALL GROUP ACCESS & MGMT TO NATIVE EXO CMDS!" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    } ;

    <# EXO connection
    $reqMods+="connect-exo;Reconnect-exo;Disconnect-exo".split(";") ;
    if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; BREAK ;}  ;
    $smsg= "(loading EXO...)" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    # first run a disconnect-exo - ENSURE there's no existing conn, or it'll contaiminate the SenderDomain list.
    if($VerbosePreference = "Continue"){
        $VerbosePrefPrior = $VerbosePreference ;
        $VerbosePreference = "SilentlyContinue" ;
        $verbose = ($VerbosePreference -eq "Continue") ;
    } ;
    disconnect-exo ;
    # napalm it all, to ensure *nothing* survived the prior pass.
    get-pssession | remove-pssession
    if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
    else { reconnect-EXO @pltRXO } ;
    #>

    if($pltRX10){
        ReConnect-Ex2010 @pltRX10 ;
    } else { Reconnect-Ex2010 ; } ;

    # reenable VerbosePreference:Continue, if set, during mod loads
    if($VerbosePrefPrior -eq "Continue"){
        $VerbosePreference = $VerbosePrefPrior ;
        $verbose = ($VerbosePreference -eq "Continue") ;
    } ;

} ;  # BEGIN-E
PROCESS {


    $Error.Clear() ;
    # call func with $PSBoundParameters and an extra (includes Verbose)
    #call-somefunc @PSBoundParameters -anotherParam

        $procd++ ;

        # pre-clear dc, before querying
        $domaincontroller = $null ;
        # we don't know which subdoms may be in play
        pushd ; # cache pwd
        if( $tPsd = "$((Get-Variable  -name "$($TenOrg)Meta").value.ADForestName -replace $rgxDriveBanChars):" ){
            if(test-path $tPsd){
                $error.clear() ;
                TRY {
                    set-location -Path $tPsd -ea STOP ;
                    $objForest = get-adforest ;
                    $doms = @($objForest.Domains) ; # ad mod get-adforest vers
                    # do simple detect 2 doms (parent & child), use child (non-parent dom):
                    if(($doms|?{$_ -ne $objforest.name}|measure).count -eq 1){
                        $subdom = $doms|?{$_ -ne $objforest.name} ;
                        $domaincontroller = get-gcfastxo -TenOrg $TenOrg -Subdomain $subdom -verbose:$($verbose) |?{$_.length} ;
                        $smsg = "get-gcfastxo:returned $($domaincontroller)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } else {
                        # as this is just EX, and not AD search, open the forestview up - all Ex OP qrys will search entire forest
                        enable-forestview
                        $domaincontroller = $null ;
                        # use the forest root
                        # otherwise would have to recursive search an ADObj to find the host subdomain, and use it with: get-adXXX -server sub.domain.com
                    } ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg= "Failed to exec cmd because: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #-=-record a STATUSERROR=-=-=-=-=-=-=
                    $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    if(gv passstatus -scope Script){$script:PassStatus += $statusdelta } ;
                    if(gv -Name PassStatus_$($tenorg) -scope Script){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                    #-=-=-=-=-=-=-=-=
                    $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    popd ; # restore dir
                    BREAK #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                } ;
            } else {
                $smsg = "UNABLE TO FIND *MOUNTED* AD PSDRIVE $($Tpsd) FROM `$$($TENorg)Meta!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #-=-record a STATUSERROR=-=-=-=-=-=-=
                $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                if(gv passstatus -scope Script){$script:PassStatus += $statusdelta } ;
                if(gv -Name PassStatus_$($tenorg) -scope Script){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                #-=-=-=-=-=-=-=-=
                BREAK #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
            } ;
        } else {
            $smsg = "UNABLE TO RESOLVE PROPER AD PSDRIVE FROM `$$($TENorg)Meta!" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #-=-record a STATUSERROR=-=-=-=-=-=-=
            $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(gv passstatus -scope Script){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
            #-=-=-=-=-=-=-=-=
            BREAK #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
        } ;

        popd ; # cd to prior dir


        #Reconnect-EXO ;
        if($pltRX10){
            ReConnect-Ex2010 @pltRX10 ;
        } else { Reconnect-Ex2010 ; } ;
        $error.clear() ;



        $error.clear() ;
        TRY {

            $oManagedBy = get-recipient $ManagedBy ;

            switch ($oManagedBy.RecipientType ) {
                "UserMailbox" {
                    # no issue
                }
                "MailUser" {
                    # no issue
                }
                "MailContact"{
                        throw "Specified ManagedBy:$($InputSplat.Owner) is a MailContact:Non-Security Principal *cannot* be ManagedBy on an object!`nPlease specify a local Security Principal (Mailbox/MailUser/RemoteMailbox/ADUser) for the ManagedBy, and rerun the script" ;
                        Cleanup ; Exit ;
                }
                default {
                    throw "$($InputSplat.Owner) Not found, or unrecognized RecipientType" ;
                    Cleanup ; Exit ;
                }
            } ; # swtch-E

            $pltNDG.ManagedBy = $oManagedBy.primarysmtpaddress ;
            $smsg = "Explicit -ManagedBy specified, forcing to specified $($oManagedBy.primarysmtpaddress)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            $members = $members | get-recipient -domaincontroller $domaincontroller -ErrorAction Continue | select -expand distinguishedname  | select -unique ;

            # resolve OU
            if ($SiteOverride) {
                $SiteCode = $($SiteOverride);
            } else {
                # we need to use the OwnerMbx - Owner currently is the alias, we want the object with it's dn
                $SiteCode = $oManagedBy.identity.tostring().split("/")[1]  ;
            } ;
            if ($env:USERDOMAIN -eq $TORMeta['legacyDomain']) {
                $FindOU = "^OU=Distribution\sGroups,";
            } ELSEif ($env:USERDOMAIN -eq $TOLMeta['legacyDomain']) {
                # CN=Lab-SEC-Email-Thomas Jefferson,OU=Email Access,OU=SEC Groups,OU=Managed Groups,OU=LYN,DC=SUBDOM,DC=DOMAIN,DC=DOMAIN,DC=com
                $FindOU = "^OU=Distribution\sGroups,"; ;
            } else {
                throw "UNRECOGNIZED USERDOMAIN:$($env:USERDOMAIN)" ;
            } ;
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
            $smsg = "SiteCode:$SiteCode" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $smsg = "OU:$OU" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;


            #$SGSplat.DisplayName = "$($SiteCode)-SEC-Email-$($Tmbx.DisplayName)-G";

            $pltNDG=[ordered]@{
                DisplayName=$DisplayName;
                Name=$DisplayName;
                Members=$members ;
                DomainController=$domaincontroller;
                Alias=([System.Text.RegularExpressions.Regex]::Replace($DisplayName,"[^1-9a-zA-Z_]",""));
                ManagedBy=$oManagedBy;
                OrganizationalUnit = $OU ;
                ErrorAction = 'Stop' ;
                whatif=$($whatif);
            } ;

            if($existDG=get-distributiongroup -id $pltndg.alias -ResultSize 1 -ea 0){
                $pltSetDG=[ordered]@{
                    identity = $existDG.primarysmtpaddress ;
                    #Members=$members ; # not supported have to add-DistributionGroupMember them in on existings
                    DomainController=$domaincontroller;
                    ManagedBy=$oManagedBy;
                    whatif=$($whatif);
                    ErrorAction = 'Stop' ;
                } ;
                $smsg = "UpdateExisting DG:Set-DistributionGroup  w`n$(($pltSetDG|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                set-DistributionGroup @pltSetDG ;
                # pre-purge
                $prembrs = get-DistributionGroupMember -id $pltSetDG.identity ;
                $pltModDGMbr=[ordered]@{identity= $pltSetDG.identity ;whatif = $($whatif) ;erroraction = 'STOP'  ;confirm=$false ;}
                $smsg = "Clear existing members:remove-DistributionGroupMember w`n$(($pltModDGMbr|out-string).trim())`n$(($prembrs |out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #$prembrs | %{remove-DistributionGroupMember @$pltModDGMbr -Member $_.alias  } ;
                $prembrs.distinguishedname | remove-DistributionGroupMember @pltModDGMbr ;
                # get-DistributionGroupMember -id $pltSetDG.identity | remove-DistributionGroupMember -id $pltSetDG.identity –whatif:$($whatif) -ea STOP ;
                # then add validated from scratch
                $smsg = "re-add VALIDATED members:add-DistributionGroupMember w`n$(($pltModDGMbr|out-string).trim())`n$(($members|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $members | add-DistributionGroupMember @pltModDGMbr ;
            } else {
                $smsg = "New-DistributionGroup  w`n$(($pltNDG|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                New-DistributionGroup @pltNDG ;
            } ;
            if(!$whatif){
                $1F=$false ;Do {if($1F){Sleep -s 5} ;  write-host "." -NoNewLine ; $1F=$true ; } Until ($existDG = get-DistributionGroup $pltNDG.alias -dom $domaincontroller -EA 0) ;
                # set hidden (can't be done with new-dg command): -HiddenFromAddressListsEnabled
                $pltSetDG=[ordered]@{
                   identity = $existDG.primarysmtpaddress ;
                   HiddenFromAddressListsEnabled = $true ;
                    whatif=$($whatif);
                    ErrorAction = 'Stop' ;
                } ;
                $smsg = "HiddenFromAddressListsEnabled:UpdateExisting DG:Set-DistributionGroup  w`n$(($pltSetDG|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                set-DistributionGroup @pltSetDG ;

                $approvedsndrDLs = $preDDG.AcceptMessagesOnlyFromDLMembers ;
                $approvedsndrDLs += $existdg.DistinguishedName ;
                # now have a mix of DN & canonicalNames, resolve them all to DN syntasx, and de-dupe
                $approvedsndrDLs  = $approvedsndrDLs | get-recipient -domaincontroller $domaincontroller -ErrorAction Continue | select -expand distinguishedname  | select -unique ;

                $pltSetDDG=[ordered]@{
                    identity=$preDDG.alias;
                    AcceptMessagesOnlyFromDLMembers = $approvedsndrDLs ;
                    DomainController=$domaincontroller;
                    ErrorAction = 'Stop' ;
                    whatif=$($whatif);
                } ;
                $smsg = "Set-DynamicDistributionGroup  w`n$(($pltSetDDG|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Set-DynamicDistributionGroup @pltSetDDG ;
                $pddg = get-DynamicDistributionGroup -id $pltSetDDG.identity -dom $domaincontroller -ErrorAction 'Stop' ; ;

                $smsg = "POST:$($pddg.alias):`n$(($pddg | fl *name*,prim*,accept*|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $smsg = "(-whatif: skipping balance of process)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            }  ;

        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #-=-record a STATUSWARN=-=-=-=-=-=-=
            $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(Get-Variable passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
            if(Get-Variable -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
            #-=-=-=-=-=-=-=-=

            Continue ; #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
        } ;



        start-sleep -Milliseconds $ThrottleMs

} ;  # PROC-E
END {
    $smsg = "$($sBnr.replace('=v','=^').replace('v=','^='))" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    Stop-Transcript ;
} ;  # END-E



