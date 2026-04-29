# new-xopMailboxSREtdo.ps1

<#
.SYNOPSIS
new-xopMailboxSREtdo.ps1 - Wrapper that creates a new Exchange Server Conference room
.NOTES
Version     : 0.0.
Author      : Todd Kadrie
Website     : http://www.toddomation.com
Twitter     : @tostka / http://twitter.com/tostka
CreatedDate : 2026-
FileName    : new-xopMailboxSREtdo.ps1
License     : MIT License
Copyright   : (c) 2026 Todd Kadrie
Github      : https://github.com/tostka/verb-XXX
Tags        : Powershell
AddedCredit : REFERENCE
AddedWebsite: URL
AddedTwitter: URL
REVISIONS
* 3:56 PM 4/28/2026 repaired trailing summary output; along with changes below
* 11:46 AM 4/22/2026 recoded for native -type Shared support, also -mbxSpecFormat Email support; added post $EmailAddress align code; fixed move-exo* call, confirmed return and post works (needed to add -outobject to call)
* 9:44 AM 4/16/2026 ren: new-xopConfRmTDO.ps1 -> new-xopMailboxSREtdo.ps1 (Shared/Room/Equipment) ;
fix: grantees in creategrant...() was string array (needs to be comma-delim string for new-maiolboxshared()); dbg force load code, date specific; 
fixed typo in MailTip & ResourceCapacity ; typo in Cui5 return test from mbx; moved set-Calendarprocessing block into Initialize-ResourceCalendar(); 
EA stop missing quotes ; Set-ADUser won't use canonical dn; run through verb-io\ConvertFrom-CanonicalUser; finished half written set-aduser -company splat
ADD: Resolve-CustomDname() (need final Dname before can check for existing) ; show-MailboxInfo(),show-MailboxPermInfo(),show-CalProcInfo  (mbx,mbxperms,calProc report trailing summary); 
updated get-gcfast call to defer to get-addomaincontroller in cmw; Now builds and emmits trailing console notification letter summary for ticket (resolves $env:username into SID-> SID, and reverses HR mangled title for lettter).
fixed dawdle to track $moveTargets[-1]; updated move-Exomailboxnow to eimmit a PostStat w batchname -> this now echoes functional batch tracking & removal code

* 4:53 PM 4/10/2026 more substantial work getting the array of 8 differnt OU's and ou namings chemes for Migrations tree; built into a functional rule set for verb-adms\get-sitmbxou()
    also added logic/process info to the CBH description
    added -skipBookingConfig (prepping to direct convert this to cover new-xopSharedMailboxTDO.ps1 with minimal revision)
* 6:42 PM 4/9/2026 extended revising f'ery, around locating - manually polling, 
often creating - missing or non-standard OU storage for shared, room/equip & 
secgrps for email access - and then updated this, new-mailboxgenerictor; 
get-sitembxou; and new-mailboxshared to accomodate the migration blank-ups 
*10:32 AM 4/7/2026 ren: new-xopConfRm.ps1 -> new-xopConfRmTDO.ps1 ; 
add: RV_TRANSCRIPT_NODEP to output '$CmdName','$CmdPathed','$CmdParentDir' ; conditional push-TLSLatest w alert; revised banner to use new $CmdName; 
warn/prompt mixed $mbxspecs and -mailtip:$($mailtip) -resourcecustom:$($resourcecustom) -resourcecapacity:$($resourcecapacity)
review/yyy prompt mbxspecs per mbx
added -AdditionalResponse support (untested)
* 5:47 PM 4/6/2026 INIT; works used for 42416
.DESCRIPTION
new-xopMailboxSREtdo.ps1 - Wrapper that creates a new Exchange Server Conference room

## Performs the following commponent tasks for a Resource mailbox:
  - Creates a -type Room/Equipment mailbox
  - 

## Logic outline for the script:
1. Inputs are resolved: 
    - If -mbxSpecs with a populated mailbos specification array ("[DISPLAYNAME];[CU5];[OWNER];[GRANTEESARRAY]","DISPLAYNAME2;TORO;OWNER2;GRANTEE2A@toro.com,GRANTEE2B@toro.com" ), 
        The array of specification strings is looped: Each is split on the semicolons (;), to specify the -DisplayName, -cu5, -Owner, and -Grantees input parameters for the specified mailbox. 
        permitting bulk creation by stacking a comma-delimited quote-wrapped set of specifications for multiple mailboxes, with the varying specifications in use for the above parameter, 
        but common settings in use for the balance of the command-line parameters...
            -Location, -MailTip, -ResourceCustom, -ResourceCapacity, -Type, -BookinPolicy, -Ticket, -SiteOverride, -doSetFederationNote, -domaincontroller, -DCExclude, -DCServerPrefix, -whatIf'
    - If -mbxSpecs is *not* populated, the script expects to find explicit -DisplayName, -cu5, -Owner, and -Grantees input parameters for a single mailbox

2. CreateGrant-Mailbox() (internal function) is run against the specifications:
    a. The OwnerMailbox/RemoteMailbox is used to resolve the SiteCode/Hosting OU, and if CU5 isn't explicitly set, the CU% on the mailbox is set to match that of the Owner
    b. Custom displaynames - per CU5 setting - are applied for DIT|HAM|AUG|
    c. verb-ex2010\new-MailboxGenericTOR() preprocessor function is run against the specifications, to prep the inputs for the new mailbox (OnPrem)
        The verb-ex2010\New-MailboxShared() is then run to perform the actual mailbox creation. (which leverages verb-ADMS\get-SiteMbxOU() to resolve OUs that host mailbox and secgrp objects per site).
    d. A trailing 'REVIEW' report on mailbox specifications and status is output to console (and transcript/log)
    d. After the mailbox is confirmed present, verb-ex2010\add-MbxAccessGrant is run agains the Owner, and Grantees
        - Locates the SiteSpecific Email Access SecurityGroup OU for the destination SiteCode
        - Creates a "[$SiteCode]-SEC-Email-[mailbox DisplayName]-G"-named ADSecurity Group, then mail-enables the group.
        - Grants FullAccess Add-MailboxPermission and SendAs Add-ADPermissions to the new Security Group
        - Adds the Grantees to the ME SecGrp
        - Confirms permissions, and outputs a Permissions summary (granting group, Permissions, and Membership)

3. If -skipBookingConfig is *not* in use, the Room is then configured for booking using set-CalendarProcessing, as specified by the:
    -BookingType: (sets CalendarProcessing per the three variant templates: Open|Restricted|Vacation)
    -ResourceDelegates (sets Moderators that can approve bookings, for Restricted rooms)
    -BookInPolicy (override users that can *always* book the room, when AllBookInPolicy:$false)

    a. View Permissions on the Calendar Folder then grant all Delegates/Grantees Editor AccessRights
    c. A trailing summary of the Booking Settings, and effectie booking access is output to console. 
4. The script then waits for ADC/EntreID sync to occur and make the replicated mailbox MailUser recipient object visible in the cloud (via get-xorecipient cmdlet).
5. Once the mailbox is replicated, a -whatif pass is run using the move-EXOmailboxNow.ps1 script:
    a. The mailbox is evaluated for a suitable RemoteRoutingAddress (*onmicrosoft.com)
    b. The local MigrationEndPoints are validated functional
    c. And then a Migration test pass is run. 
    d. If -ImmediateMove was specified, a full migration move of the mailbox is initiated: (the move-EXOmailboxNow.ps1 script is rerun with -whatif:$false -NoTest)

         

.PARAMETER mbxSpecs
String Array of semi-colon-delimited room specifications, each is processed as a separate pass: string is split on semicolons to populate input parameters for target room, grantees are comma-delimited (permits bulk processing; OVERRIDES EXPLICIT PARAMETERS IN EACH NICHE) `"[DISPLAYNAME];[CU5];[OWNER];[GRANTEESARRAY]`"[-mbxSpecs `"Room Name;TORO;Aaaaa.Aaaaa@toro.com;Aaaaa.Aaaaa@toro.com,Aaaaa.Aaaaa@toro.com`"]
.PARAMETER MbxsFormat
string that designates if Mbxs array is displayname or email address specification (DNAME|EMAIL)[-MbxsFormat Dname]
.PARAMETER DisplayName
Displayname for new room[-DisplayName 'Room Name']
.PARAMETER EmailAddress
Optional EmailAddress (used to derive Displayname from raw email address specification)[-EmailAddress]
.PARAMETER cu5
CustomAttribute5 value for room (drives Brand assignement: (Americanaugers|Boss|Charlesmachine.works|Ditchwitch|Dripirrigation|Exmark|Hayter|HHtrenchless|Intimidatorutv|Irritrol|IrritrolEurope|Lawngenie|perrotde|perrotpl|ProkasroUSA|RadiusHDD|RainMaster|Spartanmowers|Subsite|TheToroCo|TheToroCompany|Toro.be|Toro.co.uk|Toro.hu|Torodistributor|ToroExmark|Torohosted|Toroused|Uniquelighting|Ventrac))[-cu5 Ditchwitch]
.PARAMETER Owner
Owner identifier[-Owner fname.lname@domain.com]
.PARAMETER Grantees
Grantees comma-delimited list of user-identifiers, as a single string[-Grantees 'fname.lname@domain.com,fname.lname@domain.com']
.PARAMETER Ticket
Ticket number[-ticket 123456]
.PARAMETER SiteOverride
SiteCode that overrides BaseUser/Owner Site settings[-SiteOverride LYN]
.PARAMETER ForwardingAddress
The ForwardingAddress parameter specifies ain *internal* - preconfigured MailContact, other mailbox etc -  forwarding address in your organization for messages that are sent to this mailbox. You can use any value that uniquely identifies the internal recipient[-ForwardingAddress RECIPIENT@DOMAIN.COM]
.PARAMETER Location
Override default Office field that normally inherits from Owner (used in some locations is in-building location description)[-Location 'At the Toro Bloomington 600 building (south Building)']
.PARAMETER MailTip
Add a message displayed to senders when they add the recipient to an e-mail message draft (or Meeting). For Room Resources, they appear above TO: line in OL, after the requestor switches to the Appointment display[-Mailtip 'This room resource is for requesting the *entire* Lyndale Cafe area']
.PARAMETER skipBookingConfig
Switch to skip the post Set-CalendarProcessing steps for the new Room Calendar.
.PARAMETER ResourceCustom
Room Assets, defaults to 'PolycomSpeakerPhone,Computer,Whiteboard,Projector'[-ResourceCustom @('PolycomSpeakerPhone','Computer','Whiteboard','Projector','Teams Room System')]
.PARAMETER ResourceCapacity
Room Capacity integer[-ResourceCapacity 10]
.PARAMETER AdditionalResponse
Specifies the additional information to be included in responses to meeting requests
.PARAMETER Type
Mailbox type (Equipment|Room|Shared)[-Type Room]
.PARAMETER BookingType
CalendarProcessing Booking type (Open|Restricted|Vacation) [BookingType 'Restricted']
.PARAMETER BookinPolicy
Users that can always book[BookinPolicy @('user1','user2')']
.PARAMETER ResourceDelegates
Moderator Users that must approve all bookings
.PARAMETER doSetFederationNote
Switch which populates mailbox CA6 with primarysmtpaddress domain, and CA11 with mail Org name[-doSetFederationNote]
.PARAMETER domaincontroller
Optional domaincontroller (skips discovery)[-domaincontroller 'Dc1']
.PARAMETER DCExclude
string array of problem DCs to exclude from use as -domaincontroller[-DCExclude 'Dc1','Dc2']
.PARAMETER DCServerPrefix
string which is used to filter target dcs (implies local site prefixes)[-DCServerPrefix = 'PRY']
.PARAMETER MoveImmediate
Switch to attempt an immediate move as soon as object is replicated to cloud[-MoveImmediate]
.PARAMETER whatIf
Whatif switch (defaults true)  [-whatIf]
.INPUTS
None. Does not accepted piped input.(.NET types, can add description)
.OUTPUTS
None. Returns no objects or output (.NET types)
System.Boolean
[| get-member the output to see what .NET obj TypeName is returned, to use here]
.EXAMPLE
PS> .\new-xopMailboxSREtdo.ps1 -Ticket 42416 -mbxspecs "APQP War Room;TORO;Aaaaa.Aaaaa@toro.com;Peter.Clark@toro.com" -Type Room 

    17:48:09:using common $domaincontroller:LYNMS8311
    #*------v PROCESSING (1/1) : APQP War Room v------
    17:48:12:CreateGrant-Mailbox w
    Name                           Value
    ----                           -----
    ticket                         42416
    DisplayName                    APQP War Room
    MInitial
    Owner                          Peter.Clark@toro.com
    SiteOverride
    NonGeneric                     False
    Vscan                          YES
    NoPrompt                       True
    domaincontroller               LYNMS8311
    showDebug                      True
    whatIf                         False
    Specified/Default -ResourceCustom:PolycomSpeakerPhone Computer Whiteboard Projector
    Specified -Location:At the Toro Bloomington 600 building (south Building)
    Overriding default Office: At the Toro Bloomington 600 building (south Building) with the specified value
    #*------v Set-Mailbox: v------
    17:48:18:PRE: Set-Mailbox:
    Identity       : global.ad.toro.com/LYN/Email Resources/APQP War Room
    ResourceCustom : {Whiteboard, Computer, PolycomSpeakerPhone, Projector}
    Office         : At the Toro Bloomington 600 building (south Building)
    17:48:18:Set-Mailbox w:
    Name                           Value
    ----                           -----
    Identity                       global.ad.toro.com/LYN/Email Resources/APQP War Room
    ErrorAction                    STOP
    whatif                         False
    ResourceCustom                 {PolycomSpeakerPhone, Computer, Whiteboard, Projector}
    Office                         At the Toro Bloomington 600 building (south Building)
    17:48:22:POST: Set-Mailbox:
    Identity       : global.ad.toro.com/LYN/Email Resources/APQP War Room
    ResourceCustom : {Whiteboard, Computer, PolycomSpeakerPhone, Projector}
    Office         : At the Toro Bloomington 600 building (south Building)
    17:48:23:
    #*------^ Set-Mailbox: ^------
    ===UPDATED APQP War Room ResourceCustom entries:
    Whiteboard
    Computer
    PolycomSpeakerPhone
    Projector
    ResourceCapacity:
    Location/Office:At the Toro Bloomington 600 building (south Building)
    MailTip:
    ===17:48:35:CONFIRMING PERMISSIONS:
    $DisplayName:Apqp War Room
    User
    ----
    TORO\LYN-SEC-Email-APQP War Room-G
    TORO\s-marchgj
    TORO\ExchangeAdmin
    TORO\SRVC_CWMailAdmin
    TORO\Exchange Domain Servers
    TORO\Exchange Full Administrators
    TORO\Exchange Administrators
    TORO\mckibids
    TORO\klausels
    TORO\Exchange Domain Servers
    TORO\s-marchgj
    TORO\ExchangeAdmin
    17:48:49:Updating room settings for APQPWarRoom
    17:48:59:
    #*------v set-CalendarProcessing:APQPWarRoom v------
    17:49:12:set-CalendarProcessing w:
    Name                           Value
    ----                           -----
    identity                       APQPWarRoom
    AddNewRequestsTentatively      True
    AddOrganizerToSubject          True
    AutomateProcessing             AutoAccept
    AllBookInPolicy                True
    AllRequestOutOfPolicy          False
    AllRequestInPolicy             False
    BookingWindowInDays            365
    BookInPolicy
    DeleteAttachments              False
    DeleteSubject                  False
    DeleteComments                 False
    DeleteNonCalendarItems         True
    EnableResponseDetails          True
    ForwardRequestsToDelegates     True
    OrganizerInfo                  True
    ResourceDelegates
    ErrorAction                    STOP
    whatif                         False
    #*------^ set-CalendarProcessing:APQPWarRoom ^------
    17:49:13:#*======v APQPWarRoom v======
    17:49:13:APQPWarRoom IS AN EX2010 MBOX
    17:49:14:
    #*------v ==MBX DELIVERY RESTRICTIONS (APQPWarRoom) v------
    AcceptMessagesOnlyFrom                 :
    AcceptMessagesOnlyFromDLMembers        :
    AcceptMessagesOnlyFromSendersOrMembers :
    Office                                 : At the Toro Bloomington 600 building (south Building)
    17:49:14:
    #*------^ ==MBX DELIVERY RESTRICTIONS (APQPWarRoom) ^------
    17:49:14:
    #*------v ==MBX DELIVERY RESTRICTIONS (APQPWarRoom) v------
    #*------v ==CALENDAR SETTINGS (APQPWarRoom): v------
    --STANDARD POLICY SETTINGS (APQPWarRoom):
    AllowConflicts                      : False
    BookingWindowInDays                 : 365
    MaximumDurationInMinutes            : 1440
    AllowRecurringMeetings              : True
    EnforceSchedulingHorizon            : True
    ScheduleOnlyDuringWorkHours         : False
    ConflictPercentageAllowed           : 0
    MaximumConflictInstances            : 0
    ForwardRequestsToDelegates          : True
    DeleteAttachments                   : False
    DeleteComments                      : False
    RemovePrivateProperty               : True
    DeleteSubject                       : False
    AddOrganizerToSubject               : True
    DeleteNonCalendarItems              : True
    EnableResponseDetails               : True
    OrganizerInfo                       : True
    RemoveOldMeetingMessages            : True
    RemoveForwardedMeetingNotifications : False
    AddNewRequestsTentatively           : True
    ProcessExternalMeetingMessages      : False
    --NOTIFICATION SETTINGS (APQPWarRoom):
    ForwardRequestsToDelegates AddAdditionalResponse AdditionalResponse
    -------------------------- --------------------- ------------------
                          True                 False
    --KEY BOOKING SETTINGS (APQPWarRoom):
    AutomateProcessing       : AutoAccept
    TentativePendingApproval : True
    *== Access Restrictions ==*:
    --ResourceDelegates:
    --BookInPolicy:
    --RequestInPolicy:
    --RequestOutOfPolicy:
    *== Open Access Settings ==*:
    AllBookInPolicy       : True
    AllRequestInPolicy    : False
    AllRequestOutOfPolicy : False
    17:49:14:
    #*------^ ==MBX DELIVERY RESTRICTIONS (APQPWarRoom) ^------
    #*------^ ==CALENDAR SETTINGS (APQPWarRoom): ^------
    ==CALENDAR VIEW PERMISSIONS (APQPWarRoom):
     + FolderName User      AccessRights
    ---------- ----      ------------
    Calendar   Default   {AvailabilityOnly}
    Calendar   Anonymous {None} +
    ===========
    ACCESS SUMMARY:
    AllBookInPolicy:OPEN BOOKING - All Users can book the room
    AllRequestInPolicy:$false: End users CANNOT REQUEST the room
    ResourceDelegates:Room has NO Moderators
    BookInPolicy:$false:Room has NO RESTRICTED users that can auto-book
    17:49:14:#*======^ APQPWarRoom ^======
    17:49:22:PREPARING DAWDLE LOOP!()
    MGOPLastSync:
    TimeGMT              TimeLocal          
    -------              ---------          
    4/6/2026 10:28:51 PM 4/6/2026 5:28:51 PM

Demo using bulk -mbxspecs array (specifies suite of settings per mailbox, as an array) with explicit -Type:room
.EXAMPLE
PS> $whatif = $true ;
PS> [array]$mbxSpecs = "DISPLAYNAME1;TORO;OWNER1;GRANTEE1A@toro.com,GRANTEE1B@toro.com" ; $mbxs += "DISPLAYNAME2;TORO;OWNER2;GRANTEE2A@toro.com,GRANTEE2B@toro.com" ; 
PS> $pltNCR=[ordered]@{
PS>   whatif = $($whatif) ;
PS>   Type="ROOM" ; # 'Equipment', 'Room', 'Shared'
PS>   Ticket = "TICKETNO" ;
PS>   mbxSpecs = $mbxSpecs ; # specs arraystrings defined above
PS>   SiteOverride = "" ; # 'LYN' to force unresolvable offices into an alte Site OU tree
PS> } ;
PS> if($env:userdomain -ne 'TORO'){$pltNCR.DCServerPrefix -eq $NULL}
PS> write-host -fore yellow "$((get-date).ToString('HH:mm:ss')):.\new-xopConfRm.ps1 w`n$(($pltNCR|out-string).trim())" ;
PS> $bRet=Read-Host "Review... Press any key to continue (CTRL+C TO ABORT)" ;
PS> .\new-xopMailboxSREtdo.ps1 @pltNCR ;
Simplified Splat for room/equipement/shared
.EXAMPLE
PS> $whatif = $true ;
PS> $pltNCR=[ordered]@{
PS> 	whatif = $true ;
PS> 	Type="ROOM" ; # 'Equipment', 'Room', 'Shared'
PS> 	Ticket = "TICKETNO" ;
PS> 	mbxSpecs = "";
PS> 	MbxsFormat = 'DNAME' ; #'Dname', 'Email'
PS> 	DisplayName = "DISPLAYNAME1"
PS> 	cu5 = "TORO" # (Americanaugers|Boss|Charlesmachine.works|Ditchwitch|Dripirrigation|Exmark|Hayter|HHtrenchless|Intimidatorutv|Irritrol|IrritrolEurope|Lawngenie|perrotde|perrotpl|ProkasroUSA|RadiusHDD|RainMaster|Spartanmowers|Subsite|TheToroCo|TheToroCompany|Toro.be|Toro.co.uk|Toro.hu|Torodistributor|ToroExmark|Torohosted|Toroused|Uniquelighting|Ventrac)
PS> 	Owner = "OWNER1"  ;
PS> 	Grantees = "GRANTEE1A@toro.com,GRANTEE1B@toro.com" ; # comma-delimited _not an array_
PS> 	Location = "" ; # overrides default Office:[which would be Owner's Office value], appears in booking UI as Location
PS> 	MailTip = "" ;
PS> 	ResourceCustom = @("PolycomSpeakerPhone","Computer","Whiteboard","Projector") ; # 'Easel','VCR','PolycomSpeakerPhone','SpeakerPhone','Whiteboard','Projector','Computer'
PS> 	ResourceCapacity = "" ; # booking capacity : IF REQUESTOR SPECIFIED, don't guess
PS> 	AdditionalResponse = "" ; # additional information to be included in responses to meeting requests (appends below booking email body to requestor)
PS> 	BookingType = "Open" ; #'Open','Restricted','Vacation'
PS> 	BookinPolicy = "" ; #Users that can always book
PS> 	ResourceDelegates = "" ; # Moderator Users that must approve all bookings
PS> 	SiteOverride = "" ; # "LYN" to force unresolvable offices into an alt Site OU tree
PS> 	skipBookingConfig = $false ; # skips normal set-CalendarProcessing etc for Room|Equipment calendars (normally triggered by -Type)
PS> 	MoveImmediate = $false ; # set $true, to initiate immed cloud move (otherwise runs -whatif pass, and echos move command)
PS> 	doSetFederationNote = $false ; # populates CA6 with primarysmtpaddress domain, and CA11 with mail Org name
PS> 	domaincontroller = "" ; # Optional domaincontroller (skips discovery)
PS> 	DCExclude = $null ; # problem DCs to exclude from auto-discover: 'Dc1','Dc2'"
PS> 	DCServerPrefix = 'LYN' ; # Prefix string used to filter target dcs (implies local site prefixes in TTC)
PS> } ;
PS> if($env:userdomain -ne 'TORO'){$pltNCR.DCServerPrefix -eq $NULL} ; 
PS> write-host -fore yellow "$((get-date).ToString('HH:mm:ss')):.\new-xopMailboxSREtdo.ps1 w`n$(($spltNCR|out-string).trim())" ; 
PS> .\new-xopMailboxSREtdo.ps1 @spltNCR ; 
FULL PARAMETERS SPLAT (FOR EXTENDED OPTIONS, AND NON -mbxspecs USE)
.LINK
https://github.com/tostka/verb-XXX
.LINK
https://github.com/tostka/powershellbb/
.LINK
[ name related topic(one keyword per topic), or http://|https:// to help, or add the name of 'paired' funcs in the same niche (enable/disable-xxx)]
#>

PARAM(
    [Parameter(Position = 0, Mandatory = $False, HelpMessage = "String Array of semi-colon-delimited room specifications, each is processed as a separate pass: string is split on semicolons to populate input parameters for target room, grantees are comma-delimited (permits bulk processing; OVERRIDES EXPLICIT PARAMETERS IN EACH NICHE) `"[DISPLAYNAME];[CU5];[OWNER];[GRANTEESARRAY]`"[-mbxSpecs `"Room Name;TORO;Aaaaa.Aaaaa@toro.com;Aaaaa.Aaaaa@toro.com,Aaaaa.Aaaaa@toro.com`"]")]
        [AllowEmptyString()][AllowNull()]
        <#[ValidateScript({
            if ($_ -AND ($_.tochararray() -contains ';') -AND ($_.split(';').count -eq 4)) {
                $true ; 
            }elseif (($null -eq $_) -AND ($DisplayName -AND $cu5 -AND $Owner -AND $Grantees)) {
                $TRUE ; 
            } else{
                throw "malformed -mbxspecs: Must contain semicolon delimters (;), and have FOUR elements, corresponding to `"[DISPLAYNAME];[CU5];[OWNER];[GRANTEESARRAY(comma-delimited)]`", per array member string!" ; 
            }            
        })]
        #>
        [string[]]$mbxSpecs,
        #[array]
    [Parameter(HelpMessage = "string that designates if Mbxs array is displayname or email address specification (DNAME|EMAIL)[-MbxsFormat Dname]")]
        [ValidateSet('Dname', 'Email')]
        [string]$MbxsFormat = 'Dname',
    [Parameter(HelpMessage = "Displayname for new room[-DisplayName]")]
        [string]$Identifier,
    [Parameter(HelpMessage = "Optional EmailAddress (used to derive Displayname from raw email address specification)[-EmailAddress]")]
        [string]$EmailAddress,
    [Parameter(HelpMessage = "CustomAttribute5 value for room (drives Brand assignement: (Americanaugers|Boss|Charlesmachine.works|Ditchwitch|Dripirrigation|Exmark|Hayter|HHtrenchless|Intimidatorutv|Irritrol|IrritrolEurope|Lawngenie|perrotde|perrotpl|ProkasroUSA|RadiusHDD|RainMaster|Spartanmowers|Subsite|TheToroCo|TheToroCompany|Toro.be|Toro.co.uk|Toro.hu|Torodistributor|ToroExmark|Torohosted|Toroused|Uniquelighting|Ventrac))[-cu5 Ditchwitch]")]
        [string]$cu5,
    [Parameter(HelpMessage = "Owner identifier[-owner fname.lname@domain.com]")]
        [string]$Owner,
    [Parameter(HelpMessage = "Grantees comma-delimited list of user-identifiers, as a single string[-Grantees 'fname.lname@domain.com,fname.lname@domain.com']")]        
        [string]$Grantees,    
    [Parameter(Mandatory = $False, HelpMessage = "Mailbox type (Equipment|Room|Shared)[-Type Room]")]
        [ValidateSet('Equipment', 'Room', 'Shared')]
        [ValidateNotNullOrEmpty()]
        [string]$Type = "ROOM",    
    [Parameter(Mandatory = $true, HelpMessage = "Ticket Number [-Ticket '999999']")]
        [string]$Ticket,
    [Parameter(HelpMessage = "SiteCode that overrides BaseUser/Owner Site settings[-SiteOverride LYN]")]
        [string]$SiteOverride,
    # 4:32 PM 4/22/2026 not implemented yet
    #[Parameter(HelpMessage = "The ForwardingAddress parameter specifies ain *internal* - preconfigured MailContact, other mailbox etc -  forwarding address in your organization for messages that are sent to this mailbox. You can use any value that uniquely identifies the internal recipient[-ForwardingAddress RECIPIENT@DOMAIN.COM]")]
    #    [string]$ForwardingAddress,
    [Parameter(HelpMessage = "Override default Office field that normally inherits from Owner (used in some locations is in-building location description)[-Location 'At the Toro Bloomington 600 building (south Building)']")]
        [string]$Location,
    [Parameter(HelpMessage = "Add a message displayed to senders when they add the recipient to an e-mail message draft (or Meeting). For Room Resources, they appear above TO: line in OL, after the requestor switches to the Appointment display[-Mailtip 'This room resource is for requesting the *entire* Lyndale Cafe area']")]
        [string]$MailTip,
    [Parameter(HelpMessage = "Room Assets, defaults to 'PolycomSpeakerPhone,Computer,Whiteboard,Projector'[-ResourceCustom @('PolycomSpeakerPhone','Computer','Whiteboard','Projector','Teams Room System')]")]
        [ValidateSet('Easel','VCR','PolycomSpeakerPhone','SpeakerPhone','Whiteboard','Projector','Computer')]
        [string[]]$ResourceCustom = @("PolycomSpeakerPhone","Computer","Whiteboard","Projector"),
    [Parameter(HelpMessage = "Room Capacity integer[-ResourceCapacity 10]")]
        [int32]$ResourceCapacity,
    [Parameter(HelpMessage = "Specifies the additional information to be included in responses to meeting requests")]
        [string]$AdditionalResponse,
    [Parameter(Mandatory = $False, HelpMessage = "Switch to skip the post Set-CalendarProcessing steps for the new Room Calendar.[-skipBookingConfig]")]
        [switch]$skipBookingConfig,
    [Parameter(Mandatory = $False, HelpMessage = "CalendarProcessing Booking type (Open|Restricted|Vacation) [-BookingType 'Restricted']")]
        [ValidateSet('Open','Restricted','Vacation')]
        [string]$BookingType='Open',
    [Parameter(HelpMessage = "Users that can always book[BookinPolicy @('user1','user2')']")]
        [string[]]$BookinPolicy,
    [Parameter(HelpMessage = "Moderator Users that must approve all bookings[-ResourceDelegates @('user1','user2')']")]
        [string[]]$ResourceDelegates,
    [Parameter(HelpMessage = "Switch which populates mailbox CA6 with primarysmtpaddress domain, and CA11 with mail Org name[-doSetFederationNote]")]
        [switch]$doSetFederationNote,
    [Parameter(HelpMessage = "Optional domaincontroller (skips discovery)[-domaincontroller 'Dc1']")]
        [string]$domaincontroller,
    [Parameter(HelpMessage = "string array of problem DCs to exclude from use as -domaincontroller[-DCExclude 'Dc1','Dc2']")]
        [string[]]$DCExclude,
    [Parameter(HelpMessage = "string which is used to filter target dcs (implies local site prefixes)[-DCServerPrefix = 'PRY']")]
        [string]$DCServerPrefix = 'LYN',
    [Parameter(HelpMessage = "Switch to attempt an immediate move as soon as object is replicated to cloud[-MoveImmediate]")]
        [switch]$MoveImmediate,
    [Parameter(HelpMessage = "Whatif Flag  [-whatIf]")]
        [switch] $whatIf=$true
)
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
    $dbgDate = '4/27/2026' # debugging ipmo force loads variants not in modules;  rgx: \$dbgDate\s=\s'\d{1,2}/\d{2}/\d{4}'
    $sQot = [char]34 ;
    $rgxOUMigrations = ',OU=_MIGRATIONS,DC=global,DC=ad,DC=toro,DC=com$' ;
    $rgxMigationsSite = ',OU=(\w+),OU=_MIGRATIONS,DC=global,DC=ad,DC=toro,DC=com$' ;
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

    #region FUNCTIONS_INTERNAL ; #*======v FUNCTIONS_INTERNAL v======

    #region CREATEGRANT_MAILBOX ; #*------v CreateGrant-Mailbox v------
    function CreateGrant-Mailbox{
        PARAM(
            [Parameter(Mandatory=$true,HelpMessage="Incident number for the change request[[int]nnnnnn]")]
                # [int] # 10:30 AM 10/13/2021 pulled, to permit non-numeric & multi-tix
                $Ticket,
            [Parameter(Mandatory=$true,HelpMessage="Display Name for mailbox [fname lname,genericname]")]
                [string]$DisplayName,
            [Parameter(HelpMessage="Middle Initial for mailbox (for non-Generic)[a]")]
                [string]$MInitial,
            [Parameter(Mandatory=$false,HelpMessage="Optionally force CU5 (variant domain assign) [-Cu5 Exmark]")]
                [string]$Cu5,
            [Parameter(Mandatory=$true,HelpMessage="Specify the userid to be responsible for access-grant-approvals[name,emailaddr,alias]")]
                [string]$Owner,
            [string]$Grantees,
            [Parameter(Mandatory = $False, HelpMessage = "Mailbox type (Equipment|Room|Shared)[-Type Room]")]
                [ValidateSet('Equipment', 'Room', 'Shared')]
                [ValidateNotNullOrEmpty()]
                [string]$Type = "ROOM",
            [Parameter(HelpMessage="Optional parameter indicating new mailbox Is NonGeneric-type[-NonGeneric `$true]")]
                [bool]$NonGeneric,
            [Parameter(HelpMessage="Suppress YYY confirmation prompts [-NoPrompt]")]
                [switch] $NoPrompt,
            [Parameter(HelpMessage="Optional parameter controlling Vscan (CU9) access (prompts if not specified)[-Vscan YES|NO|NULL]")]
                [string]$Vscan="YES",
            [Parameter(HelpMessage="Optionally specify a 3-letter Site Code o force OU placement to vary from Owner's current site[3-letter Site code]")]
                [string]$SiteOverride,
            [switch]$doSetFederationNote,
            [Parameter(HelpMessage="Option to hardcode a specific DC [-domaincontroller xxxx]")]
                [string]$domaincontroller,
            [Parameter(HelpMessage='Whatif Flag [$switch]')]
                [switch] $whatIf
        )
        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
        $pltNmbx = [ordered]@{  
            ticket                     = $ticket ;
            DisplayName                = "$($DisplayName)"  ;
            MInitial                   = $MInitial ;
            Owner                      = $Owner ;
            SiteOverride               = $SiteOverride ;
            NonGeneric                 = $false  ;
            Vscan                      = "YES" ;
            NoPrompt                   = $true ;
            domaincontroller           = $domaincontroller ;
            #showDebug                 = $true ;
            NoOutput                   = $false ; # new-mailboxGenericTOR() supports -noOutput, and I've now spliced in same into new-mailboxshared(); so we can now return the object by default wo rediscovery
            Verbose                    = $($VerbosePreference -eq 'Continue')
            whatIf                     = $($whatif) ;
        } ;
        if ( ($OwnerMbx = (get-mailbox -identity $($pltNmbx.Owner) -ea 0)) -OR ($OwnerMbx = (get-remotemailbox -identity $($pltNmbx.Owner) -ea 0)) ) {
            if ($ownermbx.DistinguishedName -match $rgxOUMigrations) {
                $OSiteCode = [regex]::Match($ownermbx.DistinguishedName, $rgxMigationsSite).groups[1].value ;
                if ($OSiteCode) {
                    $smsg = "Resolved _MIGRATIONS tree OSiteCode:$($OSiteCode)" ;
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
                } else {
                    throw "Unable to resolve $($pltNmbx.Owner.DistinguishedName) into a SiteCode" ;
                    Break ;
                }
            } else {
                $OSiteCode = $Ownermbx.identity.tostring().split('/')[1]
            } ;
        } else {
            throw "Unable to resolve $($pltNmbx.Owner) to any existing OP or EXO mailbox" ;
            Break ;
        } ;
        # 3:43 PM 4/10/2026: I'd add SiteCode "SUB","RAD","HAM","DIT","AUG" matching for custom's on the Site as well as CU5, but at this point it's pretty complicated code to hybrid with the CU5, if CU5 + conflicing SiteCode, which trumps? 
        # for now we'll assume -cu5 drives the custom dname settings. 
        # CU5 custom brand handling, custom Dnames, ref for Company names (implemented in post set)
        if (-not $cu5 -OR ($cu5 -eq "toro")) {
            # set Company to 'The Toro Company'
        }elseif ($cu5 -AND ($cu5 -ne "toro")) {
            write-verbose -verbose:$true "CU5:$($cu5): CUSTOM DOMAIN SPEC" ;
            $pltNmbx.add("CU5", $CU5) ;
            <# cU5 variants as of 9:22 AM 4/8/2026:
                (Americanaugers|Boss|Charlesmachine.works|Ditchwitch|Dripirrigation|Exmark|Hayter|HHtrenchless|
                Intimidatorutv|Irritrol|IrritrolEurope|Lawngenie|perrotde|perrotpl|ProkasroUSA|RadiusHDD|
                RainMaster|Spartanmowers|Subsite|TheToroCo|TheToroCompany|Toro.be|Toro.co.uk|Toro.hu|
                Torodistributor|ToroExmark|Torohosted|Toroused|Uniquelighting|Ventrac)
            #>
            # do AUG Americanaugers custom prefixes for equipement & rooms
            if($cu5 -match '^Americanaugers$'){
                # prefix room names with 'AA.\.', 
                # prefix equip names with 'r:\s', unless it's a vehicle, which gets 'v:\s', so far all vehicles are have keyword 'vehicle|truck|\scar'
                # set Company to 'American Augers, Inc.'
                switch($Type){
                    'Room'{  
                        if($displayname -notmatch "^AA\."){
                            $displayname = "AA.$($displayname)" ; 
                            $pltNmbx. DisplayName = $displayname ; 
                            $smsg = "UPDATED SPECIFIED DISPLAYNAME WITH $($cu5.toupper()) PREFIX" ;
                            $smsg += "`n$($displayname)" ; 
                            write-host -foregroundcolor yellow 
                        } ; 
                    }
                    'Equipement'{
                        if($displayname -notmatch "^AA\."){
                            if($displayname -match 'Vehicle-'){
                                $displayname = "AA.$($displayname)" ; 
                                $pltNmbx. DisplayName = $displayname ; 
                            }elseif($displayname -match 'truck|\scar'){
                                $displayname = "AA.Vehicle-$($displayname)" ; 
                                $pltNmbx. DisplayName = $displayname ; 
                            }else{
                                $displayname = "AA.$($displayname)" ; 
                                $pltNmbx. DisplayName = $displayname ; 
                            }
                            $smsg = "UPDATED SPECIFIED DISPLAYNAME WITH DITCHWITCH PREFIX" ;
                            $smsg += "`n$($displayname)" ; 
                            write-host -foregroundcolor yellow 
                        }
                    }
                    'Shared'{
                        # no action
                    }
                }
            } ; 
            # do DIT custom prefixes for equipement & rooms
            if($cu5 -match '^Ditchwitch$'){
                # prefix room names with 'r:\s', 
                # prefix equip names with 'r:\s', unless it's a vehicle, which gets 'v:\s', so far all vehicles are have keyword 'vehicle|truck|\scar'
                # set Company to 'Ditch Witch'
                switch($Type){
                    'Room'{  
                        if($displayname -notmatch "^r:\s"){
                            $displayname = "r: $($displayname)" ; 
                            $pltNmbx. DisplayName = $displayname ; 
                            $smsg = "UPDATED SPECIFIED DISPLAYNAME WITH $($cu5.toupper()) PREFIX" ;
                            $smsg += "`n$($displayname)" ; 
                            write-host -foregroundcolor yellow $smsg ;
                        } ; 
                    }
                    'Equipement'{
                        if(($displayname -notmatch "^r:\s") -AND ($displayname -notmatch "^v:\s")){
                            if($displayname -match 'vehicle|truck|\scar'){
                                $displayname = "v: $($displayname)" ; 
                                $pltNmbx. DisplayName = $displayname ; 
                            }else{
                                $displayname = "r: $($displayname)" ; 
                                $pltNmbx. DisplayName = $displayname ; 
                            }
                            $smsg = "UPDATED SPECIFIED DISPLAYNAME WITH DITCHWITCH PREFIX" ;
                            $smsg += "`n$($displayname)" ; 
                            write-host -foregroundcolor yellow $smsg ;
                        }
                    }
                    'Shared'{
                        # no action - NO NAME CUSTOMIZATION FOR SHARED
                    }
                }
            } ; 
            # CU5: Exmark.com @exmark.com
            if($cu5 -match '^Exmark\.com$'){                
                # set Company to 'Exmark Manufacturing Co'                
            } ; 
            # cu5: Hayter.co.uk @toro.com 
            if($cu5 -match '^Hayter\.co\.uk$'){                
                # set Company to 'Hayter Ltd'                
            } ; 
            # do HAM cu5: HHtrenchless @hhtrenchless.com
            if($cu5 -match '^HHtrenchless$'){
                # prefix room names with 'r: HH ', 
                # prefix equip names with 'r: HH ', unless it's a vehicle, which gets 'r: HH', so far all vehicles are have keyword 'vehicle|truck|\scar'
                # set Company to 'HammerHead Trenchless'
                switch($Type){
                    'Room'{  
                        if($displayname -notmatch "^r:\sHH\s"){
                            $displayname = "r: HH $($displayname)" ; 
                            $pltNmbx. DisplayName = $displayname ; 
                            $smsg = "UPDATED SPECIFIED DISPLAYNAME WITH $($cu5.toupper()) PREFIX" ;
                            $smsg += "`n$($displayname)" ; 
                            write-host -foregroundcolor yellow $smsg ;
                        } ; 
                    }
                    'Equipement'{
                        if(($displayname -notmatch "^r:\sHH\s") -AND ($displayname -notmatch "^v:\sHH\s")){
                            if($displayname -match 'vehicle|truck|\scar'){
                                $displayname = "v: HH $($displayname)" ; 
                                $pltNmbx. DisplayName = $displayname ; 
                            }else{
                                $displayname = "r: HH $($displayname)" ; 
                                $pltNmbx. DisplayName = $displayname ; 
                            }
                            $smsg = "UPDATED SPECIFIED DISPLAYNAME WITH DITCHWITCH PREFIX" ;
                            $smsg += "`n$($displayname)" ; 
                            write-host -foregroundcolor yellow $smsg ;
                        }
                    }
                    'Shared'{
                        # no action
                    }
                }
            } ; 
            # cu5:  @irritrol.com 
            if($cu5 -match '^Irritrol\.com$'){                
                # set Company to 'Irritrol'                
            } ; 
            # cu5:  IrritrolEurope.com @irritroleurope.com
            if($cu5 -match '^Irritrol\.com$'){                
                # set Company to 'Irritrol'                
            } ; 
            # cu5: Perrot.De  @perrot.de
            if($cu5 -match '^Perrot\.De$'){                
                # set Company to 'Perrot DE'                
            } ; 
            # cu5: Perrot.PL @perrot.pl
            if($cu5 -match '^Perrot\.PL$'){                
                # set Company to 'Perrot.PL'
            } ; 
            # cu5: Prokasrousa.com @prokasrousa.com
            if($cu5 -match '^Prokasrousa\.com$'){                
                # set Company to 'Prokasro USA'
            } ; 
            # do RAD # cu5: Radiushdd.com @radiushdd.com custom prefixes for equipement & rooms
            if($cu5 -match '^Radiushdd.com$'){
                # prefix room names with 'RAD\s', 
                # prefix equip names with 'RAD\s'
                # set Company to 'Radius HDD Direct, LLC'
                switch($Type){
                    'Room'{  
                        if($displayname -notmatch "RAD\s"){
                            $displayname = "RAD $($displayname)" ; 
                            $pltNmbx. DisplayName = $displayname ; 
                            $smsg = "UPDATED SPECIFIED DISPLAYNAME WITH $($cu5.toupper()) PREFIX" ;
                            $smsg += "`n$($displayname)" ; 
                            write-host -foregroundcolor yellow $smsg ;
                        } ; 
                    }
                    'Equipement'{
                        if($displayname -notmatch "^RAD\s"){
                            $displayname = "RAD $($displayname)" ; 
                            $pltNmbx. DisplayName = $displayname ; 
                            $smsg = "UPDATED SPECIFIED DISPLAYNAME WITH DITCHWITCH PREFIX" ;
                            $smsg += "`n$($displayname)" ; 
                            write-host -foregroundcolor yellow $smsg ;
                        }
                    }
                    'Shared'{
                        # no action
                    }
                }
            } ; 
            # cu5: Spartanmowers.com @spartanmowers.com
            if($cu5 -match '^Spartanmowers\.com$'){                
                # set Company to 'Spartan Mowers'
            } ; 
            # do SUB cu5: Subsite.com  @subsite.com
            if($cu5 -match '^(Subsite\.com)$'){
                # prefix room names with 'r:\s', 
                # prefix equip names with 'r:\s', unless it's a vehicle, which gets 'v:\s', so far all vehicles are have keyword 'vehicle|truck|\scar'
                # set Company to 'Subsite Electronics'
                switch($Type){
                    'Room'{  
                        if($displayname -notmatch "^r:\sSubsite-"){
                            $displayname = "r: $($displayname)" ; 
                            $pltNmbx. DisplayName = $displayname ; 
                            $smsg = "UPDATED SPECIFIED DISPLAYNAME WITH $($cu5.toupper()) PREFIX" ;
                            $smsg += "`n$($displayname)" ; 
                            write-host -foregroundcolor yellow $smsg ;
                        } ; 
                    }
                    'Equipement'{
                        if(($displayname -notmatch "r:\sSubsite-") -AND ($displayname -notmatch "^v:\sSubsite-")){
                            if($displayname -match 'vehicle|truck|\scar'){
                                $displayname = "v: Subsite-$($displayname)" ; 
                                $pltNmbx. DisplayName = $displayname ; 
                            }else{
                                $displayname = "r: Subsite-$($displayname)" ; 
                                $pltNmbx. DisplayName = $displayname ; 
                            }
                            $smsg = "UPDATED SPECIFIED DISPLAYNAME WITH $($cu5.toupper()) PREFIX" ;
                            $smsg += "`n$($displayname)" ; 
                            write-host -foregroundcolor yellow $smsg ;
                        }
                    }
                    'Shared'{
                        # no action
                    }
                }
            } ; 
            # cu5: Thetoroco.com @thetoroco.com
            if($cu5 -match '^Spartanmowers\.com$'){                
                # set Company to 'Thetoroco'
            } ; 
            #cu5: TheToroCompany.com @thetorocompany.com
            if($cu5 -match '^Spartanmowers\.com$'){                
                # set Company to 'TheToroCompany'
            } ; 
            # cu5:Toro.be @toro.com
            if($cu5 -match '^Toro\.be$'){                
                # set Company to 'ToroBE'
            } ; 
            # do VPI cu5: Ventrac.com @ventrac.com
            if($cu5 -match '^Ventrac\.com$'){                
                # set Company to 'Venture Products, Inc.'
            } ; 
        } ;
        # spec proper type praam
        switch ($Type) {
            "Shared" {
                #$pltNmbx.add('Shared',$true)
                # for new-MailboxGenericTOR, generic/shared is assumed, so we indicate it by setting NonGeneric',$false (shared isn't a parameter)
                if($pltNmbx.keys -contains 'NonGeneric'){
                    $pltNmbx.NonGeneric = $false ; 
                }else{
                    $pltNmbx.add('NonGeneric',$false) ; 
                }
            } 
            "Room" { $pltNmbx.add("Room", $true) } 
            "Equip" { $pltNmbx.add("Equip", $true) } 
        } ;
        write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):new-MailboxGenericTOR w`n$(($pltNmbx|out-string).trim())" ;
        $tCmdlet = 'new-MailboxGenericTOR' ; $BMod = 'VERB-ex2010' ; 
        if($psISE -AND ((get-date ).tostring() -match $dbgDate)){ 
            if((gcm $tCmdlet).source -eq $BMod){
                Do{
                    gci "D:\scripts\$($tCmdlet)_func.ps1" -ea STOP | ipmo -fo -verb  ;
                }until((gcm $tCmdlet).source -ne $BMod)
            } ;
        } ; 
        if($pltNmbx.NoOutput){
            new-MailboxGenericTOR @pltNmbx ;
        }else{
            $tmbx = new-MailboxGenericTOR @pltNmbx  ; 
        }
        if (-not $whatif) {
            if(-NOT $tmbx -AND $pltNmbx.NoOutput){
                write-host "waiting 10 secs..." ;
                start-sleep -seconds 10 ;
                Do {
                    write-host "." -NoNewLine;
                    Start-Sleep -m (1000 * 5) ;
                } Until (($tmbx = get-mailbox "$($DisplayName)" -domaincontroller $domaincontroller -ea 0)) ;
            }
            $pltGrant = [ordered]@{  
                ticket                      = $ticket  ;
                TargetID                    = $tmbx.samaccountname ;
                Owner                       = $Owner ;
                PermsDays                   = 999 ;
                members                     = $Grantees ;
                NoPrompt                    = $true ;
                domaincontroller            = $domaincontroller ;
                #showDebug                  = $true  ;
                Verbose                     = $($VerbosePreference -eq 'Continue')
                whatIf                      = $whatif ;
            } ;
            if ($SiteOverride) { $pltGrant.add('SiteOverride', $SiteOverride) } ;

            $tCmdlet = 'add-MbxAccessGrant' ; $BMod = 'VERB-ex2010' ;
            if($psISE -AND ((get-date ).tostring() -match $dbgDate)){
                if((gcm $tCmdlet).source -eq $BMod){
                    Do{
                        gci "D:\scripts\$($tCmdlet)_func.ps1" -ea STOP | ipmo -fo -verb  ;
                    }until((gcm $tCmdlet).source -ne $BMod)
                } ;
            } ;            

            write-host -foregroundcolor green "n$((get-date).ToString('HH:mm:ss')):===add-MbxAccessGrant w`n$(($pltGrant|out-string).trim())" ;
            add-MbxAccessGrant @pltGrant ;
            if ($doSetFederationNote ) {
                $accdoms = (get-accepteddomain).domainname ;
                if ($accdoms -contains ($tmbx.primarysmtpaddress.split('@')[1] )) {
                    $pltSrmbx = [ordered]@{                     identity = $tmbx.samaccountname ;
                        CustomAttribute6                                 = $tmbx.primarysmtpaddress.split('@')[1] ;
                        CustomAttribute11                                = (Get-OrganizationConfig).name ;
                        ErrorAction                                      = 'Stop' ;
                        Verbose                                          = $($VerbosePreference -eq 'Continue')
                        whatif                                           = $($whatif ) ;
                    } ;
                    $smsg = "populate CA6 (branddomain) & CA11 (Org), to mark source federation in cloud" ;
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):set-Mailbox w`n$(($pltSrmbx|out-string).trim())" ;
                    $error.clear() ;
                    TRY {
                        set-mailbox @pltSrmbx ;
                    } CATCH {
                        $ErrTrapd = $Error[0] ;
                        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } else { write-warning "$((get-date).ToString('HH:mm:ss')):$($tmbx.primarysmtpaddress) domain IS NOT AN ACCEPTEDDOMAIN IN THE $($ORG)! " } ;
            } ;
            #$moveTargets += $tmbx.alias ;
            write-host "returning new mbx alias to pipeline"
            #$tmbx.alias | write-output ;
            # 11:46 AM 4/23/2026 having issues with the alias, it's not get-xorecipient'able to confirm replication, emit the UPN
            $tmbx.userprincipalname | write-output 
        } else { write-host -foregroundcolor green "(-whatif, skipping acc grant)" };
        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
    } ;
    #endregion CREATEGRANT_MAILBOX ; #*------^ END CreateGrant-Mailbox ^------

    #region SET_MAILBOXROOMDETAILS ; #*------v Set-MailboxRoomDetails v------
    function Set-MailboxRoomDetails {
        PARAM(
            [Parameter(Mandatory = $false, HelpMessage = "Ticket Number [-Ticket '999999']")]
                [string]$Ticket,
            [Parameter(Mandatory=$true,HelpMessage="Mailbox object to be processed[-Targetmailbox]")]
                [ValidateNotNullOrEmpty()]
                $TargetMailbox,
            [Parameter(Mandatory = $true, HelpMessage = "Mailbox type (Equipment|Room|Shared)[-Type Room]")]
                [ValidateSet('Equipment', 'Room', 'Shared')]
                [ValidateNotNullOrEmpty()]
                [string]$Type = "ROOM",
            [Parameter(HelpMessage = "Override default Office field that normally inherits from Owner (used in some locations is in-building location description)[-Location 'At the Toro Bloomington 600 building (south Building)']")]
            [string]$Location,
            [Parameter(HelpMessage = "Add a message displayed to senders when they add the recipient to an e-mail message draft (or Meeting). For Room Resources, they appear above TO: line in OL, after the requestor switches to the Appointment display[-Mailtip 'This room resource is for requesting the *entire* Lyndale Cafe area']")]
                [string]$MailTip,
            [Parameter(HelpMessage = "Room Assets, defaults to 'PolycomSpeakerPhone,Computer,Whiteboard,Projector'[-ResourceCustom @('PolycomSpeakerPhone','Computer','Whiteboard','Projector','Teams Room System')]")]
                [ValidateSet('Easel','VCR','PolycomSpeakerPhone','SpeakerPhone','Whiteboard','Projector','Computer')]
                [string[]]$ResourceCustom = @("PolycomSpeakerPhone","Computer","Whiteboard","Projector"),
            [Parameter(HelpMessage = "Room Capacity integer[-ResourceCapacity 10]")]
                [int32]$ResourceCapacity,
            [Parameter(HelpMessage = "Specifies the additional information to be included in responses to meeting requests")]
                [string]$AdditionalResponse,
            [Parameter(HelpMessage = "Whatif Flag  [-whatIf]")]
                [switch] $whatIf
        )
        $pltSM=[ordered]@{
            Identity=$TargetMailbox.identity ;
            #ResourceCustom="PolycomSpeakerPhone","Computer","Whiteboard","Projector"  ;
            #ResourceCapacity=$null ;
            domaincontroller = $domaincontroller ; 
            ErrorAction = 'STOP' ; 
            Verbose = $($VerbosePreference -eq 'Continue')
            whatif=$($whatif);
        } ;
        if($TargetMailbox.recipienttypedetails -match 'EquipmentMailbox|RemoteEquipmentMailbox' -AND $ResourceCustom){
            write-host -foregroundcolor yellow "EquipmentMailbox does not support ResourceCustom, skipping" ; 
        } elseif($ResourceCustom){
            $pltSM.add('ResourceCustom',$ResourceCustom)
            $smsg = "Specified/Default -ResourceCustom:$($ResourceCustom)" ;             
            write-host -foregroundcolor yellow $smsg ;
        }
        if($Type -match 'Equipment|Room'){
            if($Location){
                $pltSM.add('Office',$Location)
                $smsg = "Specified -Location:$($Location)" ; 
                $smsg += "`nOverriding default Office: $($TargetMailbox.Office) with the specified value" ;
                write-host -foregroundcolor yellow $smsg ;
            } ; 
            if($MailTip){
                $pltSM.add('MailTip',$MailTip)
                $smsg = "Specified -MailTip:$($MailTip)" ;             
                write-host -foregroundcolor yellow $smsg ;
            }
            if($ResourceCapacity){
                $pltSM.add('ResourceCapacity',$ResourceCapacity)
                $smsg = "Specified -ResourceCapacity:$($ResourceCapacity)" ;             
                write-host -foregroundcolor yellow $smsg ;
            }
            if($AdditionalResponse){
                # has activation flag as well (only functions when AutomateProcessing: AutoAccept)
                $pltSM.add('AddAdditionalResponse',$true)
                $pltSM.add('AdditionalResponse',$AdditionalResponse)
                $smsg = "Specified -AdditionalResponse:`n$($AdditionalResponse)" ;             
                write-host -foregroundcolor yellow $smsg ;
            }
        
            $propsGM = $pltSM.GetEnumerator() |?{$_.name -notmatch 'whatif|erroraction'} | select -expand Name ;
            $sBnrS="`n#*------v Set-Mailbox:$($tTarget) v------" ;
            write-host -foregroundcolor CYAN "$((get-date).ToString('HH:mm:ss')):$($sBnrS)" ;
            write-host -foregroundcolor WHITE "`n$((get-date).ToString('HH:mm:ss')):PRE: Set-Mailbox:`n$((Get-Mailbox $pltSM.identity | fl $propsGM |out-string).trim())" ;
            write-host -foregroundcolor YELLOW "`n$((get-date).ToString('HH:mm:ss')):Set-Mailbox w:`n$(($pltSM|out-string).trim())`n" ;
            TRY{
                Set-Mailbox @pltSM ;
                $pObj=Get-Mailbox $pltSM.identity ;
            } CATCH {$ErrTrapd=$Error[0] ;
                write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
            } ;    
            write-host -foregroundcolor YELLOW "`n$((get-date).ToString('HH:mm:ss')):POST: Set-Mailbox:`n$(($pObj| fl $propsGM |out-string).trim())" ;
            write-host -foregroundcolor CYAN "$((get-date).ToString('HH:mm:ss')):$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
            $smsg = "===UPDATED $($pObj.displayname) ResourceCustom entries:`n$(($pObj| select -expand ResourceCustom | out-string).trim())"
            $smsg += "`nResourceCapacity:$($pObj.ResourceCapacity)" ;
            $smsg += "`nLocation/Office:$($pObj.Office)" ;
            $smsg += "`nMailTip:$($pObj.MailTip)" ;
            write-host -foregroundcolor green  $smsg ;
        }else{
            write-verbose "non-Type:Room|Equipment: skipping any specified Location,MailTip,ResourceCapacity,AdditionalResponse" ; 
        }
        #always set company any type
        switch -regex ( $pobj.CustomAttribute5){
            '^Americanaugers$'{$Company = 'American Augers, Inc.'}
            '^Ditchwitch$'{$Company = 'Ditch Witch'}
            '^Exmark\.com$'{$Company = 'Exmark Manufacturing Co'}
            '^Hayter\.co\.uk$'{$Company = 'Hayter Ltd'}
            '^HHtrenchless$'{$Company = 'HammerHead Trenchless'}
            '^Irritrol|IrritrolEurope\.com$'{$Company = 'Irritrol'}
            '^Perrot\.De$'{$Company ='Perrot DE' }
            '^Perrot\.PL$'{$Company = 'Perrot.PL' }
            '^Prokasrousa\.com$'{$Company = 'Prokasro USA'}
            '^Radiushdd.com$'{$Company = 'Radius HDD Direct, LLC'}
            '^Spartanmowers\.com$'{$Company = 'Spartan Mowers'}
            '^(Subsite\.com)$'{$Company = 'Subsite Electronics'}
            '^Thetoroco\.com$'{$Company = 'Thetoroco'}
            '^TheToroCompany\.com$'{$Company = 'TheToroCompany'}
            '^Toro\.be$'{$Company = 'ToroBE'}
            '^Ventrac\.com$'{$Company = 'Venture Products, Inc.'}
            "toro" {$Company = 'The Toro Company'}
            default{$Company = 'The Toro Company'}
        }
        # set custom company spec (isn't available in set-mailbox; have to use set-aduser or set-user)
        # SET-ADUSER -ID -Company
        $pltSAdu = [ordered]@{
            Identity = $pltSM.identity  ;
            Company = $Company ;
            ErrorAction = 'Stop' ;
            Verbose = $($VerbosePreference -eq 'Continue')
            server = $PLTsm.domaincontroller ;
        }
        # sadu won't use canonical dn, convert it
        if($pltsm.Identity.contains('/') -AND (gcm -name ConvertFrom-CanonicalUser -ea STOP)){
            $pltSAdu.Identity = (ConvertFrom-CanonicalUser -CanonicalName $pltSM.identity)
        }ELSE{
            $smsg = "Identity is Canonical DN, but MISSING:VERB-IO\ConvertFrom-CanonicalUser()!" ; 
            write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ; 
            throw $smsg ; 
        } ; 
        write-host -foregroundcolor YELLOW "`n$((get-date).ToString('HH:mm:ss')):SET COMPANY: Set-ADUser w:`n$(($pltSAdu|out-string).trim())`n" ;
        TRY{
            set-aduser @pltSAdu
        } CATCH {
            $ErrTrapd=$Error[0] ;
            write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
            write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
        } ;
        $pObj | write-output 
    } ; 
    #endregion SET_MAILBOXROOMDETAILS ; #*------^ END Set-MailboxRoomDetails ^------

    #region INITIALIZE_RESOURCECALENDAR ; #*------v Initialize-ResourceCalendar v------
    function Initialize-ResourceCalendar {
        <#
        .PARAMETER Ticket
        Ticket number[-ticket 123456]
        .PARAMETER Grantees
        Grantees comma-delimited list of user-identifiers, as a single string[-Grantees 'fname.lname@domain.com,fname.lname@domain.com']
        .PARAMETER BookingType
        CalendarProcessing Booking type (Open|Restricted|Vacation) [BookingType 'Restricted']
        .PARAMETER BookinPolicy
        Users that can always book[BookinPolicy @('user1','user2')']
        .PARAMETER ResourceDelegates
        Moderator Users that must approve all bookings
        .PARAMETER ResourceCustom
        Room Assets, defaults to 'PolycomSpeakerPhone,Computer,Whiteboard,Projector'[-ResourceCustom @('PolycomSpeakerPhone','Computer','Whiteboard','Projector','Teams Room System')]
        .PARAMETER ResourceCapacity
        Room Capacity integer[-ResourceCapacity 10]
        .PARAMETER AdditionalResponse
        Specifies the additional information to be included in responses to meeting requests
        .PARAMETER whatIf
        Whatif switch (defaults true)  [-whatIf]
        .EXAMPLE

        #>
        PARAM(
            [Parameter(Mandatory = $false, HelpMessage = "Ticket Number [-Ticket '999999']")]
                [string]$Ticket,
            [Parameter(Mandatory=$true,HelpMessage="Mailbox object to be processed[-Targetmailbox]")]
                [ValidateNotNullOrEmpty()]
                $TargetMailbox,
            [Parameter(Mandatory = $False, HelpMessage = "CalendarProcessing Booking type (Open|Restricted|Vacation) [-BookingType 'Restricted']")]
                [ValidateSet('Open','Restricted','Vacation')]
                [string]$BookingType='Open',
            [Parameter(HelpMessage = "Users that can always book[BookinPolicy @('user1','user2')']")]
                [string[]]$BookinPolicy,
            [Parameter(HelpMessage = "Moderator Users that must approve all bookings[-ResourceDelegates @('user1','user2')']")]
                [string[]]$ResourceDelegates,            
            [Parameter(HelpMessage = "Grantees comma-delimited list of user-identifiers, as a single string[-Grantees 'fname.lname@domain.com,fname.lname@domain.com']")]        
                [string]$Grantees,
            <#
            [Parameter(HelpMessage = "Room Assets, defaults to 'PolycomSpeakerPhone,Computer,Whiteboard,Projector'[-ResourceCustom @('PolycomSpeakerPhone','Computer','Whiteboard','Projector','Teams Room System')]")]
                [ValidateSet('Easel','VCR','PolycomSpeakerPhone','SpeakerPhone','Whiteboard','Projector','Computer')]
                [string[]]$ResourceCustom = @("PolycomSpeakerPhone","Computer","Whiteboard","Projector"),
            [Parameter(HelpMessage = "Room Capacity integer[-ResourceCapacity 10]")]
                [int32]$ResourceCapacity,
            [Parameter(HelpMessage = "Specifies the additional information to be included in responses to meeting requests")]
                [string]$AdditionalResponse,
            [Parameter(HelpMessage = "Optional domaincontroller (skips discovery)[-domaincontroller 'Dc1']")]
                [string]$domaincontroller,
            #>
            [Parameter(HelpMessage = "Whatif Flag  [-whatIf]")]
                [switch] $whatIf
        )
        BEGIN{
        
        }
        PROCESS{
            foreach ($MBX in $TargetMailbox) {
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Updating room settings for $($MBX.userprincipalname)" ;
                #set-EXORoom -Identity $MBX.userprincipalname -AutomateProcessing AutoAccept -BookingWindowInDay 365 -EnforceSchedulingHorizon $true -ScheduleOnlyDuringWorkHours $false -ShowInAddressBook $true -whatif:$whatif ;
                switch($BookingType){
                    'Open'{
                        $CalendarProcessingDefaults = [ordered]@{
                            Identity = $null ;
                            AddNewRequestsTentatively                           = $true ;
                            AddOrganizerToSubject                               = $true ;
                            AutomateProcessing                                  = "AutoAccept" ;
                            AllBookInPolicy                                     = $true ;
                            AllRequestOutOfPolicy                               = $false ;
                            AllRequestInPolicy                                  = $false ;
                            BookingWindowInDays                                 = 365 ;
                            BookInPolicy                                        = $null ;
                            DeleteAttachments                                   = $false ;
                            DeleteSubject                                       = $false ;
                            DeleteComments                                      = $false ;
                            DeleteNonCalendarItems                              = $true ;
                            EnableResponseDetails                               = $True;
                            ForwardRequestsToDelegates                          = $True;
                            OrganizerInfo                                       = $True;
                            ResourceDelegates                                   = $null ;
                            Verbose                                             = $($VerbosePreference -eq 'Continue')
                            ErrorAction                                         = 'STOP' ; 
                            whatif                                              = $($whatif) ;
                        } ;
                    }'restricted'{
                        # splat to reset/clear prior booking settings
                        $spltBookingCLEAR = [ordered]@{
                            identity = $null ;
                            ResourceDelegates                     = $null ;
                            BookInPolicy                          = $null ;
                            ErrorAction                           = 'STOP' ; 
                            Verbose                               = $($VerbosePreference -eq 'Continue')
                            whatif                                = $($whatif) ;
                        }
                        $spltBookingCLEAR.identity = $MBX.userprincipalname ;
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):PreBlank the specs:`nset-CalendarProcessing w`n$(($spltBookingCLEAR|out-string).trim())" ;                    
                        $error.clear() ;
                        TRY {
                            set-CalendarProcessing @spltBookingCLEAR ;
                        } CATCH {
                            $ErrTrapd = $Error[0] ;
                            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                            write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)"
                        } ;
                        # update splat
                        $CalendarProcessingDefaults = [ordered]@{
                            identity              = $null ;
                            ResourceDelegates     = $null ;
                            AllBookInPolicy       = $false ;
                            BookInPolicy          = $null;
                            AllRequestInPolicy    = $true ;
                            RequestOutOfPolicy    = $null ;
                            AllRequestOutOfPolicy = $false ;
                            RequestInPolicy       = $null ;
                            ErrorAction           = 'STOP' ; 
                            Verbose               = $($VerbosePreference -eq 'Continue')
                            whatif                = $($whatif) ;
                        } ;
                        if($ResourceDelegates){
                            $CalendarProcessingDefaults.ResourceDelegates = $ResourceDelegates ;
                        }
                        if($BookInPolicy){
                            $CalendarProcessingDefaults.BookInPolicy = $BookInPolicy ;
                        }
                        # do folder access grants
                        $AccessRights = "Editor"
                        write-host -f yellow "===$($MBX.userprincipalname ) - Grants" ;
                        $pltGrant = [ordered]@{
                            identity = "$($MBX.userprincipalname):\Calendar" ;
                            user                              = $null ;
                            AccessRights                      = $AccessRights ;
                            Verbose                           = $($VerbosePreference -eq 'Continue')
                            whatif                            = $($whatif) ;
                        } ;
                        $props = $pltGrant.GetEnumerator() |?{$_.name -notmatch 'whatif|erroraction'} | select -expand Name ;
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):PRE: Get-xoCalendarProcessing `n$((Get-xoMailboxFolderPermission -id $pltGrant.identity | fl $props |out-string).trim())" ;
                        foreach ($grantee in (($Grantees -split ',')|%{$_.trim()})) {
                            write-host "--$($grantee)" ;
                            $pltGrant.user = $grantee ;
                            if (Get-xoMailboxFolderPermission -id $pltGrant.identity -User $pltGrant.$grantee) {
                                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Set-xoMailboxFolderPermission w`n$(($pltGrant|out-string).trim())" ;
                                $Exit = 0 ;
                                Do {
                                    Try {
                                        Set-xoMailboxFolderPermission @pltGrant  ;
                                        $Exit = $Retries ;
                                    } Catch {
                                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                                        Start-Sleep -Seconds $RetrySleep ;
                                        rxo ;
                                        $Exit ++ ;
                                        Write-Verbose "Try #: $Exit" ;
                                        If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                                    }  ;
                                } Until ($Exit -eq $Retries) ;
                            } else {
                                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Add-xoMailboxFolderPermission w`n$(($pltGrant|out-string).trim())" ;
                                $Exit = 0 ;
                                Do {
                                    Try {
                                        Add-xoMailboxFolderPermission @pltGrant  ;
                                        $Exit = $Retries ;
                                    } Catch {
                                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                                        Start-Sleep -Seconds $RetrySleep ;
                                        rxo2 ;
                                        $Exit ++ ;
                                        Write-Verbose "Try #: $Exit" ;
                                        If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                                    }  ;
                                } Until ($Exit -eq $Retries) ;
                            } ;
                        } ;
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):POST: Get-xoCalendarProcessing `n$((Get-xoMailboxFolderPermission -id $pltGrant.identity | fl $props |out-string).trim())" ;
                    }
                    'vacation'{
                        $CalendarProcessingDefaults = [ordered]@{
                            Identity                  = $null ;
                            AutomateProcessing        = "AutoAccept" ;
                            AllowConflicts            = $true ;
                            ConflictPercentageAllowed = 100 ;
                            MaximumConflictInstances  = 1000 ;
                            BookingWindowInDays       = 365 ;
                            MaximumDurationInMinutes  = 525600 ;
                            EnforceSchedulingHorizon  = $True ;
                            DeleteAttachments         = $false ;
                            DeleteSubject             = $false ;
                            DeleteComments            = $false ;
                            DeleteNonCalendarItems    = $true ;
                            AddOrganizerToSubject     = $true ;
                            ErrorAction               = 'STOP' ; 
                            Verbose                   = $($VerbosePreference -eq 'Continue')
                            whatif                    = $($whatif) ;
                        } ;
                    }
                }
                $sBnrS = "`n#*------v set-CalendarProcessing:$($MBX.userprincipalname) v------" ;
                write-host -foregroundcolor CYAN "$((get-date).ToString('HH:mm:ss')):$($sBnrS)" ;
                $CalendarProcessingDefaults.identity = $MBX.userprincipalname ;
                $props = $CalendarProcessingDefaults.GetEnumerator() |?{$_.name -notmatch 'whatif|erroraction'} | select -expand name ; 
                write-host -foregroundcolor WHITE "`n$((get-date).ToString('HH:mm:ss')):PRE: set-CalendarProcessing:`n$((get-CalendarProcessing $CalendarProcessingDefaults.identity | fl $props |out-string).trim())" ;
                write-host -foregroundcolor YELLOW "`n$((get-date).ToString('HH:mm:ss')):set-CalendarProcessing w:`n$(($CalendarProcessingDefaults|out-string).trim())`n" ;
                if($spltBooking.ResourceDelegates){
                    write-host "::ResourceDelegates:`n$(($spltBooking.ResourceDelegates|out-string).trim())"
                } ;
                if($spltBooking.BookInPolicy){
                    write-host "::BookInPolicy:`n$(($spltBooking.BookInPolicy|out-string).trim())"
                } ;            
                TRY {
                    set-CalendarProcessing @CalendarProcessingDefaults ;
                    $pObj = get-CalendarProcessing $CalendarProcessingDefaults.identity ;
                } CATCH {
                    $ErrTrapd = $Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
                write-host -foregroundcolor YELLOW "`n$((get-date).ToString('HH:mm:ss')):POST: set-CalendarProcessing:`n$(($pObj| fl $props |out-string).trim())" ;
                if ($spltBooking.ResourceDelegates) {
                    write-host -foregroundcolor yellow  "::ResourceDelegates:`n$(($pCP.ResourceDelegates|out-string).trim())"
                } ;
                if ($spltBooking.BookInPolicy) {
                    write-host -foregroundcolor yellow  "::BookInPolicy:`n$(($pCP.BookInPolicy|out-string).trim())"
                } ;
                write-host -foregroundcolor CYAN "$((get-date).ToString('HH:mm:ss')):$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
            } ; # loop-E creategrant CalProc          
        }
        END{
            foreach ($MBX in $TargetMailbox) {
                $oReport = [ordered]@{
                    Mailbox = $null ; 
                    CalendarProcessing = $null ; 
                    CalendarFolderPermission = $null ; 
                } ; 
                $verbose = $FALSE ;
                $sBnr = "#*======v $($MBX.userprincipalname ) v======" ;
                write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
                if (!(gcm get-recipient -ea 0)) { rx10 } ;
                $OpRcp = get-recipient $MBX.userprincipalname  ;
                $CPrcps = "ResourceDelegates", "BookInPolicy", "RequestInPolicy", "RequestOutOfPolicy";
                switch ($OpRcp.recipienttype) {
                    "MailUser" {
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($MBX.userprincipalname ) IS AN EXO MBOX" ;
                        reconnect-exo2 ;
                        set-alias ps1grcp get-xorecipient ;
                        set-alias ps1gmbx get-xomailbox ;
                        set-alias ps1gcp get-xocalendarprocessing ;
                        set-alias ps1gmfp get-xomailboxfolderpermission ;
                    } ;
                    "UserMailbox" {
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($MBX.userprincipalname ) IS AN EX2010 MBOX" ;
                        reconnect-ex2010 ;
                        set-alias ps1grcp get-recipient ;
                        set-alias ps1gmbx get-mailbox ;
                        set-alias ps1gcp get-calendarprocessing ;
                        set-alias ps1gmfp get-mailboxfolderpermission ;
                    } ;
                } ;
                $oReport.Mailbox = ps1gmbx -identity $MBX.userprincipalname  ;
                $oReport.CalendarProcessing = ps1gcp -identity $MBX.userprincipalname  ;
                $oReport.CalendarFolderPermission = ps1gmfp -Identity "$($MBX.userprincipalname ):\Calendar" ;
                <# 9:33 AM 4/15/2026 flipping to nested obj, can't use gv to locate props
                foreach ($prop in $CPrcps) {
                    write-verbose "Expand:$($prop) into `$o$($prop)" ;
                    set-variable -name "o$($prop)" -value ((((gv CalendarProcessing).value.psobject.properties | ? { $_.name -eq $prop }).value | ps1grcp | select -expand primarysmtpaddress) -join ';')  ;
                } ;
                #>
                if($oreport.CalendarProcessing.ResourceDelegates){ $oResourceDelegates = ($oreport.CalendarProcessing.ResourceDelegates | ps1grcp | select -expand primarysmtpaddress) -join ';'}
                if($oreport.CalendarProcessing.BookInPolicy){ $oBookInPolicy = ($oreport.CalendarProcessing.BookInPolicy | ps1grcp | select -expand primarysmtpaddress) -join ';'}
                if($oreport.CalendarProcessing.RequestInPolicy){ $oRequestInPolicy = ($oreport.CalendarProcessing.RequestInPolicy | ps1grcp | select -expand primarysmtpaddress) -join ';'}
                if($oreport.CalendarProcessing.RequestOutOfPolicy){ $oRequestOutOfPolicy = ($oreport.CalendarProcessing.RequestOutOfPolicy | ps1grcp | select -expand primarysmtpaddress) -join ';'}

                $sBnrS = "`n#*------v ==MBX DELIVERY RESTRICTIONS ($($MBX.userprincipalname )) v------" ;
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnrS)" ;
                $prpsMbx = @{Name = 'AcceptMessagesOnlyFrom'; Expression = { $_.AcceptMessagesOnlyFrom | select DistinguishedName | out-string } }, @{Name = 'AcceptMessagesOnlyFromDLMembers'; Expression = { $_.AcceptMessagesOnlyFromDLMembers } }, @{Name = 'AcceptMessagesOnlyFromSendersOrMembers'; Expression = { $_.AcceptMessagesOnlyFromSendersOrMembers | select DistinguishedName | out-string } }, 'Office';
                write-host "$(($oReport.Mailbox | fl  $prpsMbx |out-string).trim())";
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
                $sBnrS += "`n#*------v ==CALENDAR SETTINGS ($($MBX.userprincipalname )): v------" ;
                $prpsCPPol = 'AllowConflicts', 'BookingWindowInDays', 'MaximumDurationInMinutes', 'AllowRecurringMeetings', 'EnforceSchedulingHorizon', 'ScheduleOnlyDuringWorkHours', 'ConflictPercentageAllowed', 'MaximumConflictInstances', 'ForwardRequestsToDelegates', 'DeleteAttachments', 'DeleteComments', 'RemovePrivateProperty', 'DeleteSubject', 'AddOrganizerToSubject', 'DeleteNonCalendarItems', 'EnableResponseDetails', 'OrganizerInfo', 'RemoveOldMeetingMessages', 'RemoveForwardedMeetingNotifications', 'AddNewRequestsTentatively', 'ProcessExternalMeetingMessages' ;
                $prpsCPNotif = 'ForwardRequestsToDelegates', 'AddAdditionalResponse', @{Name = 'AdditionalResponse'; Expression = { $_.AdditionalResponse | out-string } } ;
                $prpsCPKeyBook = 'AutomateProcessing', 'TentativePendingApproval' ;
                $prpsCPOpen = 'AllBookInPolicy', 'AllRequestInPolicy', 'AllRequestOutOfPolicy';
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnrS)" ;
                $smsg = "`n`n--STANDARD POLICY SETTINGS ($($MBX.userprincipalname )):`n$(($oReport.CalendarProcessing | fl $prpsCPPol | out-string).trim())" ;
                $smsg += "`n`n--NOTIFICATION SETTINGS ($($MBX.userprincipalname )):`n$(($oReport.CalendarProcessing | select $prpsCPNotif | out-string).trim())" ;
                $smsg += "`n`n--KEY BOOKING SETTINGS ($($MBX.userprincipalname )):`n$(($oReport.CalendarProcessing | fl $prpsCPKeyBook | out-string).trim())";
                $smsg += "`n`n*== Access Restrictions ==*:" ;
                $smsg += "`n`n--ResourceDelegates:`n$(($oResourceDelegates|out-string).trim())" ;
                $smsg += "`n`n--BookInPolicy:`n$(($oBookInPolicy|out-string).trim())" ;
                $smsg += "`n`n--RequestInPolicy:`n$(($oRequestInPolicy|out-string).trim())" ;
                $smsg += "`n`n--RequestOutOfPolicy:`n$(($oRequestOutOfPolicy|out-string).trim())" ;
                $smsg += "`n`n*== Open Access Settings ==*:`n$(($oReport.CalendarProcessing | fl $prpsCPOpen | out-string).trim())" ;
                write-host -f Gray $smsg ;
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnrS.replace('-v','-^').replace('v-','^-'))`n" ;
                write-host -f gray "==CALENDAR VIEW PERMISSIONS ($($MBX.userprincipalname )):`n" + ($oReport.CalendarFolderPermission | ft -a  FolderName, User, AccessRights | out-string ).trim() + "`n===========" ;
                write-host -foregroundcolor green "ACCESS SUMMARY:" ;
                if ($oReport.CalendarProcessing.AllBookInPolicy) {
                     write-host -fore yellow "AllBookInPolicy:OPEN BOOKING - All Users can book the room`n"
                } else {
                     write-host -fore yellow "AllBookInPolicy:`$false: End users CANNOT book the room (without Moderator Approval)`n"
                } ;
                if ($oReport.CalendarProcessing.AllBookInPolicy -eq $false -AND $oRequestInPolicy) {
                     write-host -fore yellow "RequestInPolicy:SOME specific users can REQUEST the room:`n$(($oRequestInPolicy|out-string).trim())`n"
                } ;
                if ($oReport.CalendarProcessing.AllBookInPolicy -eq $false -AND $oReport.CalendarProcessing.AllRequestInPolicy) {
                     write-host -fore yellow "AllRequestInPolicy:OPEN MODERATION REQUESTS - All Users can REQUEST the room (submits request for Delegates approval)`n"
                } else {
                     write-host -fore yellow "AllRequestInPolicy:`$false: End users CANNOT REQUEST the room `n"
                } ;
                if ($oReport.CalendarProcessing.resourcedelegates) {
                     write-host -fore yellow "ResourceDelegates:Room has designated Moderators:`n$(($oResourceDelegates|out-string).trim())`n"
                } else {
                     write-host -fore yellow "ResourceDelegates:Room has NO Moderators"
                } ;
                if ($oReport.CalendarProcessing.BookInPolicy) {
                    write-host -fore yellow "BookInPolicy:Room has EMPOWERED users that can AUTO-BOOK (bypassing Moderation):`n$(($oBookInPolicy|out-string).trim())`n"
                } else {
                    write-host -fore yellow "BookInPolicy:`$false:Room has NO RESTRICTED users that can auto-book"
                } ;
                write-verbose "returning summary object to pipeline" ; 
                [pscustomobject]$oReport | write-output ; 
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))`n" ;
                remove-alias ps1grcp ;
                remove-alias ps1gmbx ;
                remove-alias ps1gcp ;
                remove-alias ps1gmfp ;
            }
        }
    } ;
    #endregion INITIALIZE_RESOURCECALENDAR ; #*------^ END Initialize-ResourceCalendar ^------

    #region RESOLVE_CUSTOMDNAME ; #*------v Resolve-CustomDname v------
    function Resolve-CustomDname{
        <# CU5 custom brand handling, custom Dnames, ref for Company names (implemented in post set)
        To check for existing mbx, we need the final customized dname and we can't sub this completely, as the CU5 on newmbx is conditional
        and can't be assigned in this internal fct, so this is the logic for that later work, 
        duplicating it's code. [facepalm] 
        #>
        PARAM(
            [Parameter(Mandatory=$true,HelpMessage="Display Name for mailbox [fname lname,genericname]")]
                    [string]$DisplayName,
            [Parameter(Mandatory=$false,HelpMessage="Optionally force CU5 (variant domain assign) [-Cu5 Exmark]")]
                    [string]$Cu5,
            #[Parameter(Mandatory=$true,HelpMessage="Specify the userid to be responsible for access-grant-approvals[name,emailaddr,alias]")]
                    #[string]$Owner,
            [Parameter(Mandatory = $False, HelpMessage = "Mailbox type (Equipment|Room|Shared)[-Type Room]")]
                    [ValidateSet('Equipment', 'Room', 'Shared')]
                    [ValidateNotNullOrEmpty()]
                    [string]$Type = "ROOM"
            #[Parameter(HelpMessage="Optional parameter indicating new mailbox Is NonGeneric-type[-NonGeneric `$true]")]
                    #[bool]$NonGeneric,
            #[Parameter(HelpMessage="Optionally specify a 3-letter Site Code o force OU placement to vary from Owner's current site[3-letter Site code]")]
                    #[string]$SiteOverride
        )
        # CU5 custom brand handling, custom Dnames, ref for Company names (implemented in post set)
        if (-not $cu5 -OR ($cu5 -eq "toro")) {
        }elseif ($cu5 -AND ($cu5 -ne "toro")) {
            write-verbose -verbose:$true "CU5:$($cu5): CUSTOM DOMAIN SPEC" ;
            if($cu5 -match '^Americanaugers$'){
                switch($Type){
                    'Room'{  
                        if($displayname -notmatch "^AA\."){
                            $displayname = "AA.$($displayname)" ; 
                        } ; 
                    }
                    'Equipement'{
                        if($displayname -notmatch "^AA\."){
                            if($displayname -match 'Vehicle-'){
                                $displayname = "AA.$($displayname)" ; 
                            }elseif($displayname -match 'truck|\scar'){
                                $displayname = "AA.Vehicle-$($displayname)" ; 
                            }else{
                                $displayname = "AA.$($displayname)" ; 
                            }
                        }
                    }
                    'Shared'{
                        # no action
                    }
                }
            } ; 
            # do DIT custom prefixes for equipement & rooms
            if($cu5 -match '^Ditchwitch$'){
                switch($Type){
                    'Room'{  
                        if($displayname -notmatch "^r:\s"){
                            $displayname = "r: $($displayname)" ; 
                        } ; 
                    }
                    'Equipement'{
                        if(($displayname -notmatch "^r:\s") -AND ($displayname -notmatch "^v:\s")){
                            if($displayname -match 'vehicle|truck|\scar'){
                                $displayname = "v: $($displayname)" ; 
                            }else{
                                $displayname = "r: $($displayname)" ; 
                            }
                        }
                    }
                    'Shared'{
                        # no action
                    }
                }
            } ; 
            if($cu5 -match '^Exmark\.com$'){} ; 
            if($cu5 -match '^Hayter\.co\.uk$'){} ; 
            if($cu5 -match '^HHtrenchless$'){
                switch($Type){
                    'Room'{  
                        if($displayname -notmatch "^r:\sHH\s"){
                            $displayname = "r: HH $($displayname)" ; 
                        } ; 
                    }
                    'Equipement'{
                        if(($displayname -notmatch "^r:\sHH\s") -AND ($displayname -notmatch "^v:\sHH\s")){
                            if($displayname -match 'vehicle|truck|\scar'){
                                $displayname = "v: HH $($displayname)" ; 
                            }else{
                                $displayname = "r: HH $($displayname)" ; 
                            }
                        }
                    }
                    'Shared'{}
                }
            } ; 
            if($cu5 -match '^Irritrol\.com$'){} ; 
            if($cu5 -match '^Irritrol\.com$'){} ; 
            if($cu5 -match '^Perrot\.De$'){} ; 
            if($cu5 -match '^Perrot\.PL$'){} ; 
            if($cu5 -match '^Prokasrousa\.com$'){} ; 
            if($cu5 -match '^Radiushdd.com$'){
                switch($Type){
                    'Room'{  
                        if($displayname -notmatch "RAD\s"){
                            $displayname = "RAD $($displayname)" ; 
                        } ; 
                    }
                    'Equipement'{
                        if($displayname -notmatch "^RAD\s"){
                            $displayname = "RAD $($displayname)" ; 
                        }
                    }
                    'Shared'{}
                }
            } ; 
            if($cu5 -match '^Spartanmowers\.com$'){} ; 
            if($cu5 -match '^(Subsite\.com)$'){
                switch($Type){
                    'Room'{  
                        if($displayname -notmatch "^r:\sSubsite-"){
                            $displayname = "r: $($displayname)" ; 
                        } ; 
                    }
                    'Equipement'{
                        if(($displayname -notmatch "r:\sSubsite-") -AND ($displayname -notmatch "^v:\sSubsite-")){
                            if($displayname -match 'vehicle|truck|\scar'){
                                $displayname = "v: Subsite-$($displayname)" ; 
                            }else{
                                $displayname = "r: Subsite-$($displayname)" ; 
                            }
                        }
                    }
                    'Shared'{}
                }
            } ; 
            if($cu5 -match '^Spartanmowers\.com$'){} ; 
            if($cu5 -match '^Spartanmowers\.com$'){} ; 
            if($cu5 -match '^Toro\.be$'){} ; 
            if($cu5 -match '^Ventrac\.com$'){} ; 
        } ;
        write-verbose "Return updated Displayname to pipeline:$($displayname)" ; 
        $displayname | write-output ; 
    } ; 
    #endregion RESOLVE_CUSTOMDNAME ; #*------^ END Resolve-CustomDname ^------

    #region SHOW_MAILBOXINFO ; #*------v show-MailboxInfo v------
    function show-MailboxInfo {
        PARAM(
            [Parameter(Mandatory=$true,HelpMessage="Display Name for mailbox [fname lname,genericname]")]
                  [string]$Alias,
            [Parameter(HelpMessage = "Owner identifier[-owner fname.lname@domain.com]")]
                [string]$Owner,
            [Parameter(HelpMessage = "Optional domaincontroller (skips discovery)[-domaincontroller 'Dc1']")]
                [string]$domaincontroller
        )
        TRY{
            $mbxo = get-mailbox -Identity $Alias -domaincontroller $domaincontroller -EA STOP;
            if ( ($OwnerMbx = (get-mailbox -identity $($pltNmbx.Owner) -ea 0)) -OR ($OwnerMbx = (get-remotemailbox -identity $($pltNmbx.Owner) -ea 0)) ) {
            } ; 
            $cmbxo= $mbxo | Get-CASMailbox -domaincontroller $domaincontroller -EA STOP;
            $prpMbx = @{Name='LogonName';Expression={$_.SamAccountName }},'Name','DisplayName','Alias','database',
                'UserPrincipalName','Office','RetentionPolicy','CustomAttribute5','CustomAttribute9','RecipientType','RecipientTypeDetails'
            $aduprops="GivenName,Surname,Manager,Company,Office,Title,StreetAddress,City,StateOrProvince,c,co,countryCode,PostalCode,Phone,Fax,Description" ;
            $prpADU2 = 'GivenName','Surname','Manager','Company','Office','Title','StreetAddress','City','StateOrProvince','c','co','countryCode','PostalCode','Phone','Fax','Description' ; 
            $ADu = get-ADuser -Identity $mbxo.samaccountname -properties * -server $domaincontroller -EA STOP| Select-Object *;
    
            $mbxSummary = @"
===REVIEW SETTINGS:=== 

User Email:`t$(($mbxo.WindowsEmailAddress.tostring()).trim())
Owner Email:`t$(($OwnerMbx.WindowsEmailAddress.tostring()).trim())
Mailbox Information:
$(($mbxo | select $prpMbx | out-string).trim())

$(
    if($mbxo.ResourceCustom){
                "Specified/Default -ResourceCustom:$($mbxo.ResourceCustom)" | write-output 
    }
    if($mbxo.MailTip){
                "Specified -MailTip:$($mbxo.MailTip)" | write-output  
    }
    if($mbxo.ResourceCapacity){
                "Specified -ResourceCapacity:$($mbxo.ResourceCapacity)"  | write-output 
    }
    if($mbxo.AdditionalResponse -ANd $mbxo.AddAdditionalResponse){
                "Specified -AdditionalResponse:`n$($mbxo.AdditionalResponse)"  | write-output 
    }
)

$(($Adu | select $prpADU2 | out-string).trim())
$(
  if($NonGeneric -eq $false){
      "ActiveSyncMailboxPolicy:$($cmbxo.ActiveSyncMailboxPolicy.tostring())" ; write-output ; 
  } ;
) ; 
Description: $($Adu.Description.tostring())
Info: $($Adu.info.tostring())
$(
    if(!($MbxSplat.Shared -OR $MbxSplat.Room -OR $MbxSplat.Equipment  )){
        "Initial Password: $(($InputSplat.pass | out-string).trim())" | write-output 
    } ;
)

"@ ; 
            $mbxSummary | write-output ; 
        } CATCH {$ErrTrapd=$Error[0] ;
            write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
            write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
        } ;
    } ; 
    #endregion SHOW_MAILBOXINFO ; #*------^ END show-MailboxInfo ^------

    #region SHOW_MAILBOXPERMINFO ; #*------v show-MailboxPermInfo v------
    function show-MailboxPermInfo {    PARAM(
            [Parameter(Mandatory=$true,HelpMessage="Display Name for mailbox [fname lname,genericname]")]
                  [string]$Alias,
            [Parameter(HelpMessage = "Optional domaincontroller (skips discovery)[-domaincontroller 'Dc1']")]
                [string]$domaincontroller
        )
        TRY{
            $propsMbxP = 'user','AccessRights','IsInhertied','Deny' ; 
            $propsAMbxP = 'User','ExtendedRights','Inherited','Deny' ; 
            $mbxo = get-mailbox -Identity $Alias -domaincontroller $domaincontroller -EA STOP;
            # need the SG name to filter the perms
            if ($mbxo.DistinguishedName -match $rgxOUMigrations) {
                $SiteCode = [regex]::Match($mbxo.DistinguishedName, $rgxMigationsSite).groups[1].value ;
                if ($SiteCode) {
                    $smsg = "Resolved _MIGRATIONS tree SiteCode:$($SiteCode)" ;
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
                } else {
                    throw "Unable to resolve $($pltNmbx.Owner.DistinguishedName) into a SiteCode" ;
                    Break ;
                }
            } else {
                $SiteCode = $mbxo.identity.tostring().split('/')[1]
            } ;
            # build the sg name from dname & sitecode
            $sgName = "$($SiteCode)-SEC-Email-$($mbxo.DisplayName)-G"
            if($oSG = get-group -Identity $sgName -DomainController $domaincontroller){
                $rMbxP = get-mailboxpermission -identity $($mbxo.Identity) -user $oSG.samaccountname -domaincontroller $domaincontroller |
                    ?{$_.user -match ".*-(SEC|Data)-Email-.*$"} ; 
                $rAMbxP = Get-ADPermission -identity $($mbxo.Identity) -domaincontroller $domaincontroller -user $oSG.distinguishedName ; 
                $mbrs = Get-ADGroupMember -identity $oSG.distinguishedName -server $DomainController | 
                    Select-Object distinguishedName ;
            } ; 
    
            $mbxPermSummary = @"

===REVIEW SETTINGS:===`n----Updated Permissions:`n`nChecking Mailbox/AD Permission on $($mbxo.samaccountname) mailbox `n to accessing user:`n $($oSG.SamAccountName)`n---" ;

`n$(($rMbxP | format-list $propsMbxP |out-string).trim())

`n==User mbx grant: Confirming $($mbxo.name) member of $($grpN):

$(
    if ($mbxo.distinguishedname -match $rgxUserOUs) {
        "`n$(($rAMbxP|out-string | format-list $propsAMbxP ).trim())" | WRITE-OUTPUT 
    } else {
        "TMBX $($mbxo.samaccountname) is in a non-User OU: Term Hide/Unhide groups do not apply..." | write-output 
    }  ;
)

Updated $($oSG.Displayname) Membership...
$(
    if ($mbrs) {        
        "`n$(($mbrs |out-string).trim())`n-----------------------" | write-output ; 
    } else {
        $smsg = "(NO MEMBERS RETURNED)`n-----------------------" | write-output ;
    } ;
)

"@ ; 
            $mbxPermSummary | write-output ; 
        } CATCH {$ErrTrapd=$Error[0] ;
            write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
            write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
        } ;
    } ; 
    #endregion SHOW_MAILBOXPERMINFO ; #*------^ END show-MailboxPermInfo ^------

    #region SHOW_CALPROCINFO ; #*------v show-CalProcInfo v------
    function show-CalProcInfo {    
        PARAM(
            [Parameter(Mandatory=$true,HelpMessage="BookUpdate CalProc summary pscustomobject")]
                  $BookUpdate,
            [Parameter(HelpMessage = "Optional domaincontroller (skips discovery)[-domaincontroller 'Dc1']")]
                [string]$domaincontroller
        )
        PROCESS{
            foreach($BUpdate in $BookUpdate){
                TRY{
                    $OpRcp = get-recipient $BUpdate.Mailbox.userprincipalname  ;
                    $CPrcps = "ResourceDelegates", "BookInPolicy", "RequestInPolicy", "RequestOutOfPolicy";
                    $prpsMbx = @{Name = 'AcceptMessagesOnlyFrom'; Expression = { $_.AcceptMessagesOnlyFrom | select DistinguishedName | out-string } }, @{Name = 'AcceptMessagesOnlyFromDLMembers'; Expression = { $_.AcceptMessagesOnlyFromDLMembers } }, @{Name = 'AcceptMessagesOnlyFromSendersOrMembers'; Expression = { $_.AcceptMessagesOnlyFromSendersOrMembers | select DistinguishedName | out-string } }, 'Office';
                    $prpsCPPol = 'AllowConflicts', 'BookingWindowInDays', 'MaximumDurationInMinutes', 'AllowRecurringMeetings', 'EnforceSchedulingHorizon', 'ScheduleOnlyDuringWorkHours', 'ConflictPercentageAllowed', 'MaximumConflictInstances', 'ForwardRequestsToDelegates', 'DeleteAttachments', 'DeleteComments', 'RemovePrivateProperty', 'DeleteSubject', 'AddOrganizerToSubject', 'DeleteNonCalendarItems', 'EnableResponseDetails', 'OrganizerInfo', 'RemoveOldMeetingMessages', 'RemoveForwardedMeetingNotifications', 'AddNewRequestsTentatively', 'ProcessExternalMeetingMessages' ;
                    $prpsCPNotif = 'ForwardRequestsToDelegates', 'AddAdditionalResponse', @{Name = 'AdditionalResponse'; Expression = { $_.AdditionalResponse | out-string } } ;
                    $prpsCPKeyBook = 'AutomateProcessing', 'TentativePendingApproval' ;
                    $prpsCPOpen = 'AllBookInPolicy', 'AllRequestInPolicy', 'AllRequestOutOfPolicy';
                    #$oReport.Mailbox = ps1gmbx -identity $BUpdate.Mailbox.userprincipalname  ;
                    #$oReport.CalendarProcessing = ps1gcp -identity $BUpdate.Mailbox.userprincipalname  ;
                    #$oReport.CalendarFolderPermission = ps1gmfp -Identity "$($BUpdate.Mailbox.userprincipalname ):\Calendar" ;
                    if($BUpdate.Calendarprocessing.ResourceDelegates){ $oResourceDelegates = ($BUpdate.Calendarprocessing.ResourceDelegates | ps1grcp | select -expand primarysmtpaddress) -join ';'}
                    if($BUpdate.Calendarprocessing.BookInPolicy){ $oBookInPolicy = ($BUpdate.Calendarprocessing.BookInPolicy | ps1grcp | select -expand primarysmtpaddress) -join ';'}
                    if($BUpdate.Calendarprocessing.RequestInPolicy){ $oRequestInPolicy = ($BUpdate.Calendarprocessing.RequestInPolicy | ps1grcp | select -expand primarysmtpaddress) -join ';'}
                    if($BUpdate.Calendarprocessing.RequestOutOfPolicy){ $oRequestOutOfPolicy = ($BUpdate.Calendarprocessing.RequestOutOfPolicy | ps1grcp | select -expand primarysmtpaddress) -join ';'}
        
                    $CalProcSummary = @"

$(
    if ($BUpdate.Calendarprocessing.AllBookInPolicy) {
        "The Resource Calendar has default, 'anyone can book', access:No moderators, or booking restrictions." | write-output 
    }
    if($BUpdate.Calendarprocessing.BookInPolicy){
        "The Resource Calendar has access booking restrictions: `nUsers on the BookInPolicy list (below) can always book the room" | write-output 
    }
    if($BUpdate.Calendarprocessing.ResourceDelegates){
        "The Resource Calendar has Delegated/Moderated booking restrictions: `nUser can request bookings (which are initially marked 'Tentative'),`nand the ResourceDelegates (below) will receive approval email to approve these requests`n(at which time the booking is completed)" | write-output 
    }
    if($BUpdate.Calendarprocessing.RequestOutOfPolicy){ 
        "The Resource Calendar has RequestOutOfPolicy users:`nThese users can book meetings, beyond normal Calendar working hours" | write-output 
    }
)

#*======v $($BUpdate.Mailbox.userprincipalname ) v======

#*------v ==MBX DELIVERY RESTRICTIONS ($($BUpdate.Mailbox.userprincipalname )) v------
$(($BUpdate.Mailbox | fl  $prpsMbx |out-string).trim())
#*------v ==CALENDAR SETTINGS ($($BUpdate.Mailbox.userprincipalname )): v------" ;

--STANDARD POLICY SETTINGS ($($BUpdate.Mailbox.userprincipalname )):`n$(($BUpdate.Calendarprocessing | fl $prpsCPPol | out-string).trim())

--NOTIFICATION SETTINGS ($($BUpdate.Mailbox.userprincipalname )):`n$(($BUpdate.Calendarprocessing | select $prpsCPNotif | out-string).trim())

--KEY BOOKING SETTINGS ($($BUpdate.Mailbox.userprincipalname )):`n$(($BUpdate.Calendarprocessing | fl $prpsCPKeyBook | out-string).trim())

*== Access Restrictions ==*
--ResourceDelegates:`n$(($oResourceDelegates|out-string).trim())
--BookInPolicy:`n$(($oBookInPolicy|out-string).trim())
--RequestInPolicy:`n$(($oRequestInPolicy|out-string).trim())
--RequestOutOfPolicy:`n$(($oRequestOutOfPolicy|out-string).trim())
*== Open Access Settings ==*:`n$(($BUpdate.Calendarprocessing | fl $prpsCPOpen | out-string).trim())

==CALENDAR VIEW PERMISSIONS ($($BUpdate.Mailbox.userprincipalname )):
$($BUpdate.CalendarFolderPermission | ft -a  FolderName, User, AccessRights | out-string ).trim() 
===========

ACCESS SUMMARY:
$(
    if ($BUpdate.Calendarprocessing.AllBookInPolicy) {
         "AllBookInPolicy:OPEN BOOKING - All Users can book the room`n" | write-output 
    } else {
         "AllBookInPolicy:`$false: End users CANNOT book the room (without Moderator Approval)`n"  | write-output 
    } ;
    if ($BUpdate.Calendarprocessing.AllBookInPolicy -eq $false -AND $oRequestInPolicy) {
         "RequestInPolicy:SOME specific users can REQUEST the room:`n$(($oRequestInPolicy|out-string).trim())`n" | write-output 
    } ;
    if ($BUpdate.Calendarprocessing.AllBookInPolicy -eq $false -AND $BUpdate.Calendarprocessing.AllRequestInPolicy) {
         "AllRequestInPolicy:OPEN MODERATION REQUESTS - All Users can REQUEST the room (submits request for Delegates approval)`n" | write-output 
    } else {
         "AllRequestInPolicy:`$false: End users CANNOT REQUEST the room `n" | write-output 
    } ;
    if ($BUpdate.Calendarprocessing.resourcedelegates) {
         "ResourceDelegates:Room has designated Moderators:`n$(($oResourceDelegates|out-string).trim())`n" | write-output 
    } else {
         "ResourceDelegates:Room has NO Moderators" | write-output 
    } ;
    if ($BUpdate.Calendarprocessing.BookInPolicy) {
        "BookInPolicy:Room has EMPOWERED users that can AUTO-BOOK (bypassing Moderation):`n$(($oBookInPolicy|out-string).trim())`n" | write-output 
    } else {
        "BookInPolicy:`$false:Room has NO RESTRICTED users that can auto-book" | write-output 
    } ;
) ; 
"@ ; 
                    $mbxPermSummary | write-output ; 
                } CATCH {$ErrTrapd=$Error[0] ;
                    write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                } ;
            } ; # loop-e bookupdate objectS
        }  # PROC-E
    }
    #endregion SHOW_CALPROCINFO ; #*------^ END show-CalProcInfo ^------


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
    $moveTargets = @() ;    
    $cultTxt = $((Get-Culture).TextInfo) ;
    $ttl = $mbxSpecs.count ;

    [string[]]$WorkSummary = @() ; 
    $wsOpen = @"

Per your request:

SUBJECT
SYMPTOM


"@ ; 
$WorkSummary += $wsOpen ; 
    

    # EVAL MIXED $MBXSPECS AND -MailTip:$($Mailtip) -ResourceCustom:$($ResourceCustom) -ResourceCapacity:$($ResourceCapacity)
    if( ( ($mbxSpecs | measure).count -gt 1) -AND ($Mailtip -OR $ResourceCustom -OR $ResourceCapacity -OR $Location -OR $AdditionalResponse)){
        $smsg = "-MBXSPECS WITH MULTIPLE MAILBOXES SPECIFIFIED" ;   
        $smsg += "`nWITH ONE OR MORE POST-CREATE PARAMETERS:" ;
        $smsg += "`n-MailTip:$($Mailtip) -ResourceCustom:$($ResourceCustom) -ResourceCapacity:$($ResourceCapacity)`n-Location:$($Location) `n-AdditionalResponse:`$($AdditionalResponse)" ;
        $smsg += "`nTHOSE PARAMETERSE ARE APPLIED TO *ALL* MAILBOXES PROCESSED!" ;
        $smsg += "`nPLEASE CONFIRM YOU WANT TO MOVE AHEAD CREATING MULTIPLE RESOURCE MAILBOXES WITH THE POST-PARAMETERS SPECIFIED ON EACH!" ;
        write-warning $SMSG ; 
        $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ;
        if ($bRet.ToUpper() -eq "YYY") {
            $smsg = "(Moving on)" ;
            write-host -foregroundcolor green $smsg  ;
        } else {
            $smsg = "(ABORT)" ;
            write-host -foregroundcolor yellow $smsg  ;
            break ; #exit 1
        } ; 
        
    }
    #endregion SUBMAIN ; #*======^ END SUB MAIN ^======
} ;  # BEGIN-E
PROCESS {
    $Error.Clear() ;
    $Procd = 0 ;
    if (-not $mbxSpecs -AND $MbxsFormat -AND $Owner -AND $Grantees){
   
            write-verbose "Running Parameters build"
            # just roll a mbxspecs from the params, and do that
            if($MbxsFormat -eq 'DNAME'){
                $mbxSpecs = @($DisplayName, $cu5, $Owner, $Grantees) -join ';'
            }elseif($MbxsFormat -eq 'EMAIL'){
                $mbxSpecs = @($EmailAddress, $cu5, $Owner, $Grantees) -join ';'
            } ; 
            $smsg = "Resolved params into MbxSpecs:"
            $smsg += "`n$(($mbxSpecs|out-string).trim())" ; 
            write-verbose $smsg ; 
    } ; 
    if ($mbxSpecs){
        #$MbxsFormat = 'Dname', Email
        # resource:Type:Room|Equipment $mbxspecs are assumed to have 4 elements and be semicolon delimited
        if ($mbxSpecs -AND ($mbxspecs.tochararray() -contains ';') -AND ($mbxSpecs.split(';').count -eq 4)) {
            foreach ($mbx in $mbxSpecs) {
                if($MbxsFormat -eq 'DNAME'){
                    #[array]$mbxSpecs = "DISPLAYNAME1;TORO; OWNER1;GRANTEE1A@toro.com,GRANTEE1B@toro.com" ; $mbxSpecs += "DISPLAYNAME2;TORO;OWNER2;GRANTEE2A@toro.com,GRANTEE2B@toro.com" ;
                    $DisplayName, $cu5, $Owner, $Grantees = $mbx.split(';').trim() ;
                }elseif($MbxsFormat -eq 'EMAIL'){
                    $EmailAddress, $cu5, $Owner, $Grantees = $mbx.split(';').trim() ;
                    write-verbose "detect periods, go into email address" ;
                    if ($EmailAddress.split('@')[0].contains('.')) {
                        if($EmailAddress.split('@')[0].split('.').count -gt 2){
                            $smsg = "MORE THAN A SINGLE PERIOD IN THE DNAME! PERIOD IS USED IN NEW-MAILBOXSHARE() TO SPLIT FNAME/LNAME!" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            THROW $SMSG ; 
                            RETURN ; 
                        }
                        write-host "Dname contains periods, preserving for inclusion in new-MailboxShared as period-delimtied email addr" ;
                        write-verbose "if all caps, split on @, replace any ! with spaces" ;
                        #if ($EmailAddress -cmatch '([A-Z])') { 
                        # fix, *actual* all caps rgx                        
                        if ($EmailAddress -cmatch '^[A-Z]+$') { 
                            # all caps: indicates replacable delimiters, will ned TitleCase post to get dname
                            #$DisplayName = $EmailAddress.split("@")[0].replace('!', ' ') 
                            # swap out .? ; no pass it through, and purge it in new-mailboxshared() as a way to fname/lname split the dname
                            $DisplayName = $cultTxt.ToTitleCase($EmailAddress.split("@")[0].replace('!', ' ').replace("_", " ").toLower())
                        } else { 
                            $DisplayName = $EmailAddress.split("@")[0].replace("_", " ")
                        } ;
                    } else {
                        if ($EmailAddress -cmatch '^[A-Z]+$') { 
                            # all caps, sub !/_ -> \s, Titlecase
                            $DisplayName = $cultTxt.ToTitleCase($EmailAddress.split("@")[0].replace('!', ' ').replace("_", " ").toLower())
                        } else {
                            #$DisplayName = $cultTxt.ToTitleCase(($EmailAddress.split("@")[0].replace("_", " ").replace(".", " ").toLower()))  ; 
                            # pass through periods, new-sharedmailbox will build fname/lname out of the period loc
                            $DisplayName = $cultTxt.ToTitleCase(($EmailAddress.split("@")[0].replace("_", " ").toLower()))  ; 
                        } ;
                    } ;
                } ; 
                $Procd++ ;
                $smsg = $sBnrS = "`n#*------v PROCESSING ($($Procd)/$($ttl)) : $($DisplayName) v------" ;
                write-host -foregroundcolor CYAN "$((get-date).ToString('HH:mm:ss')):$($sBnrS)" ;
                $pltNmbx = [ordered]@{  
                    ticket = $ticket ;
                    DisplayName                = "$($DisplayName)"  ;
                    MInitial                   = "" ;
                    Owner                      = $Owner ;
                    Cu5                        = $Cu5
                    SiteOverride               = $SiteOverride ;
                    NonGeneric                 = $false  ;
                    Type                       = $Type ; 
                    Vscan                      = "YES" ;
                    NoPrompt                   = $true ;
                    domaincontroller           = $domaincontroller ;
                    Grantees                   = $Grantees ; 
                    #showDebug                 = $true ;
                    Verbose                    = $($VerbosePreference -eq 'Continue')
                    whatIf                     = $($whatif) ;
                } ;
                $smsg = "-mbxSpecs: $($Procd)/$($ttl): resolved specifications:" ; 
                $smsg += "`nCreateGrant-Mailbox w`n$(($pltNmbx|out-string).trim())" ;
                $smsg += "`n`n(`$Grantees (applied later):$(( $grantees -join ', ' |out-string).trim()))`n`n" ; 
                write-host -foregroundcolor yellow $smsg ; 
                $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ;
                if ($bRet.ToUpper() -eq "YYY") {
                    $smsg = "(Moving on)" ;
                    write-host -foregroundcolor green $smsg  ;
                } else {
                    $smsg = "(ABORT!)" ;
                    write-host -foregroundcolor yellow $smsg  ;
                    return ; #break ; #exit 1
                } ; 
                $premoveT = $moveTargets |  measure | select -expand count  ;                 
                # we need the customized dname before we can check for existing
                $tDname = Resolve-CustomDname -DisplayName $DisplayName -Cu5 $Cu5 -Type $Type #-Owner $Owner -SiteOverride $SiteOverride ;                   
                $pltSMRD = [ordered]@{
                    TargetMailbox = $null  ;
                    Type = $Type ; 
                    Location = $Location ;
                    MailTip = $Mailtip ;
                    ResourceCustom = $ResourceCustom ;
                    ResourceCapacity = $ResourceCapacity ;
                    Verbose = $($VerbosePreference -eq 'Continue')
                    whatIf = $whatif ;
                }
                if($xistMbx = get-mailbox $tDname -domaincontroller $domaincontroller -ea 0){
                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):Mailbox with DisplayName:$($tDname) already exists! Skipping mailbox creation, attempting to continue with Booking Config (skipping creation & grants)..." ;
                    #$moveTargets += $xistMbx.alias ;  
                    $moveTargets += $xistMbx.userprincipalname ;
                    if(($moveTargets| measure-object).count -eq $premoveT){
                        $smsg = "Update fail: " ; 
                        write-warning $smsg ;
                        return ; 
                    } ;                     
                }else{
                    # had pipeline blown, return should be the [string]Alias, filter anything else out on return
                    # 11:47 AM 4/23/2026 it's now returning the UPN, v the alias (which was borked for get-xorecipient replic test)
                    $moveTargets += CreateGrant-Mailbox @pltNmbx | 
                        ?{$_.gettype().fullname -ne 'System.Management.Automation.PSObject' -AND $_.gettype().fullname -eq 'System.String'}
                    if(($moveTargets| measure-object).count -eq $premoveT){
                        if(-not $whatif){
                            $smsg = "Update fail: " ; 
                            write-warning $smsg ;
                            return ; 
                        }else{
                            $smsg = "-whatif: Skipping Post Updates: " ; 
                            write-host -foregroundcolor green $smsg ; 
                            return ; 
                        }
                    } ; 
                    Do {
                        write-host "." -NoNewLine;
                        start-sleep 5 ;                             
                    } Until ($xistMbx = get-mailbox -id $movetargets[-1] -domaincontroller $domaincontroller ) ;                    
                }
                if(-not $whatif){                    
                    $tDname = $xistMbx.displayname
                    if($EmailAddress -AND ($xistmbx.PrimarySmtpAddress -ne $EmailAddress)){
                        $smsg = "-EmailAddress:GENERATED PrimarySmtpAddress $($xistmbx.PrimarySmtpAddress) *NOTEQUALS* SPECIFIED `$EmailAddress $($EmailAddress)" ; 
                        $smsg += "`nFORCING OUT OF EAC-MGMT WITH CUSTOM VALUE, TO ENFORCE ALIGNMENT!" ;
                        write-warning $smsg ; 
                        #region PSPLT3T ; #*------v (psb-PSPLT3T.cbp) v------
                        # FIND/REPL: SMbxA IDENTITY set-mailbox 
                        $whatif = $true ;
                        $pltSMbxA=[ordered]@{
                            IDENTITY = $xistMbx.ExchangeGuid  ;
                            EmailAddressPolicyEnabled = $false ;
                            PrimarySmtpAddress = $EmailAddress ; 
                            domaincontoller = $domaincontroller ; 
                            erroraction = 'STOP' ; 
                            Verbose = $($VerbosePreference -eq 'Continue')
                            whatif = $($whatif) ;
                        } ;
                        $thisDirName,$thisDomain = $EmailAddress.split('@') ; 
                        try {$conflict = $null ; $conflict = get-recipient -id (@($dirname,$domainname) -join '@') -domaincontroller $DomainController -ea 0 } CATCH {} ;
                        $incr = 1 ;  
                        Do{
                            if($conflict){
                                $incr++ ; 
                                $pltSMbxA.PrimarySmtpAddress = "$($dirname)$($incr)@($($domainname)" ; 
                            }
                            try {$conflict = $null ; $conflict = get-recipient -id $pltSMbxA.PrimarySmtpAddress -domaincontroller $DomainController -ea 0} CATCH {} ;
                        }while($conflict) ; 
                        write-verbose "calculate propertries for post query, from splat" ; 
                        $rgxExclParams = 'Verbose|Debug|ErrorAction|WarningAction|ErrorVariable|WarningVariable|OutVariable|OutBuffer|whatif|Confirm|force' ; 
                        [string[]]$prpSMbxA = @('UserprincipalName','displayname','Name','Alias') ;
                        [string[]]$prpSMbxA = $(@($prpSMbxA);@($pltSMbxA.GetEnumerator().name | ?{$_ -notmatch $rgxExclParams})) ; 
                        $prpSMbxA += 'WindowsEmailAddress' ; 
                        write-verbose "(Purge no value keys from splat, and conditionally pass through whatifpref)" ; 
                        $mts = $pltSMbxA.GetEnumerator() |?{ ($_.value -eq $null) -OR ($_.value -eq '') -OR ($_.value.length -eq 0)} ; $mts |foreach-object{$pltSMbxA.remove($_.Name)} ; remove-variable mts -ea 0 ; 
                        $smsg = "set-mailbox w`n$(($pltSMbxA|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;                        
                        TRY{
                            set-mailbox @pltSMbxA ;
                            $smsg = "`n`n==>POST:`n$((GET-mailbox -Identity $pltSMbxA.identity| fl $prpSMbxA| out-string).trim())`n`n" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } CATCH {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        } ;
                        #endregion PSPLT3T ; #*------^ END (psb-PSPLT3T.cbp) ^------
                   
                    }elseif($EmailAddress -AND ($xistmbx.PrimarySmtpAddress -eq $EmailAddress)){
                        $smsg = "(-EmailAddress:Generated PrimarySmtpAddress $($xistmbx.PrimarySmtpAddress) EQUALS specified `$EmailAddress $($EmailAddress))" ; 
                        write-host -foregroundcolor green $smsg ; 
                    }else{
                        write-verbose "(no -EmailAddress specification; no action)"
                    }
                    if($Type -match 'Room|Equipement'){
                        $pltSMRD.TargetMailbox = $xistMbx ;
                        #$xistmailbox = Set-MailboxRoomDetails -TargetMailbox $xistMbx -Location:$($Location) -MailTip:$($Mailtip) -ResourceCustom:$($ResourceCustom) -ResourceCapacity:$($ResourceCapacity) -whatIf:$($whatif) ; 
                        write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):Set-MailboxRoomDetails w`n$(($pltSMRD|out-string).trim())" ; 
                        $xistmailbox = Set-MailboxRoomDetails @pltSMRD ; 
                        $WorkSummary += "I have created a new conference ROOM/EQUIPMENT calendar mailbox with the following specifications:`n`n"
                    }else{
                        $smsg = "-Type:$($Type): RoomUpdates don't apply (Set-MailboxRoomDetails): " ; 
                        write-host -foregroundcolor green $smsg ;     
                        $WorkSummary += "I have created the requested SHARED mailbox with the following specifications:`n`n"
                    }
                    $WorkSummary += show-MailboxInfo -Alias $xistMbx.Alias -domaincontroller $domaincontroller; 
                    $WorkSummary += "`n`nThe Mailbox Access is configured with the following permissions grants:`n`n"
                    $WorkSummary += show-MailboxPermInfo -Alias $xistMbx.Alias -domaincontroller $domaincontroller; 
                    
                }else{
                    $smsg = "-whatif: Skipping Post Updates: " ; 
                    write-host -foregroundcolor green $smsg ; 
                    return ; 
                }
            } # loop-E
        }else{
                write-warning "$((get-date).ToString('HH:mm:ss')):Invalid mbx spec format! Each mbx spec should be in the format: `nDISPLAYNAME;CU5;OWNER;GRANTEE1,GRANTEE2,...`nExample:`n`"Conference Room 1;TORO;John Doe;"
        }    
    } else{
        $smsg = "EITHER a -mbxspecs array of inputs, or a minimum combination of: -DisplayName -AND -Owner -AND -Grantees must be specified to run this script!" ; 
        $smsg += "`n(ABORTING)" ;
        write-warning $SMSG ; 
        throw $SMSG ; 
        RETURN ; 
    }# if-E # $mbxSpecs ARRAY VS EXPLICIT PARAMS
    if(-not $whatif) {
        write-host -foregroundcolor green "===$((get-date).ToString('HH:mm:ss')):CONFIRMING PERMISSIONS:" ;
        #foreach ($mbx in $mbxSpecs) {
        # no the dname is uncustomized, not accurate, loop the $movetargets (aliases) instead
        $mbxArray = @() ; 
        foreach ($alias in $movetargets) {
            try{$currMbx = get-mailbox -id $alias -ea STOP ; $mbxArray += $currMbx} CATCH {
                $ErrTrapd=$Error[0] ;
               write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
               $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
               write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
             } ;
            #$DisplayName =$mbx.split(";")[0]
            #$DisplayName = $DisplayName.replace('_', ' ').replace('.', ' ').toLower()
            #$DisplayName = $cultTxt.ToTitleCase($DisplayName) ;
            #$smsg = "`$DisplayName:$($DisplayName)" ;
            $smsg = "`$DisplayName:$($currMbx.DisplayName)" ;
            write-host -foregroundcolor green $smsg
            #$mbp = get-mailboxpermission -identity "$($DisplayName)" | ? { $_.user -like 'toro*' } | select user;
            $mbp = $currMbx | get-mailboxpermission | ? { $_.user -like 'toro*' } | select user;
            $smsg = "$(($mbp  |out-string).trim())" ;
            write-host -foregroundcolor green $smsg ;
        } ;
    } ;
    # update room settings
    if(-not $whatif -AND $mbxArray ){
        if((($mbxARray.recipienttypedetails -replace 'Mailbox','') -match 'Room|Equipment') -ANd -not $skipBookingConfig){
            $pltInRC=[ordered]@{
                Ticket = $Ticket ;
                TargetMailbox = $mbxArray   ;
                BookingType = $BookingType  ;
                BookinPolicy = $BookinPolicy  ;
                ResourceDelegates = $ResourceDelegates  ;
                Grantees = $Grantees  ;
                Verbose = $($VerbosePreference -eq 'Continue')
                whatIf = $($whatif) ; 
            } ; 
            $smsg = "Initialize-ResourceCalendar w`n$(($pltInRC|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $BookUpdate = Initialize-ResourceCalendar @pltInRC ; 
            $WorkSummary += show-CalProcInfo -BookUpdate $BookUpdate -domaincontroller $domaincontroller; ; 
        }elseif((($mbxARray.recipienttypedetails -replace 'Mailbox','') -notmatch 'Room|Equipment')){
            WRITE-VERBOSE "SHARED/GENERIC mailbox, no Set-CalendarProcessing applies)" ;
        }elseif( (($mbxARray.recipienttypedetails -replace 'Mailbox','') -match 'Room|Equipment')-ANd $skipBookingConfig){
            $sHRule6 = if($psise){'/\'*5}else{'/\'*[int]($Host.UI.RawUI.WindowSize.Width/3/2) }
            $whHRule6 =@{BackgroundColor = 'Red' ; ForegroundColor = 'Yellow' } ; 
            write-host $sHRule6 @whHRule6 ;
            $smsg = "`n`n-skipBookingConfig: SKIPPED RESOURCE BOOKING CONFIGURATION!`n`n" ;            
            write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" ; 
            write-host $sHRule6 @whHRule6 ;
        }else{
            write-warning "unrecognized combo: recipienttypedetails -notmatch 'Room|Equipment' -AND  -not `$skipBookingConfig!"
        } ; 
    } ; 
    if (-not $whatif -AND $alias) {
        #connect-mg ;
        write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):PREPARING DAWDLE LOOP!($($tmbx.PrimarySmtpAddress))`nMGOPLastSync:`n$((get-MGOPLastSync| ft -a TimeGMT,TimeLocal|out-string).trim())" ;
        Do {
            rxo ;
            write-host "." -NoNewLine;
            Start-Sleep -s 30        
        } Until ((get-xorecipient $moveTargets[-1] -EA 0)) ;
        write-host "`n*READY TO MOVE*!`a" ;
        start-sleep -s 1 ;
        write-host "*READY TO MOVE*!`a" ;
        start-sleep -s 1 ;
        write-host "*READY TO MOVE*!`a`n" ;
        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Running:`n`nmove-EXOmailboxNow.ps1 -TargetMailboxes $($sQot + ($moveTargets -join '","') + $sQot) -showDebug -whatIf`n`n" ;
        . move-EXOmailboxNow.ps1 -TargetMailboxes $moveTargets -showDebug:$($false) -whatIf ;
        if(-not $MoveImmediate){            
            $smsg = "-whatif pass completed ^ above ^`nDO YOU WANT TO MOVE AHEAD WITH immediate CLOUD MIGRATION?" ; 
            $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ;
        } ; 
        if ( ($bRet.ToUpper() -eq "YYY" ) -OR $MoveImmediate) {
            $smsg = "(Moving on)" ;
            write-host -foregroundcolor green $smsg  ;
            $postSTat = . move-EXOmailboxNow.ps1 -TargetMailboxes $moveTargets -showDebug:$($false) -NoTEST -whatIf:$false -outputReport:$true ;
            if($postStat.BatchName){
                    
                $smsg= "`nContinue monitoring with:`n get-xomoverequest -BatchName $($postStat.BatchName) | Get-xoMoveRequestStatistics | fl DisplayName,status,percentcomplete,itemstransferred,BadItemsEncountered`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            }

        } else {
            $smsg = "(*skip*)" ;
            write-host -foregroundcolor yellow $smsg  ;
            $strMoveCmd = ". move-EXOmailboxNow.ps1 -TargetMailboxes `$moveTargets -showDebug -NoTEST  -whatIf:`$false" ;
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Move Command (copied to cb):`n`n$($strMoveCmd)`n" ;
            $strMoveCmd | out-clipboard ;
            return ; #exit 1
        } ;             
        $strCleanCmd = "get-xomoverequest -BatchName ExoMoves-* | ?{`$_.status -eq 'Completed'} | Remove-xoMoveRequest -whatif" ;
        write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):Post-completion Cleanup Command :`n`n$($strCleanCmd)`n" ;
        if($SID = get-user -id $env:username | ?{$_.name -notmatch 'ExchangeAdmin'}){
            # returns Mailuser|Rmbx, and User|User, sort rtd and select 1st (to aim for the rmbx): otherwise, it returns 2 names 
            $UID = get-user -id ($sid.name.replace('S-','')) | sort RecipientTypeDetails | select -first 1 ; 
        } ;
        $hsWorkClose += @"

Finally, I initiated a move of the mailbox to the cloud.

$(
    if($postStat){
        "$(get-date -format 'HH:mm:tt'): INFO: CLOUD MIGRATION STATUS:" | write-output ;
        "$(($PostStat | fl 'DisplayName','Status','PercentComplete','ItemsTransferred','BadItemsEncountered' |out-string).trim())" | write-output ;
    }
)

Please let me know if you have any other questions, or I can be of further service.

Thanks,
$(
    $sig = @() ; 
    #if($UID){ $sig += "$($UID.name)`n" | write-output } ; 
    if($UID){ $sig += $UID.name}else{$sig = "`n`n"} ; 
    if($uid.title -AND $uid.title -match  ','){
        $rname = $uid.title.split(',') ; 
        [array]::Reverse($rname) ; 
        #(($rname -join ' ' | write-output).trim()).trim() | write-output ; 
        $sig+=(($rname -join ' ').trim()).trim()  ; 
    }elseif($uid.title){
        #($uid.title  | write-output).trim() | write-output ; 
        $sig+= $uid.title ; 
    }else{
        #"`n`n" | write-output  ; 
        $sig = "`n`n"
    }
    if($sig){$sig -join "`n" | write-output }
)
Toro Messaging Team

"@ ;
        $WorkSummary += $hsWorkClose ; 

        $smsg = $sBnr="#*======v CLOSE SUMMARY: v======" ; 
        $smsg += $WorkSummary ; 
        $smsg += $sBnr.replace('=v','=^').replace('v=','^=')
        write-host $smsg ; 
    } else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(WHATIF: skipping Add-MailboxPermission, Set-CalendarProcessing & Cloud Migration)" } ;
} # PROC-E
END{

}