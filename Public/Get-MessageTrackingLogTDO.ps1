# Get-MessageTrackingLogTDO.ps1

# To convert vx10 func to .ps1 script, for vx10-less use: rem the function declare lines, and saveas Get-MessageTrackingLogTDO.ps1, update exec splats to use .\Get-MessageTrackingLogTDO.ps1 vs unpathed/ext func name
#*------v Function Get-MessageTrackingLogTDO v------
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
    Github      : htt-ps://github.com/tostka/verb-XXX
    Tags        : Powershell,Exchange,MessageTracking,Get-MessageTrackingLog,ActiveDirectory
    REVISIONS
    * 10;23 am 4/30/2025 - get-ADExchangeServerTDO: * 10;05 am 4/30/2025 fixed code for Edge role in raw PS, missing evaris for Ex: added discovery from reg & stock file system dirs for version etc.
        - Connect-ExchangeServerTDO: * 10;07 am 4/30/2025 fixed borked edge conn, typo, and rev logic for Ex & role detection in raw PS - lacks evaris for exchange (EMS/REMS only), so leverage reg & stock install loc hunting to discover setup.exe for vers & role confirm).
        - start-log: made Tag cleanup conditional (on funcs avail; legacy of bringing full start-log into the mix).
    * 1:35 PM 4/29/2025: added resolve-NetworkLocalTDO() (for vnet); 
        - Initialize-xopEventIDTable(): * 2;58 pm 4/28/2025 Updated table again, and found Ex2016/19 eventid specifications online, added. Did find that 
            the online doc doesn't document the edge SendExternal event id (added below, manually).      
        - ENCODED_CONTANTS block: added eval'd varis: $isLocalExchangeServer, $ByPassLocalExchangeServerTest, $IsEdgeTransport
        - new NETWORK_INFO block: $netsettings = resolve-NetworkLocalTDO() ;
        - new $isNonDomainServer test, $isLocalExchangeServer (transport svc & regkey), $IsEdgeTransport (regkey test)
        - revised $TenOrg forced resolution, to leverage DNS A fqdn suffix steering (collected via new resolve-NetworkLocalTDO())
        - UseOPAD: skip AD load, $useForestWide, $objforest, $forestdom, $domaincontroller & get-GCFast(), if $isNonDomainServer
        - added block that tests if($IsEdgeTransport -AND $psise) and warns likely AccessDenied's from ex commands (works fine in EMS outside of ise)
        - revised $OrgTag resolution, to detect Edge 'First Org' orgname, and substitute in $($TenOrg)edge log name tagging
        - stripped back Sender array expansion (ported from EXO, which supports mults): gmtl *doesnt* support [string[]]$Sender, throws error.
        - validated on TTC Ex10 (did CMW hubs & edge below)
    * 3:12 pm 4/28/2025 Noted, get-OrgConfig returns 'First Org... on edge, so coded in TenOrgedge as OrgTag in that niched case (was coming back 'O_FO' ; 
        extensive debugging updates in the services_control BP block, and detection and handling of TenOrg around non-dom-joined edge role machine; 
            Still having issues in psie debugging: throws localizaiton error - appears due to the EMS connection; works fine in raw EMS console on edge. 
            work either way on non-edge systems. Fine, but you can't step-debug it in ise on edge boxes
             - you can, but you can't get the get-messagetrackinglog command to exec properrly, throws Access Denied trying to access the logs
             errors correspond to those reported by folks trying to centralize edge tracking, from remote hosts. So it's definitely an REMS v EMS issue.
    * 4:25 PM 4/24/2025 validated TOR & full func verif CMW, after: Svcs_conn: added pretest for version in connect-exchangeserverTDO, splat (suppress dupe add error), only necc if both an explicit vers is specd & its on CMW
        - also added OrgName abbreviated to the exports, to differentiate sources
    * 8:01 PM 4/23/2025 add: Version, to match connect-ExchangeServerTDO() param, -DetailLimit for config of detailed outptut sizes; 
        # service_conn block: added $ExopVers, which defers to local $Version, if populated
        # also completely skips all the verb-x10 functions, when $TenOrg -eq CMW (no mod installs on their servers, this runs freestanding, self-contained only, no central va support
        - fixed bug/misplacement of most recent echo; restricted detail outputs to select -last $DetailLimit msgs; 
        - reworked Failed reporting; added an explicit 'other', for anything not covred by the existing targeted filtering.
        - substantially expanded Defer coverage: with volume defered in SAP tracks into other site, needed way to confirm retry/deliver
            coded in indexed hash of all SEND events, on msgid, with fast lookup of the Send for each Defer, and reporting/confirm on the exact amount of delay ea experienced (up to $detaillimit output).
        - fixed fundemental breaks in trailing echo blocks - weren't even functional, as last revised.
    * 5:43 PM 4/22/2025 CMW non-verb-ex2010: .ps1 version:
        - biggest issue, had to move func decl's to the top of BEGIN{ or they wouldn't register 
        - had to splice in additional functions, load-ADMS, get-GCFast, replaced the simplified write-log with the full version (w cbh & rem'd blocks pulled)
        - patch version in calls to Connect-ExchangeServerTDO (to 2010)

    * 4:51 PM 4/22/2025 added resolve-EnvironmentTDO alias: resolve-Environment; ADD: resolve-environment() & support, and updated start-log support; TLS_LATEST_FORCE ; missing regions; SWRITELOG ; SSTARTLOG ; 
        updated -Version supporting Connect-ExchangeServerTDO  ; convertFrom-MarkdownTable() to support... ; Initialize-xopEventIDTable; 
        fixed bug in -resultsize code; code to leverage Initialize-xopEventIDTable and output uniqued eventid's returnedon gmtl passes (doc output inline)
        copied over latest service conn code & slog for renv()
    * 3:23 PM 12/2/2024 throwing param transform errs on start & end (wyhen typed): pull typing, and do it post assignh, can't assign empty '' or $null to t a datetime coerced vary ;pre-reduce splat hash to populated values, in exmpl & BP use;
         rem out the parameterset code, and just do manual conflicting -start/-end -days tests and errors
    * 2:34 PM 11/26/2024 updated to latest 'Connect-ExchangeServerTDO()','get-ADExchangeServerTDO()', set to defer to existing
    * 4:20 PM 11/25/2024 updated from get-exomessagetraceexportedtdo(), more silent suppression, integrated dep-less ExOP conn supportadd delimters to echos, to space more, readability ;  fixed typo in eventid histo output
    * 3:16 PM 11/21/2024 working: added back Connectorid (postfiltered from results); add: $DaysLimit = 30 ; added: MsgsFail, MsgsDefer, MsgsFailRcpStat; 
    * 2:00 PM 11/20/2024 rounded out to iflv level, no dbg yet
    * 5:00 PM 10/14/2024 at this point roughed in ported updates from get-exomsgtracedetailed, no debugging/testing yet; updated params & cbh to semi- match vso\get-exomsgtracedetailed(); convert to a function (from ps1)
    * 11:30 AM 7/16/2024 CBH example typo fix
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

    NOTE: To port this into the no-verb-ex2010 module support in CMW, rem the function declare lines.

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
    .PARAMETER Version
    Specific Exchange Server Version to connect to('Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000')[-Version 'Ex2016']
    .PARAMETER Resultsize
    The ResultSize parameter specifies the maximum number of results to return. If you want to return all requests that match the query, use unlimited for the value of this parameter. The default value is 1000.
    .PARAMETER DetailLimit
    Integer number of maximum messages to be follow-up MessageTraceDetail'd (defaults to 20) [-DetailLimit 100]
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
    PS>     Version = 'Ex2010' ; 
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
    PS>     Version = 'Ex2010' ; 
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
    PS>     Version = 'Ex2016' ; 
    PS>     ResultSize="" ;
    PS>     Tag='' ;
    PS> } ;
    PS> $pltGMTL = [ordered]@{} ;
    PS> $pltI.GetEnumerator() | ?{ $_.value}  | ForEach-Object { $pltGMTL.Add($_.Key, $_.Value) } ;
    PS> $vn = (@("xopMsgs$($pltI.ticket)",$pltI.Tag) | ?{$_}) -join '_' write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):.\Get-MessageTrackingLogTDO w`n$(($pltGMTL|out-string).trim())`n(assign to `$$($vn))" ;
    PS> if(gv $vn -ea 0){rv $vn} ;
    PS> if($tmsgs = .\Get-MessageTrackingLogTDO @pltGMTL){sv -na $vn -va $tmsgs ;
    PS> write-host "(assigned to `$$vn)"} ;
    Demo with dot-exec'd script syntax, same as demo1, with .\'s and version ex16 (for renames from _func.ps1 -> .ps1, and rem'ing the function declar lines). 
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
        [Parameter(HelpMessage="Specific Exchange Server Version to connect to('Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000')[-Version 'Ex2016']")]
            [ValidateSet('Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000')]
            [string[]]$Version = 'Ex2010',
        [Parameter(Mandatory=$False,HelpMessage="The ResultSize parameter specifies the maximum number of results to return. If you want to return all requests that match the query, use unlimited for the value of this parameter. The default value is 1000.")]
            [ValidateScript({
              if( ($_ -match 'unlimited') -OR ($_.gettype().fullname -eq 'System.Int32') -OR ($null -eq $_) -OR ('' -eq $_) ){
                  $true ; 
              } else { 
                  throw "Resultsize must be an integer or the string 'unlimited' (or blank)" ; 
              } ;
            })]
            $ResultSize,
        [Parameter(HelpMessage="Integer number of maximum messages to be follow-up MessageTraceDetail'd (defaults to 20) [-DetailLimit 100]")]
            [int]$DetailLimit = 20,
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

        #region FUNCTIONS_INTERNAL ; #*======v FUNCTIONS_INTERNAL v======
        # Pull the CUser mod dir out of psmodpaths:
        #$CUModPath = $env:psmodulepath.split(';')|?{$_ -like '*\Users\*'} ;

        #region RESOLVE_ENVIRONMENTTDO ; #*------v resolve-EnvironmentTDO v------
        if(-not(gci function:resolve-EnvironmentTDO -ea 0)){
            #*----------v Function resolve-EnvironmentTDO() v----------
            function resolve-EnvironmentTDO {
                <#
                    .SYNOPSIS
                    resolve-EnvironmentTDO.ps1 - Resolves local environment into usable Script or Function-descriptive values (for reuse in logging and i/o access)
                    .NOTES
                    Version     : 0.0.2
                    Author      : Todd Kadrie
                    Website     : http://www.toddomation.com
                    Twitter     : @tostka / http://twitter.com/tostka
                    CreatedDate : 2025-04-04
                    FileName    : resolve-EnvironmentTDO.ps1
                    License     : (non asserted)
                    Copyright   : (non asserted)
                    Github      : https://github.com/tostka/verb-ex2010
                    Tags        : Powershell,ExchangeServer,Version
                    AddedCredit : theSysadminChannel
                    AddedWebsite: https://thesysadminchannel.com/get-exchange-cumulative-update-version-and-build-numbers-using-powershell/
                    AddedTwitter: URL
                    REVISION
                    * 4:13 PM 4/4/2025 init
                    .EXAMPLE
                    PS> write-verbose "Typically from the BEGIN{} block of an Advanced Function, or immediately after PARAM() block" ; 
                    PS> $Verbose = [boolean]($VerbosePreference -eq 'Continue') ;
                    PS> $rPSCmdlet = $PSCmdlet ;
                    PS> $rPSScriptRoot = $PSScriptRoot ;
                    PS> $rPSCommandPath = $PSCommandPath ;
                    PS> $rMyInvocation = $MyInvocation ;
                    PS> $rPSBoundParameters = $PSBoundParameters ;
                    PS> $pltRvEnv=[ordered]@{
                    PS>     PSCmdletproxy = $rPSCmdlet ;
                    PS>     PSScriptRootproxy = $rPSScriptRoot ;
                    PS>     PSCommandPathproxy = $rPSCommandPath ;
                    PS>     MyInvocationproxy = $rMyInvocation ;
                    PS>     PSBoundParametersproxy = $rPSBoundParameters
                    PS>     verbose = [boolean]($PSBoundParameters['Verbose'] -eq $true) ;
                    PS> } ;
                    PS> write-verbose "(Purge no value keys from splat)" ;
                    PS> $mts = $pltRVEnv.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltRVEnv.remove($_.Name)} ; rv mts -ea 0 ;
                    PS> $smsg = "resolve-EnvironmentTDO w`n$(($pltRVEnv|out-string).trim())" ;
                    PS> if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    PS> $rvEnv = resolve-EnvironmentTDO @pltRVEnv ;  
                    PS> write-host "Returned `$rvEnv:`n$(($rvEnv|out-string).trim())" ; 
                #>
                [Alias('resolve-Environment')]
                [CmdletBinding()]
                PARAM(
                    [Parameter(HelpMessage="Proxied Powershell Automatic Variable object that represents the cmdlet or advanced function that’s being run. (passed by external assignment to a variable, which is then passed to this function)")] 
                        $PSCmdletproxy,        
                    [Parameter(HelpMessage="Proxied Powershell Automatic Variable that contains the full path to the script that invoked the current command. The value of this property is populated only when the caller is a script. (passed by external assignment to a variable, which is then passed to this function).")] 
                        $PSScriptRootproxy,
                    [Parameter(HelpMessage="Proxied Powershell Automatic Variable that contains the full path and file name of the script that’s being run. This variable is valid in all scripts. (passed by external assignment to a variable, which is then passed to this function).")] 
                        $PSCommandPathproxy,
                    [Parameter(HelpMessage="Proxied Powershell Automatic Variable that contains information about the current command, such as the name, parameters, parameter values, and information about how the command was started, called, or invoked, such as the name of the script that called the current command. (passed by external assignment to a variable, which is then passed to this function).")]
                        $MyInvocationproxy,
                    [Parameter(HelpMessage="Proxied Powershell Automatic Variable that contains a dictionary of the parameters that are passed to a script or function and their current values. This variable has a value only in a scope where parameters are declared, such as a script or function. You can use it to display or change the current values of parameters or to pass parameter values to another script or function. (passed by external assignment to a variable, which is then passed to this function).")]
                        $PSBoundParametersproxy
                ) ; 
                BEGIN {
                    $Verbose = [boolean]($VerbosePreference -eq 'Continue') ; 
                    <#
                    $PSCmdletproxy = $PSCmdlet ; # an object that represents the cmdlet or advanced function that's being run. Available on functions w CmdletBinding (& $args will not be available). (Blank on non-CmdletBinding/Non-Adv funcs).
                    $PSScriptRootproxy = $PSScriptRoot ; # the full path of the executing script's parent directory., PS2: valid only in script modules (.psm1). PS3+:it's valid in all scripts. (Funcs: ParentDir of the file that hosts the func)
                    $PSCommandPathproxy = $PSCommandPath ; # the full path and filename of the script that's being run, or file hosting the funct. Valid in all scripts.
                    $MyInvocationproxy = $MyInvocation ; # populated only for scripts, function, and script blocks.
                    #>
                    # - $MyInvocation.MyCommand.Name returns name of a function, to identify the current command,  name of the current script (pop'd w func name, on Advfuncs)
                    # - Ps3+:$MyInvocation.PSScriptRoot : full path to the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
                    # - Ps3+:$MyInvocation.PSCommandPath : full path and filename of the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
                    #     ** note: above pair contain information about the _invoker or calling script_, not the current script
                    #$PSBoundParametersproxy = $PSBoundParameters ; 

                    if($host.version.major -ge 3){$hshOutput=[ordered]@{Dummy = $null ;} }
                    else {$hshOutput = New-Object Collections.Specialized.OrderedDictionary} ;
                    If($hshOutput.Contains("Dummy")){$hshOutput.remove("Dummy")} ;
                    $tv = 'PSCmdletproxy','PSScriptRootproxy','PSCommandPathproxy','MyInvocationproxy','PSBoundParametersproxy'
                    # stock the autovaris, if populated
                    $tv | % { 
                        $hshOutput.add($_, (get-variable -name $_ -ea 0).Value) 
                    } ;
                    write-verbose "`$hshOutputn$(($hshOutput|out-string).trim())" ; 
                    $fieldsnull = 'runSource','CmdletName','PSParameters','ParamsNonDefault' 
                    if([boolean]($hshOutput.MyInvocationproxy.MyCommand.commandtype -eq 'Function' -AND $hshOutput.MyInvocationproxy.MyCommand.Name)){
                        #$tv+= @('isFunc','funcname','isFuncAdv') ; 
                        $fieldsnull = $(@($fieldsnull);@(@('isFunc','funcname','isFuncAdv'))) ; 
                        #$tv+= @('FuncDir') ; 
                        $fieldsnull = $(@($fieldsnull);@(@('FuncDir'))) ; 
                    } ; 
                    if([boolean]($hshOutput.MyInvocationproxy.MyCommand.commandtype -eq 'ExternalScript' -OR $hshOutput.PSCmdletproxy.MyInvocation.InvocationName -match '\.ps1$')){
                        #$tv += @('isScript','ScriptName','ScriptBaseName','ScriptNameNoExt','ScriptDir','isScriptUnpathed') ; 
                        $fieldsnull = $(@($fieldsnull);@('isScript','ScriptName','ScriptBaseName','ScriptNameNoExt','ScriptDir','isScriptUnpathed')) ; 
                    } ; 
                    $tv = $(@($tv);@($fieldsnull)) ; 
                    # append resolved elements to the hash as $null 
                    $fieldsnull  | % { $hshOutput.add($_,$null) } ;
                    write-verbose "`$hshOutputn$(($hshOutput|out-string).trim())" ; 

                    if($hshOutput.isFunc = [boolean]($hshOutput.MyInvocationproxy.MyCommand.commandtype -eq 'Function' -AND $hshOutput.MyInvocationproxy.MyCommand.Name)){
                        $hshOutput.FuncName = $hshOutput.MyInvocationproxy.MyCommand.Name ; write-verbose "`$hshOutput.FuncName: $($hshOutput.FuncName)" ; 
                    } ;
                    if($hshOutput.isFunc -AND (gv PSCmdletproxy -ea 0).value -eq $null){
                        $hshOutput.isFuncAdv = $false 
                    }elseif($hshOutput.isFunc){
                        $hshOutput.isFuncAdv = [boolean]($hshOutput.isFunc -AND $hshOutput.PSCmdletproxy.MyInvocation.InvocationName -AND ($hshOutput.FuncName -eq $hshOutput.PSCmdletproxy.MyInvocation.InvocationName)) ; 
                    } ; 
                    if($hshOutput.isFunc -AND $hshOutput.PSScriptRootproxy){
                        $hshOutput.FuncDir = $hshOutput.PSScriptRootproxy ; 
                    } ; 
                    $hshOutput.isScript = [boolean]($hshOutput.MyInvocationproxy.MyCommand.commandtype -eq 'ExternalScript' -OR $hshOutput.PSCmdletproxy.MyInvocation.InvocationName -match '\.ps1$') ; 
                    $hshOutput.isScriptUnpathed = [boolean]($hshOutput.PSCmdletproxy.MyInvocation.InvocationName  -match '^\.') ; # dot-sourced invocation, no paths will be stored in `$MyInvocation objects 
                    [array]$score = @() ; 
                    if($hshOutput.PSCmdletproxy.MyInvocation.InvocationName){ 
                        # blank on basic funcs, popd on AdvFuncs
                        if($hshOutput.PSCmdletproxy.MyInvocation.InvocationName -match '\.ps1$'){$score+= 'ExternalScript' 
                        }elseif($hshOutput.PSCmdletproxy.MyInvocation.InvocationName  -match '^\.'){
                            write-warning "dot-sourced invocation detected!:$($hshOutput.PSCmdletproxy.MyInvocation.InvocationName)`n(will be unable to leverage script path etc from `$MyInvocation objects)" ; 
                            write-verbose "(dot sourcing is implicit script exec)" ; 
                            $score+= 'ExternalScript' ; 
                        } else {$score+= 'Function' }; # blank under function exec, has func name under AdvFuncs
                    } ; 
                    if($hshOutput.PSCmdletproxy.CommandRuntime){
                        # blank on nonAdvfuncs, 
                        if($hshOutput.PSCmdletproxy.CommandRuntime.tostring() -match '\.ps1$'){$score+= 'ExternalScript' } else {$score+= 'Function' } ; # blank under function exec, func name on AdvFuncs
                    } ; 
                    $score+= $hshOutput.MyInvocationproxy.MyCommand.commandtype.tostring() ; # returns 'Function' for basic & Adv funcs
                    $grpSrc = $score | group-object -NoElement | sort count ;
                    if( ($grpSrc |  measure | select -expand count) -gt 1){
                        write-warning  "$score mixed results:$(($grpSrc| ft -a count,name | out-string).trim())" ;
                        if($grpSrc[-1].count -eq $grpSrc[-2].count){
                            write-warning "Deadlocked non-majority results!" ;
                        } else {
                            $hshOutput.runSource = $grpSrc | select -last 1 | select -expand name ;
                        } ;
                    } else {
                        write-verbose "consistent results" ;
                        $hshOutput.runSource = $grpSrc | select -last 1 | select -expand name ;
                    };
                    if($hshOutput.runSource -eq 'Function'){
                        if($hshOutput.isFuncAdv){
                            $smsg = "Calculated `$hshOutput.runSource:Advanced $($hshOutput.runSource)"
                        } else { 
                            $smsg = "Calculated `$hshOutput.runSource: Basic $($hshOutput.runSource)"
                        } ; 
                    }elseif($hshOutput.runSource -eq 'ExternalScript'){
                        $smsg =  "Calculated `$hshOutput.runSource:$($hshOutput.runSource)" ;
                    } ; 
                    write-verbose $smsg ;
                    'score','grpSrc' | get-variable | remove-variable ; # cleanup temp varis
                    $hshOutput.CmdletName = $hshOutput.PSCmdletproxy.MyInvocation.MyCommand.Name ; # function self-name (equiv to script's: $MyInvocation.MyCommand.Path), pop'd on AdvFunc
                    #region PsParams ; #*------v PSPARAMS v------
                    $hshOutput.PSParameters = New-Object -TypeName PSObject -Property $hshOutput.PSBoundParametersproxy ;
                    # DIFFERENCES $hshOutput.PSParameters vs $PSBoundParameters:
                    # - $PSBoundParameters: System.Management.Automation.PSBoundParametersDictionary (native obj)
                    # test/access: ($PSBoundParameters['Verbose'] -eq $true) ; $PSBoundParameters.ContainsKey('Referrer') #hash syntax
                    # CAN use as a @PSBoundParameters splat to push through (make sure populated, can fail if wrong type of wrapping code)
                    # - $hshOutput.PSParameters: System.Management.Automation.PSCustomObject (created obj)
                    # test/access: ($hshOutput.PSParameters.verbose -eq $true) ; $hshOutput.PSParameters.psobject.Properties.name -contains 'SenderAddress' ; # cobj syntax
                    # CANNOT use as a @splat to push through (it's a cobj)
                    write-verbose "`$hshOutput.PSBoundParametersproxy:`n$(($hshOutput.PSBoundParametersproxy|out-string).trim())" ;
                    # pre psv2, no $hshOutput.PSBoundParametersproxy autovari to check, so back them out:
                    if($hshOutput.PSCmdletproxy.MyInvocation.InvocationName){
                        # has func name under AdvFuncs
                        if($hshOutput.PSCmdletproxy.MyInvocation.InvocationName  -match '^\.'){
                            $smsg = "detected dot-sourced invocation: Skipping `$PSCmdlet.MyInvocation.InvocationName-tied cmds..." ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } else { 
                            write-verbose 'Collect all non-default Params (works back to psv2 w CmdletBinding)'
                            $hshOutput.ParamsNonDefault = (Get-Command $hshOutput.PSCmdletproxy.MyInvocation.InvocationName).parameters | 
                                Select-Object -expand keys | 
                                Where-Object{$_ -notmatch '(Verbose|Debug|ErrorAction|WarningAction|ErrorVariable|WarningVariable|OutVariable|OutBuffer)'} ;
                        } ; 
                    } else { 
                        $smsg = "(blank `$hshOutput.PSCmdletproxy.MyInvocation.InvocationName, skipping Parameters collection)" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    } ; 
                    if($hshOutput.isScript){
                        $hshOutput.ScriptDir = $scriptName = '' ;     
                        if($hshOutput.isScript){
                            $hshOutput.ScriptDir = $hshOutput.PSScriptRootproxy; 
                            $hshOutput.ScriptName = $hshOutput.PSCommandPathproxy ; 
                            if($hshOutput.ScriptDir -eq '' -AND $hshOutput.runSource -eq 'ExternalScript'){$hshOutput.ScriptDir = (Split-Path -Path $hshOutput.MyInvocationproxy.MyCommand.Source -Parent)} # Running from File
                        };

                        if($hshOutput.ScriptDir -eq '' -AND (Test-Path variable:psEditor)) {
                            write-verbose "Running from VSCode|VS" ; 
                            $hshOutput.ScriptDir = (Split-Path -Path $psEditor.GetEditorContext().CurrentFile.Path -Parent) ; 
                                if($hshOutput.ScriptName -eq ''){$hshOutput.ScriptName = $psEditor.GetEditorContext().CurrentFile.Path }; 
                        } ;
                        if ($hshOutput.ScriptDir -eq '' -AND $host.version.major -lt 3 -AND $hshOutput.MyInvocationproxy.MyCommand.Path.length -gt 0){
                            $hshOutput.ScriptDir = $hshOutput.MyInvocationproxy.MyCommand.Path ; 
                            write-verbose "(backrev emulating `$hshOutput.PSScriptRootproxy, `$hshOutput.PSCommandPathproxy)"
                            $hshOutput.ScriptName = split-path $hshOutput.MyInvocationproxy.MyCommand.Path -leaf ;
                            $hshOutput.PSScriptRootproxy = Split-Path $hshOutput.ScriptName -Parent ;
                            $hshOutput.PSCommandPathproxy = $hshOutput.ScriptName ;
                        } ;
                        if ($hshOutput.ScriptDir -eq '' -AND $hshOutput.MyInvocationproxy.MyCommand.Path.length){
                            if($hshOutput.ScriptName -eq ''){$hshOutput.ScriptName = $hshOutput.MyInvocationproxy.MyCommand.Path} ;
                            $hshOutput.ScriptDir = $hshOutput.PSScriptRootproxy = Split-Path $hshOutput.MyInvocationproxy.MyCommand.Path -Parent ;
                        }
                        if ($hshOutput.ScriptDir -eq ''){throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$hshOutput.MyInvocationproxy IS BLANK!" } ;
                        if($hshOutput.ScriptName){
                            if(-not $hshOutput.ScriptDir ){$hshOutput.ScriptDir = Split-Path -Parent $hshOutput.ScriptName} ; 
                            $hshOutput.ScriptBaseName = split-path -leaf $hshOutput.ScriptName ;
                            $hshOutput.ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($hshOutput.ScriptName) ;
                        } ; 
                        # blank $cmdlet name comming through, patch it for Scripts:
                        if(-not $hshOutput.CmdletName -AND $hshOutput.ScriptBaseName){
                            $hshOutput.CmdletName = $hshOutput.ScriptBaseName
                        }
                        # last ditch patch the values in if you've got a $hshOutput.ScriptName
                        if($hshOutput.PSScriptRootproxy.Length -ne 0){}else{ 
                            if($hshOutput.ScriptName){$hshOutput.PSScriptRootproxy = Split-Path $hshOutput.ScriptName -Parent }
                            else{ throw "Unpopulated, `$hshOutput.PSScriptRootproxy, and no populated `$hshOutput.ScriptName from which to emulate the value!" } ; 
                        } ; 
                        if($hshOutput.PSCommandPathproxy.Length -ne 0){}else{ 
                            if($hshOutput.ScriptName){$hshOutput.PSCommandPathproxy = $hshOutput.ScriptName }
                            else{ throw "Unpopulated, `$hshOutput.PSCommandPathproxy, and no populated `$hshOutput.ScriptName from which to emulate the value!" } ; 
                        } ; 
                        if(-not ($hshOutput.ScriptDir -AND $hshOutput.ScriptBaseName -AND $hshOutput.ScriptNameNoExt  -AND $hshOutput.PSScriptRootproxy  -AND $hshOutput.PSCommandPathproxy )){ 
                            throw "Invalid Invocation. Blank `$hshOutput.ScriptDir/`$hshOutput.ScriptBaseName/`$hshOutput.ScriptBaseName" ; 
                            BREAK ; 
                        } ; 
                    } ; 
                    if($hshOutput.isFunc){
                        if($hshOutput.isFuncAdv){
                            # AdvFunc-specific cmds
                        }else {
                            # Basic Func-specific cmds
                        } ; 
                        if($hshOutput.PSCommandPathproxy -match '\.psm1$'){
                            write-host "MODULE-HOMED FUNCTION:Use `$hshOutput.CmdletName to reference the running function name for transcripts etc (under a .psm1 `$hshOutput.ScriptName will reflect the .psm1 file  fullname)"
                            if(-not $hshOutput.CmdletName){write-warning "MODULE-HOMED FUNCTION with BLANK `$CmdletNam:$($CmdletNam)" } ;
                        } # Running from .psm1 module
                        if(-not $hshOutput.CmdletName -AND $hshOutput.FuncName){
                            $hshOutput.CmdletName = $hshOutput.FuncName
                        } ; 
                    } ; 
                    $smsg = "`$hshOutput  w`n$(($hshOutput|out-string).trim())" ; 
                    #write-host $smsg ; 
                    write-verbose $smsg ; 
                } ;  # BEG-E
                PROCESS {};  # PROC-E
                END {
                    if($hshOutput){
                        write-verbose "(return `$hshOutput to pipeline)" ; 
                        New-Object PSObject -Property $hshOutput | write-output 
                    } ; 
                }
            } ; 
            #*------^ END Function resolve-EnvironmentTDO() ^------ 
        } ;
        #endregion RESOLVE_ENVIRONMENTTDO ; #*------^ END resolve-EnvironmentTDO ^------
    
        #region WRITE_LOG ; #*------v write-log v------
        if(-not(gci function:write-log -ea 0)){
            function write-log  {
                <#
                .SYNOPSIS
                Write-Log.ps1 - Write-Log writes a message to a specified log file with the current time stamp, and write-verbose|warn|error's the matching msg.
                .NOTES
                REVISIONS
                * 1:42 PM 11/8/2024 CBH expl fixes
                .PARAMETER Object <System.Object>
                    Objects to display in the host.
                    .PARAMETER Path
                    The path to the log file to which you would like to write. By default the function will create the path and file if it does not exist.
                    .PARAMETER Level
                    Specify the criticality of the log information being written to the log (defaults Info): (Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success)[-level Info]
                    .PARAMETER useHost
                    Switch to use write-host rather than write-[verbose|warn|error] (does not apply to H1|H2|H3|DEBUG which alt via uncolored write-host) [-useHost]
                    .PARAMETER NoEcho
                    Switch to suppress console echos (e.g log to file only [-NoEcho]
                    .PARAMETER NoClobber
                    Use NoClobber if you do not wish to overwrite an existing file.
                    .PARAMETER BackgroundColor
                    Specifies the background color. There is no default. The acceptable values for this parameter are:
                    (Black | DarkBlue | DarkGreen | DarkCyan | DarkRed | DarkMagenta | DarkYellow | Gray | DarkGray | Blue | Green | Cyan | Red | Magenta | Yellow | White)
                    .PARAMETER ForegroundColor <System.ConsoleColor>
                    Specifies the text color. There is no default. The acceptable values for this parameter are:
                    (Black | DarkBlue | DarkGreen | DarkCyan | DarkRed | DarkMagenta | DarkYellow | Gray | DarkGray | Blue | Green | Cyan | Red | Magenta | Yellow | White)
                    .PARAMETER NoNewline <System.Management.Automation.SwitchParameter>
                    The string representations of the input objects are concatenated to form the output. No spaces or newlines are inserted between
                    the output strings. No newline is added after the last output string.
                    .PARAMETER Separator <System.Object>
                    Specifies a separator string to insert between objects displayed by the host.
                    .PARAMETER PadChar
                    Character to use for padding (defaults to a space).[-PadChar '-']
                    .PARAMETER usePID
                    Switch to use the `$PID in the `$env:HostIndentSpaces name (Env:HostIndentSpaces`$PID)[-usePID]
                    .PARAMETER Indent
                    Switch to use write-HostIndent-type code for console echos(see get-help write-HostIndent)[-Indent]
                    .PARAMETER Flatten
                    Switch to strip empty lines when using -Indent (which auto-splits multiline Objects)[-Flatten]
                    .PARAMETER ShowDebug
                    Parameter to display Debugging messages [-ShowDebug switch]
                    .PARAMETER demo
	                Switch to output a demo display of each Level, and it's configured color scheme (requires specification of a 'dummy' message string to avoid an error).[-Demo]
                    .EXAMPLE
                    PS>  Write-Log -Message 'Log message'   ;
                    Writes the message to default log loc (c:\Logs\PowerShellLog.log, -level defaults to Info).
                    .EXAMPLE
                    PS> Write-Log -Message 'Restarting Server.' -Path c:\Logs\Scriptoutput.log ;
                    Writes the content to the specified log file and creates the path and file specified.
                    .EXAMPLE
                    PS> write-log -level warn "some information" -Path c:\tmp\tmp.txt
                    WARNING: 10:17:59: some information
                    Demo default use of the native write-warning cmdlet (default behavior when -useHost is not used)
                    .EXAMPLE
                    PS> write-log -level warn "some information" -Path c:\tmp\tmp.txt -usehost
                        10:19:14: WARNING: some information
                    Demo use of the "warning" color scheme write-host cmdlet (behavior when -useHost *IS* used)
                    .EXAMPLE
                    PS> Write-Log -level Prompt -Message "Enter Text:" -Path c:\tmp\tmp.txt -usehost  ;
                    PS> invoke-soundcue -type question ;
                    PS> $enteredText = read-host ;
                    Echo's a distinctive Prompt color scheme for the message (vs using read-host native non-color-differentiating -prompt parameter), and writes a 'Prompt'-level entry to the log, uses my verb-io:invoke-soundCue to play a the system question sound; then uses promptless read-host to take typed input.
                    PS> Write-Log -level Prompt -Message "Enter Password:" -Path c:\tmp\tmp.txt -usehost  ;
                    PS> invoke-soundcue -type question ;
                    PS> $SecurePW = read-host -AsSecureString ;
                    Variant that demos collection of a secure password using read-host's native -AsSecureString param.
                    .EXAMPLE
                    PS>  $smsg = "ENTER CERTIFICATE PFX Password: (use 'dummy' for UserName)" ;
                    PS>  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT }
                    PS>  else{ write-host -foregroundcolor Blue -backgroundcolor White "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    PS>  $pfxcred=(Get-Credential -credential dummy) ;
                    PS>  Export-PfxCertificate -Password $pfxcred.password -Cert= $certpath -FilePath c:\path-to\output.pfx;
                    Demo use of write-log -level prompt, leveraging the get-credential popup GUI to collect a secure password (without use of username)
                    .EXAMPLE
                    PS>  # init content in script context ($MyInvocation is blank in function scope)
                    PS>  $logfile = join-path -path $ofile -childpath "$([system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName))-BATCH-$(get-date -format 'yyyyMMdd-HHmmtt')-LOG.txt"  ;
                    PS>  $logging = $True ;
                    PS>  $sBnr="#*======v `$tmbx:($($Procd)/$($ttl)):$($tmbx) v======" ;
                    PS>  $smsg="$($sBnr)" ;
                    PS>  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug|H1|H2|H3
                    PS>  else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Demo with conditional write-log (with -useHost switch, to trigger native write-host use), else failthru to write-host output
                    PS>  .EXAMPLE
                    PS>  $transcript = join-path -path (Split-Path -parent $MyInvocation.MyCommand.Definition) -ChildPath "logs" ;
                    PS>  if(!(test-path -path $transcript)){ "Creating missing log dir $($transcript)..." ; mkdir $transcript  ; } ;
                    PS>  $transcript=join-path -path $transcript -childpath "$([system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName))"  ;
                    PS>  $transcript+= "-Transcript-BATCH-$(get-date -format 'yyyyMMdd-HHmmtt')-trans-log.txt"  ;
                    PS>  # add log file variant as target of Write-Log:
                    PS>  $logfile=$transcript.replace("-Transcript","-LOG").replace("-trans-log","-log")
                    PS>  if($whatif){
                    PS>      $logfile=$logfile.replace("-BATCH","-BATCH-WHATIF") ;
                    PS>      $transcript=$transcript.replace("-BATCH","-BATCH-WHATIF") ;
                    PS>  } else {
                    PS>      $logfile=$logfile.replace("-BATCH","-BATCH-EXEC") ;
                    PS>      $transcript=$transcript.replace("-BATCH","-BATCH-EXEC") ;
                    PS>  } ;
                    PS>  if($Ticket){
                    PS>      $logfile=$logfile.replace("-BATCH","-$($Ticket)") ;
                    PS>      $transcript=$transcript.replace("-BATCH","-$($Ticket)") ;
                    PS>  } else {
                    PS>      $logfile=$logfile.replace("-BATCH","-nnnnnn") ;
                    PS>      $transcript=$transcript.replace("-BATCH","-nnnnnn") ;
                    PS>  } ;
                    PS>  $logging = $True ;
                    PS>  $sBnr="#*======v START PASS:$($ScriptBaseName) v======" ;
                    PS>  $smsg= "$($sBnr)" ;
                    PS>  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                    More complete boilerplate including $whatif & $ticket
                    .EXAMPLE
                    PS>  $pltSL=@{ NoTimeStamp=$false ; Tag = $null ; showdebug=$($showdebug) ; whatif=$($whatif) ; Verbose=$($VerbosePreference -eq 'Continue') ; } ;
                    PS>  $pltSL.Tag = "$(split-path -path $CSVPath -leaf)"; # build tag from a variable
                    PS>  # construct log name on calling script/function fullname
                    PS>  if($PSCommandPath){ $logspec = start-Log -Path $PSCommandPath @pltSL }
                    PS>  else { $logspec = start-Log -Path ($MyInvocation.MyCommand.Definition) @pltSL } ;
                    PS>  if($logspec){
                    PS>      $logging=$logspec.logging ;
                    PS>      $logfile=$logspec.logfile ;
                    PS>      $transcript=$logspec.transcript ;
                    PS>      $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                    PS>      start-Transcript -path $transcript ;
                    PS>  } else {throw "Unable to configure logging!" } ;
                    PS>  $sBnr="#*======v $(${CmdletName}): v======" ;
                    PS>  $smsg = $sBnr ;
                    PS>  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    PS>  else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Demo leveraging splatted start-log(), and either $PSCommandPath or $MyInvocation (support varies by host/psversion) to build the log name.
                    .EXAMPLE
                    PS> write-log -demo -message 'Dummy' ;
                    Demo (using required dummy error-suppressing messasge) of sample outputs/color combos for each Level configured).
                    .EXAMPLE
                    PS>  $smsg = "`n`n===TESTIPAddress: was *validated* as covered by the recursed ipv4 specification:" ;
                    PS>  $smsg += "`n" ;
                    PS>  $smsg += "`n---> This host *should be able to* send email on behalf of the configured SPF domain (at least in terms of SPF checks)" ;
                    PS>  $env:hostindentspaces = 8 ;
                    PS>  $lvl = 'Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success'.split('|') ;
                    PS>  foreach ($l in $lvl){Write-Log -LogContent $smsg -Path $tmpfile -Level $l -useHost -Indent} ;
                    Demo indent function across range of Levels (alt to native -Demo which also supports -indent).
                    .EXAMPLE
                    PS>  write-verbose 'set to baseline' ;
                    PS>  reset-HostIndent ;
                    PS>  write-host "`$env:HostIndentSpaces:$($env:HostIndentSpaces)" ;
                    PS>  write-verbose 'write an H1 banner'
                    PS>  $sBnr="#*======v  H1 Banner: v======" ;
                    PS>  $smsg = $sBnr ;
                    PS>  Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1;
                    PS>  write-verbose 'push indent level+1' ;
                    PS>  push-HostIndent ;
                    PS>  write-host "`$env:HostIndentSpaces:$($env:HostIndentSpaces)" ;
                    PS>  write-verbose 'write an INFO entry with -Indent specified' ;
                    PS>  $smsg = "This is information (indented)" ;
                    PS>  Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info -Indent:$true ;
                    PS>  write-verbose 'push indent level+2' ;
                    PS>  push-HostIndent ;
                    PS>  write-host "`$env:HostIndentSpaces:$($env:HostIndentSpaces)" ;
                    PS>  write-verbose 'write a PROMPT entry with -Indent specified' ;
                    PS>  $smsg = "This is a subset of information (indented)" ;
                    PS>  Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt -Indent:$true ;
                    PS>  write-verbose 'pop indent level out one -1' ;
                    PS>  pop-HostIndent ;
                    PS>  write-verbose 'write a Success entry with -Indent specified' ;
                    PS>  $smsg = "This is a Successful information (indented)" ;
                    PS>  Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success -Indent:$true ;
                    PS>  write-verbose 'reset to baseline for trailing banner'
                    PS>  reset-HostIndent ;
                    PS>  write-host "`$env:HostIndentSpaces:$($env:HostIndentSpaces)" ;
                    PS>  write-verbose 'write the trailing H1 banner'
                    PS>  $smsg = "$($sBnr.replace('=v','=^').replace('v=','^='))" ;
                    PS>  Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1;
                    PS>  write-verbose 'clear indent `$env:HostIndentSpaces' ;
                    PS>  clear-HostIndent ;
                    PS>  write-host "`$env:HostIndentSpaces:$($env:HostIndentSpaces)" ;
                        $env:HostIndentSpaces:0
                        16:16:17: #  #*======v  H1 Banner: v======
                        $env:HostIndentSpaces:4
                            16:16:17: INFO:  This is information (indented)
                        $env:HostIndentSpaces:8
                                16:16:17: PROMPT:  This is a subset of information (indented)
                            16:16:17: SUCCESS:  This is a Successful information (indented)
                        $env:HostIndentSpaces:0
                        16:16:17: #  #*======^  H1 Banner: ^======
                        $env:HostIndentSpaces:
                    Demo broad process for use of verb-HostIndent funcs and write-log with -indent parameter.
                    .EXAMPLE
                    PS>  write-host "`n`n" ;
                    PS>  $smsg = "`n`n==ALL Grouped Status.errorCode :`n$(($EVTS.status.errorCode | group| sort count -des | format-table -auto count,name|out-string).trim())" ;
                    PS>  $colors = (get-colorcombo -random) ;
                    PS>  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info -Indent @colors -flatten }
                    PS>  else{ write-host @colors  "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    PS>  PS>  write-host "`n`n" ;
                    When using -Indent with group'd or other cmd-multiline output, you will want to:
                    1. use the...
                        $smsg = $(([results]|out-string).trim())"
                        ...structure to pre-clean & convert from [FormatEntryData] to [string]
                        (avoids errors, due to formatteddata *not* having split mehtod)
                    2. Use -flatten to avoid empty _colored_ lines between each entry in the output (and sprinkle write-host "`n`n"'s pre/post for separation).
                    These issues only occur under -Indent use, due to the need to `$Object.split to get each line of indented object properly collored and indented.
                    .EXAMPLE
                    PS> $pltH1=@{foregroundcolor='black';backgroundcolor='darkyellow'};
                    PS> write-host "Running demo of current settings..." @pltH1
                    PS> $combos = "H1; #*======v STATUSMSG: SBNR v======","H2;`n#*------v PROCESSING : sBnrS v------","H3;`n#*~~~~~~v SUB-PROCESSING : sBnr3 v~~~~~~","H4;`n#*``````v DETAIL : sBnr4 v``````","H5;`n#*______v FOCUS : sBnr5 v______","INFO;This is typical output","PROMPT;What is your quest?","SUCCESS;Successful execution!","WARN;THIS DIDN'T GO AS PLANNED","ERROR;UTTER FAILURE!","VERBOSE;internal comment executed"
                    PS> $tmpfile = [System.IO.Path]::GetTempFileName().replace('.tmp','.txt') ;
                    PS> foreach($cmbo in $combos){
                    PS>     $level,$text = $cmbo.split(';') ;
                    PS>     $pltWL=@{
                    PS>         message= $text ;
                    PS>         Level=$Level ;
                    PS>         Path=$tmpfile  ;
                    PS>         useHost=$true;
                    PS>     } ;
                    PS>     if($Indent){$PltWL.add('Indent',$true)} ;
                    PS>     $whsmsg = "write-log w`n$(($pltWL|out-string).trim())`n" ;
                    PS>     write-host $whsmsg ;
                    PS>     write-logNoDep @pltWL ;
                    PS> } ;
                    PS> remove-item -path $tmpfile ;
                    Demo code adapted from the -demo param, for manual passes.
                #>[CmdletBinding()]
                PARAM (
                        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true,
                            HelpMessage = "Message is the content that you wish to add to the log file")]
                            [ValidateNotNullOrEmpty()]
                            [Alias("LogContent",'Message')]
                            [System.Object]$Object,
                        [Parameter(Mandatory = $false,
                            HelpMessage = "The path to the log file to which you would like to write. By default the function will create the path and file if it does not exist.")]
                            [Alias('LogPath')]
                            [string]$Path = 'C:\Logs\PowerShellLog.log',
                        [Parameter(Mandatory = $false,
                            HelpMessage = "Specify the criticality of the log information being written to the log (defaults Info): (Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success)[-level Info]")]
                            [ValidateSet('Error','Warn','Info','H1','H2','H3','H4','H5','Debug','Verbose','Prompt','Success')]
                            [string]$Level = "Info",
                        [Parameter(
                            HelpMessage = "Switch to use write-host rather than write-[verbose|warn|error] [-useHost]")]
                            [switch] $useHost,
                        [Parameter(
                            HelpMessage="Specifies the background color. There is no default. The acceptable values for this parameter are:
                    (Black | DarkBlue | DarkGreen | DarkCyan | DarkRed | DarkMagenta | DarkYellow | Gray | DarkGray | Blue | Green | Cyan | Red | Magenta | Yellow | White)")]
                            [System.ConsoleColor]$BackgroundColor,
                        [Parameter(
                            HelpMessage="Specifies the text color. There is no default. The acceptable values for this parameter are:
                (Black | DarkBlue | DarkGreen | DarkCyan | DarkRed | DarkMagenta | DarkYellow | Gray | DarkGray | Blue | Green | Cyan | Red | Magenta | Yellow | White)")]
                            [System.ConsoleColor]$ForegroundColor,
                        [Parameter(
                            HelpMessage="The string representations of the input objects are concatenated to form the output. No spaces or newlines are inserted between
                the output strings. No newline is added after the last output string.")]
                            [System.Management.Automation.SwitchParameter]$NoNewline,
                        [Parameter(
                            HelpMessage = "Switch to use write-HostIndent-type code for console echos(see get-help write-HostIndent)[-Indent]")]
                            [Alias('in')]
                            [switch] $Indent,
                        [Parameter(
                            HelpMessage="Switch to use the `$PID in the `$env:HostIndentSpaces name (Env:HostIndentSpaces`$PID)[-usePID]")]
                            [switch]$usePID,
                        [Parameter(
                            HelpMessage = "Switch to strip empty lines when using -Indent (which auto-splits multiline Objects)[-Flatten]")]
                            #[Alias('flat')]
                            [switch] $Flatten,
                        [Parameter(
                            HelpMessage="Specifies a separator string to insert between objects displayed by the host.")]
                        [System.Object]$Separator,
                        [Parameter(
                            HelpMessage="Character to use for padding (defaults to a space).[-PadChar '-']")]
                        [string]$PadChar = ' ',
                        [Parameter(
                            HelpMessage="Number of spaces to pad by default (defaults to 4).[-PadIncrment 8]")]
                        [int]$PadIncrment = 4,
                        [Parameter(
                            HelpMessage = "Switch to suppress console echos (e.g log to file only [-NoEcho]")]
                            [switch] $NoEcho,
                        [Parameter(Mandatory = $false,
                            HelpMessage = "Use NoClobber if you do not wish to overwrite an existing file.")]
                            [switch]$NoClobber,
                        [Parameter(
                            HelpMessage = "Debugging Flag [-showDebug]")]
                            [switch] $showDebug,
                        [Parameter(
                            HelpMessage = "Switch to output a demo display of each Level, and it's configured color scheme (requires specification of a 'dummy' message string to avoid an error).[-Demo]")]
                            [switch] $demo
                    )  ;
                BEGIN {
                    #region CONSTANTS-AND-ENVIRO #*======v CONSTANTS-AND-ENVIRO v======
                    # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
                    ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
                    if(($PSBoundParameters.keys).count -ne 0){
                        $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
                        write-verbose "$($CmdletName): `$PSBoundParameters:`n$(($PSBoundParameters|out-string).trim())" ;
                    } ; 
                    $Verbose = ($VerbosePreference -eq 'Continue') ;     
                    #$VerbosePreference = "SilentlyContinue" ;
                    #endregion CONSTANTS-AND-ENVIRO #*======^ END CONSTANTS-AND-ENVIRO ^======

                    $pltWH = @{
                            Object = $null ;
                    } ;
                    if ($PSBoundParameters.ContainsKey('BackgroundColor')) {
                        $pltWH.add('BackgroundColor',$BackgroundColor) ;
                    } ;
                    if ($PSBoundParameters.ContainsKey('ForegroundColor')) {
                        $pltWH.add('ForegroundColor',$ForegroundColor) ;
                    } ;
                    if ($PSBoundParameters.ContainsKey('NoNewline')) {
                        $pltWH.add('NoNewline',$NoNewline) ;
                    } ;
                    if($Indent){
                        if ($PSBoundParameters.ContainsKey('Separator')) {
                            $pltWH.add('Separator',$Separator) ;
                        } ;
                        write-verbose "$($CmdletName): Using `$PadChar:`'$($PadChar)`'" ;
            
                        #if we want to tune this to a $PID-specific variant, use:
                        if($usePID){
                            $smsg = "-usePID specified: `$Env:HostIndentSpaces will be suffixed with this process' `$PID value!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $HISName = "Env:HostIndentSpaces$($PID)" ;
                        } else {
                            $HISName = "Env:HostIndentSpaces" ;
                        } ;
                        if(($smsg = Get-Item -Path "Env:HostIndentSpaces$($PID)" -erroraction SilentlyContinue).value){
                            write-verbose $smsg ;
                        } ;
                        if (-not ([int]$CurrIndent = (Get-Item -Path $HISName -erroraction SilentlyContinue).Value ) ){
                            [int]$CurrIndent = 0 ;
                        } ;
                        write-verbose "$($CmdletName): Discovered `$$($HISName):$($CurrIndent)" ;
            
                    } ;
                    if(get-command get-colorcombo -ErrorAction SilentlyContinue){$buseCC=$true} else {$buseCC=$false} ;
                   
                    if ($host.Name -eq 'Windows PowerShell ISE Host' -AND $host.version.major -lt 3){
                            write-verbose "PSISE under psV2 has wacky inconsistent colors - only *some* even display, others default to white`nso we choose fundementally wrong colors, to approximate the target colors" ;
                            $pltError=@{foregroundcolor='DarkYellow';backgroundcolor='Red'};
                            $pltWarn=@{foregroundcolor='DarkMagenta';backgroundcolor='DarkCyan'};
                            $pltInfo=@{foregroundcolor='Blue';backgroundcolor='darkGreen'};
                            $pltH1=@{foregroundcolor='black';backgroundcolor='darkyellow'};
                            $pltH2=@{foregroundcolor='darkblue';backgroundcolor='cyan'};
                            $pltH3=@{foregroundcolor='black';backgroundcolor='cyan'};
                            $pltH4=@{foregroundcolor='black';backgroundcolor='DarkMagenta'};
                            $pltH5=@{foregroundcolor='cyan';backgroundcolor='Green'};
                            $pltDebug=@{foregroundcolor='red';backgroundcolor='black'};
                            $pltVerbose=@{foregroundcolor='darkgray';backgroundcolor='darkred'};
                            $pltPrompt=@{foregroundcolor='White';backgroundcolor='DarkBlue'};
                            $pltSuccess=@{foregroundcolor='DarkGray';backgroundcolor='green'};
                    } else {
                        
                        $pltError=@{foregroundcolor='yellow';backgroundcolor='darkred'};
                        $pltWarn=@{foregroundcolor='DarkMagenta';backgroundcolor='yellow'};
                        $pltInfo=@{foregroundcolor='gray';backgroundcolor='darkblue'};
                        $pltH1=@{foregroundcolor='black';backgroundcolor='darkyellow'};
                        $pltH2=@{foregroundcolor='darkblue';backgroundcolor='gray'};
                        $pltH3=@{foregroundcolor='black';backgroundcolor='darkgray'};
                        $pltH4=@{foregroundcolor='gray';backgroundcolor='DarkCyan'};
                        $pltH5=@{foregroundcolor='cyan';backgroundcolor='DarkGreen'};
                        $pltDebug=@{foregroundcolor='red';backgroundcolor='black'};
                        $pltVerbose=@{foregroundcolor='darkgray';backgroundcolor='black'};
                        $pltPrompt=@{foregroundcolor='DarkMagenta';backgroundcolor='darkyellow'};
                        $pltSuccess=@{foregroundcolor='Blue';backgroundcolor='green'};
                    } ; 

                    if ($PSCmdlet.MyInvocation.ExpectingInput) {
                        write-verbose "Data received from pipeline input: '$($InputObject)'" ; 
                    } else {
                        #write-verbose "Data received from parameter input: '$($InputObject)'" ; 
                        write-verbose "(non-pipeline - param - input)" ; 
                    } ; 
                }  ;
                PROCESS {

                    if($Demo){
                            write-host "Running demo of current settings..." @pltH1
                            $combos = "h1m;H1","h2m;H2","h3m;H3","h4m;H4","h5m;H5",
                                "whm;INFO","whp;PROMPT","whs;SUCCESS","whw;WARN","wem;ERROR","whv;VERBOSE" ;
                            $h1m =" #*======v STATUSMSG: SBNR v======" ;
                            $h2m = "`n#*------v PROCESSING : sBnrS v------" ;
                            $h3m ="`n#*~~~~~~v SUB-PROCESSING : sBnr3 v~~~~~~" ;
                            $h4m="`n#*``````v DETAIL : sBnr4 v``````" ;
                            $h5m="`n#*______v FOCUS : sBnr5 v______" ;
                            $whm = "This is typical output" ;
                            $whp = "What is your quest?" ;
                            $whs = "Successful execution!" ;
                            $whw = "THIS DIDN'T GO AS PLANNED" ;
                            $wem = "UTTER FAILURE!" ;
                            $whv = "internal comment executed" ;
                            $tmpfile = [System.IO.Path]::GetTempFileName().replace('.tmp','.txt') ;
                            foreach($cmbo in $combos){
                                $txt,$name = $cmbo.split(';') ;
                                $Level = $name ;
                                if($Level -eq 'H5'){
                                    write-host "Gotcha!";
                                } ;
                                $whplt = (gv "plt$($name)").value ;
                                $text = (gv $txt).value ;
                                #$smsg="`$plt$($name):($($whplt.foregroundcolor):$($whplt.backgroundcolor)):`n`n$($text)`n`n" ;
                                $whsmsg="`$plt$($name):($($whplt.foregroundcolor):$($whplt.backgroundcolor)):`n`n" ;
                                $pltWL=@{
                                    message= $text ;
                                    Level=$Level ;
                                    Path=$tmpfile  ;
                                    useHost=$true;
                                } ;
                                if($Indent){$PltWL.add('Indent',$true)} ;
                                $whsmsg += "write-log w`n$(($pltWL|out-string).trim())`n" ;
                                write-host $whsmsg ;
                                write-log @pltWL ;
                            } ;
                            remove-item -path $tmpfile ;
                    } else {
            
                        if($Indent){
                            # move split/flatten into per-object level (was up in BEGIN):
                            # if $object has multiple lines, split it:
                            # have to coerce the system.object to string array, to get access to a .split method (raw object doese't have it)
                            # and you have to recast the type to string array (can't assign a string[] to [system.object] type vari
                            if($Flatten){
                                    if($object.gettype().name -eq 'FormatEntryData'){
                                        # this converts tostring() as the string: Microsoft.PowerShell.Commands.Internal.Format.FormatEntryData
                                        # issue is (group |  ft -a count,name)'s  that aren't put through $((|out-string).trim())
                                        write-verbose "skip split/flatten on these (should be pre-out-string'd before write-logging)" ;
                                    } else {
                                        TRY{
                                            [string[]]$Object = [string[]]$Object.ToString().Split([Environment]::NewLine) ; 
                                        } CATCH{
                                            write-verbose "Workaround err: The variable cannot be validated because the value System.String[] is not a valid value for the Object variable." ; 
                                            [string[]]$Object = ($Object|out-string).trim().Split([Environment]::NewLine) ; 
                                        } ; 
                                    } ;
                            } else {
                                [string[]]$Object = [string[]]$Object.ToString().Split([Environment]::NewLine)
                            } ;
                        } ; 

                        # If the file already exists and NoClobber was specified, do not write to the log.
                        if ((Test-Path $Path) -AND $NoClobber) {
                            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name."  ;
                            Return  ;
                        } elseif (!(Test-Path $Path)) {
                            Write-Verbose "Creating $Path."  ;
                            $NewLogFile = New-Item $Path -Force -ItemType File  ;
                        } else {
                          # Nothing to see here yet.
                        }  ;

                        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"  ;
                        $EchoTime = "$((get-date).ToString('HH:mm:ss')): " ;
            
                        $pltWH.Object = $EchoTime ; 
                        $pltColors = @{} ; 
                        # Write message to error, warning, or verbose pipeline and specify $LevelText
                        switch ($Level) {
                            'Error' {
                                $LevelText = 'ERROR: ' ;
                                $pltColors = $pltError ;
                                if ($useHost) {} else {if (!$NoEcho) { Write-Error ($smsg + $Object) } } ;
                            }
                            'Warn' {
                                $LevelText = 'WARNING: ' ;
                                $pltColors = $pltWarn ;
                                if ($useHost) {} else {if (!$NoEcho) { Write-Warning ($smsg + $Object) } } ;
                            }
                            'Info' {
                                $LevelText = 'INFO: ' ;
                                $pltColors = $pltInfo ;
                            }
                            'H1' {
                                $LevelText = '# ' ;
                                $pltColors = $pltH1 ;
                            }
                            'H2' {
                                $LevelText = '## ' ;
                                $pltColors = $pltH2 ;
                            }
                            'H3' {
                                $LevelText = '### ' ;
                                $pltColors = $pltH3 ;
                            }
                            'H4' {
                                $LevelText = '#### ' ;
                                $pltColors = $pltH4 ;
                            }
                            'H5' {
                                $LevelText = '##### ' ;
                                $pltColors = $pltH5 ;
                            }
                            'Debug' {
                                $LevelText = 'DEBUG: ' ;
                                $pltColors = $pltDebug ;
                                if ($useHost) {} else {if (!$NoEcho) { Write-Degug $smsg } }  ;
                            }
                            'Verbose' {
                                $LevelText = 'VERBOSE: ' ;
                                $pltColors = $pltVerbose ;
                                if ($useHost) {}else {if (!$NoEcho) { Write-Verbose ($smsg) } } ;
                            }
                            'Prompt' {
                                $LevelText = 'PROMPT: ' ;
                                $pltColors = $pltPrompt ;
                            }
                            'Success' {
                                $LevelText = 'SUCCESS: ' ;
                                $pltColors = $pltSuccess ;
                            }
                        } ;
                        # build msg string down here, once, v in ea above
                        # always defer to explicit cmdline colors
                        if($pltColors.foregroundcolor){
                            if(-not ($pltWH.keys -contains 'foregroundcolor')){
                                $pltWH.add('foregroundcolor',$pltColors.foregroundcolor) ;
                            } elseif($pltWH.foregroundcolor -eq $null){
                                $pltWH.foregroundcolor = $pltColors.foregroundcolor ;
                            } ;
                        } ;
                        if($pltColors.backgroundcolor){
                            if(-not ($pltWH.keys -contains 'backgroundcolor')){
                                $pltWH.add('backgroundcolor',$pltColors.backgroundcolor) ;
                            } elseif($pltWH.backgroundcolor -eq $null){
                                $pltWH.backgroundcolor = $pltColors.backgroundcolor ;
                            } ;
                        } ;
 
                        if ($useHost) {
                            if(-not $Indent){
                                if($Level -match '(Debug|Verbose)' ){
                                    if(($Object|  measure).count -gt 1){
                                        $pltWH.Object += "$($LevelText) ($(($Object|out-string).trim()))" ;
                                    } else {
                                        #$pltWH.Object += ($LevelText + '(' + $Object + ')') ;
                                        $pltWH.Object += "$($LevelText) ($($Object))" ;
                                    } ;
                                } else {
                                    if(($Object|  measure).count -gt 1){
                                        $pltWH.Object += "$($LevelText) $(($Object|out-string).trim())" ;
                                    } else {
                                        #$pltWH.Object += $LevelText + $Object ;
                                        $pltWH.Object += "$($LevelText) $($Object)" ;
                                    } ;
                                } ;
                                $smsg = "write-host w`n$(($pltWH|out-string).trim())" ;
                                write-verbose $smsg ;
                                #write-host @pltErr $smsg ;
                                write-host @pltwh ;
                            } else {
                                foreach ($obj in $object){
                                    $pltWH.Object = $EchoTime ;
                                    if($Level -match '(Debug|Verbose)' ){
                                        if($obj.length -gt 0){
                                            $pltWH.Object += "$($LevelText) ($($obj))" ;
                                        } else {
                                            $pltWH.Object += "$($LevelText)" ;
                                        } ;
                                    } else {
                                        $pltWH.Object += "$($LevelText) $($obj)" ;
                                    } ;
                                    $smsg = "write-host w`n$(($pltWH|out-string).trim())" ;
                                    write-verbose $smsg ;
                                    Write-Host -NoNewline $($PadChar * $CurrIndent)  ;
                                    write-host @pltwh ;
                                } ;
                            } ;
                        }
                        # Write log entry to $Path
                        "$FormattedDate $LevelText : $Object" | Out-File -FilePath $Path -Append -encoding UTF8 ;
                    } ;  # if-E -Demo ; 
                }  ; # PROC-E
                End {}  ;
                  
            }
        } ; 
        #endregion WRITE_LOG ; #*------^ END write-log  ^------

        #region START_LOG ; #*------v Start-Log v------
        if(-not(get-command start-log -ea 0)){
            function Start-Log {
                <#
                .SYNOPSIS
                Start-Log.ps1 - Configure base settings for use of write-Log() logging
                .NOTES
                REVISIONS
               * 9:07 AM 4/30/2025 make Tag cleanup conditional on avail of the target vtxt\funcs
               #>
                [CmdletBinding()]
                PARAM(
                    [Parameter(Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Path to target script (defaults to `$PSCommandPath) [-Path .\path-to\script.ps1]")]
                    # rem out validation, for module installed in AllUsers etc, we don't want to have to spec a real existing file. No bene to testing
                    #[ValidateScript({Test-Path (split-path $_)})] 
                    $Path,
                    [Parameter(HelpMessage="Tag string to be used with -Path filename spec, to construct log file name [-tag 'ticket-123456]")]
                    [string]$Tag,
                    [Parameter(HelpMessage="Flag that suppresses the trailing timestamp value from the generated filenames[-NoTimestamp]")]
                    [switch] $NoTimeStamp,
                    [Parameter(HelpMessage="Flag that leads the returned filename with the Tag parameter value[-TagFirst]")]
                    [switch] $TagFirst,
                    [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
                    [switch] $showDebug,
                    [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
                    [switch] $whatIf=$true
                ) ;
                #${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ;
                #$PSParameters = New-Object -TypeName PSObject -Property $rPSBoundParameters ;
                $Verbose = ($VerbosePreference -eq 'Continue') ; 
                $transcript = join-path -path (Split-Path -parent $Path) -ChildPath "logs" ;
                if (-not (test-path -path $transcript)) { write-host "Creating missing log dir $($transcript)..." ; mkdir $transcript  ; } ;
                #$transcript = join-path -path $transcript -childpath "$([system.io.path]::GetFilenameWithoutExtension($Path))" ; 
                if($Tag){
                    # clean for fso use, if funcs avail
                    if((gci function:Remove-StringDiacritic -ea 0)){$Tag = Remove-StringDiacritic -String $Tag } else {write-host "(missing:verb-text\Remove-StringDiacritic, skipping)";}  # verb-text ; 
                    if((gci function:Remove-StringLatinCharacters -ea 0)){$Tag = Remove-StringLatinCharacters -String $Tag } else {write-host "(missing:verb-textRemove-StringLatinCharacters, skipping)";} # verb-text
                    if((gci function:Remove-InvalidFileNameChars -ea 0)){$Tag = Remove-InvalidFileNameChars -Name $Tag } else {write-host "(missing:verb-textRemove-InvalidFileNameChars, skipping)";}; # verb-io, (inbound Path is assumed to be filesystem safe)
                    if($TagFirst){
                        $smsg = "(-TagFirst:Building filenames with leading -Tag value)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        $transcript = join-path -path $transcript -childpath "$($Tag)-$([system.io.path]::GetFilenameWithoutExtension($Path))" ; 
                        #$transcript = "$($Tag)-$($transcript)" ; 
                    } else { 
                        $transcript = join-path -path $transcript -childpath "$([system.io.path]::GetFilenameWithoutExtension($Path))" ; 
                        $transcript += "-$($Tag)" ; 
                    } ;
                } else {
                    $transcript = join-path -path $transcript -childpath "$([system.io.path]::GetFilenameWithoutExtension($Path))" ; 
                }; 
                $transcript += "-Transcript-BATCH"
                if(-not $NoTimeStamp){ $transcript += "-$(get-date -format 'yyyyMMdd-HHmmtt')" } ; 
                $transcript += "-trans-log.txt"  ;
                # add log file variant as target of Write-Log:
                $logfile = $transcript.replace("-Transcript", "-LOG").replace("-trans-log", "-log") ;
                if(get-variable whatif -ea 0){
                    if ($whatif) {
                        $logfile = $logfile.replace("-BATCH", "-BATCH-WHATIF") ;
                        $transcript = $transcript.replace("-BATCH", "-BATCH-WHATIF") ;
                    } else {
                        $logfile = $logfile.replace("-BATCH", "-BATCH-EXEC") ;
                        $transcript = $transcript.replace("-BATCH", "-BATCH-EXEC") ;
                    } ;
                } ; 
                $logging = $True ;

                $hshRet= [ordered]@{
                    logging=$logging ;
                    logfile=$logfile ;
                    transcript=$transcript ;
                } ;
                if($showdebug -OR $verbose){
                    write-verbose -verbose:$true "$(($hshRet|out-string).trim())" ;  ;
                } ;
                Write-Output $hshRet ;
            }
        } ; 
        #endregion START_LOG ;  ; #*------^ END start-log ^------

        #region CONNECT_EXCHANGESERVERTDO ; #*------v Connect-ExchangeServerTDO v------
        if(-not(gci function:Connect-ExchangeServerTDO -ea 0)){
            Function Connect-ExchangeServerTDO {
                <#
                .SYNOPSIS
                Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
                will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellREmote (REMS) connect to each server, 
                stopping at the first successful connection.
                .NOTES
                REVISIONS
                * 10;07 am 4/30/2025 fixed borked edge conn, typo, and rev logic for Ex & role detection in raw PS - lacks evaris for exchange (EMS/REMS only), so leverage reg & stock install loc hunting to discover setup.exe for vers & role confirm).
                * 2:46 PM 4/22/2025 add: -Version (default to Ex2010), and postfiltered returned ExchangeServers on version. If no -Version, sort on newest Version, then name, -descending.
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
                .PARAMETER Version
                Specific Exchange Server Version to connect to('Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000')[-Version 'Ex2016']
                .PARAMETER TenOrg
                Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
                .INPUTS
                None. Does not accepted piped input.(.NET types, can add description)
                .OUTPUTS
                [system.object] Returns a system object containing a successful PSSession
                System.Boolean
                .EXAMPLE
                PS> $PSSession = Connect-ExchangeServerTDO -siteName SITENAME -RoleNames @('HUB','CAS') -verbose 
                Demo's connecting to a functional Hub or CAS server in the SITENAME site with verbose outputs, the `PSSession variable will contain information about the successful connection. Makes automatic Exchangeserver discovery calls into AD (using ADSI) leveraging the separate get-ADExchangeServerTDO()
                .EXAMPLE
                PS> TRY{$Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name}CATCH{$Site=$env:COMPUTERNAME} ;
                PS> $PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
                Demo including support for EdgeRole, which is detected on it's lack of AD Site specification (which gets fed through to call, by setting the Site to the machine itself).
                .EXAMPLE
                PS> $PSSession = Connect-ExchangeServerTDO -siteName SITENAME -RoleNames @('HUB','CAS') -Version Ex2016 -verbose 
                Demo's connecting to a functional Hub or CAS server Version Ex2016 in the SITENAME site with verbose outputs, the `PSSession variable will contain information about the successful connection. Makes automatic Exchangeserver discovery calls into AD (using ADSI) leveraging the separate get-ADExchangeServerTDO()
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
                    [Parameter(Position=2,HelpMessage="Specific Exchange Server Version to connect to('Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000')[-Version 'Ex2016']")]
                        [ValidateSet('Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000')]
                        [string[]]$Version = 'Ex2010',
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
                    
                    # 5:15 PM 4/22/2025 on CMW, have to patch version to Ex2016

                    #*------v Function _connect-ExOP v------
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
                                $pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ;} ;
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
                                # 3:59 PM 1/9/2025 appears credprompting is due to it's missing the import-module $ExIPSS ! 
                                $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                $Global:E10Mod = Import-Module $ExIPSS @pltIMod ;
                                $ExPSS | write-output ;
                                $ExPSS= $ExIPSS = $null ;
                            } ; 
                        } ;
                    #*------^ END Function _connect-ExOP ^------
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
                            $pltCXOP.Add('Credential',$pltGADX.credential) ;
                        } ;
                        if($Version){
                            switch ($Version){
                              'Ex2000'{$rgxExVersNum = '6' } 
                              'Ex2003'{$rgxExVersNum = '6.5' } 
                              'Ex2007'{$rgxExVersNum = '8.*' } 
                              'Ex2010'{$rgxExVersNum = '14.*'} 
                              'Ex2013'{$rgxExVersNum = '15.0' } 
                              'Ex2016'{$rgxExVersNum = '15.1'} 
                              'Ex2019'{$rgxExVersNum = '15.2' } 
                            } ; 
                            $exchServers  = $exchServers | ?{ [double]([regex]::match( $_.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value) -match $rgxExVersNum } ; 

                        } else {
                            write-verbose "no -Version: Sorting Newest first, then names, descending" ; 
                            $exchServers  = $exchServers | sort version,name -desc
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
                                $smsg = "_connect-ExOP w`n$(($pltCXOP|out-string).trim())" ;
                                $smsg += "`nServer $($exServer.FQDN)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
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
        #endregion CONNECT_EXCHANGESERVERTDO ; #*------^ END Connect-ExchangeServerTDO ^------

        #region GET_ADEXCHANGESERVERTDO ; #*------v get-ADExchangeServerTDO v------
        if(-not(gci function:get-ADExchangeServerTDO -ea 0)){
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
                * 10;05 am 4/30/2025 fixed code for Edge role in raw PS, missing evaris for Ex: added discovery from reg & stock file system dirs for version etc.
                * 3:57 PM 11/26/2024 updated simple write-host,write-verbose with full pswlt support;  syncd dbg & vx10 copies.
                * 12:57 PM 6/11/2024 Validated, Ex2010 & Ex2019, hub, mail & edge roles: tested ☑️ on CMW mail role (Curly); and Jumpbox; copied in CBH from repo copy, which has been updated/debugged compat on CMW Edge 
                * 2:05 PM 8/28/2023 REN -> Get-ExchangeServerInSite -> get-ADExchangeServerTDO (aliased orig); to better steer profile-level options - including in cmw org, added -TenOrg, and default Site to constructed vari, targeting new profile $XXX_ADSiteDefault vari; Defaulted -Roles to HUB,CAS as well.
                * 3:42 PM 8/24/2023 spliced together combo of my long-standing, and some of the interesting ideas BF's version had. Functional prod:
                    - completely removed ActiveDirectory module dependancies from BF's code, and reimplemented in raw ADSI calls. Makes it fully portable, even into areas like Edge DMZ roles, where ADMS would never be installed.

                * 3:17 PM 8/23/2023 post Edge testing: some logic fixes; add: -Names param to filter on server names; -Site & supporting code, to permit lookup against sites *not* local to the local machine (and bypass lookup on the local machine) ; 
                    ren $Ex10siteDN -> $ExOPsiteDN; ren $Ex10configNC -> $ExopconfigNC
                * 1:03 PM 8/22/2023 minor cleanup
                * 10:31 AM 4/7/2023 added CBH expl of postfilter/sorting to draw predictable pattern 
                * 4:36 PM 4/6/2023 validated Psv51 & Psv20 and Ex10 & 16; added -Roles & -RoleNames params, to perform role filtering within the function (rather than as an external post-filter step). 
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
                                    if($env:ExchangeInstalled){
                                        $smsg += "`nand has `$env:ExchangeInstalled populated: Likely Edge Server" ;
                                        # unpop'd in native PS, only in EMS/REMS
                                    } elseif(($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup')) -or
                                        ($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup')) -or
                                        $ByPassLocalExchangeServerTest){
                                        $smsg +="`nand Reg confirms ExchangeServer\v1x\Setup (`$isLocalExchangeServer)" ; 
                                        if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole') -or
                                                (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'))
                                        {
                                            $smsg +="`nand Reg confirms \v1x\EdgeTransportRole (`$IsEdgeTransport)" ; 
                                            $IsEdgeTransport = $true
                                        } ; 
                                    }  ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                                    else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    # only exists in EMS/REMS, NOT PS raw
                                    if($vers = (get-item "$($env:ExchangeInstallPath)\Bin\Setup.exe").VersionInfo.FileVersionRaw ){
                                        <# [PS] C:\scripts>((get-item "$($env:ExchangeInstallPath)\Bin\Setup.exe").VersionInfo.FileVersionRaw )
                                        Major  Minor  Build  Revision
                                        -----  -----  -----  --------
                                        15     1      2507   39
                                        #>
                                    }else{
                                        if($binPath = (resolve-path  "$($env:ProgramFiles)\Microsoft\Exchange Server\V1*\Bin\Setup.exe" -ea 0).path){
                                            # find setup in stock path discovery (won't work if manual install non-std loc)                                            
                                        } else { 
                                            # loop the letter drives checking for progfiles copies
                                            (get-psdrive -PSProvider FileSystem |?{$_ -match '[D-Z]'}  | select -expand name)|foreach-object{
                                                $drv = $_ ; 
                                                if($binPath = (resolve-path  "$($drv)$($env:ProgramFiles.substring(1,($env:ProgramFiles.length-1)))\Microsoft\Exchange Server\V1*\Bin\Setup.exe" -ea 0).path){
                                                    break ; 
                                                } ; 
                                            };
                                            if($binPath){
                                                $vers = (get-item $binPath).VersionInfo.FileVersionRaw
                                            }else {

                                            } ;
                                        } ; 
                                    } ; 
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
        #endregion GET_ADEXCHANGESERVERTDO ;#*------^ END Function get-ADExchangeServerTDO ^------ ;

        #region load_ADMS  ; #*------v load-ADMS v------
        if(-not(gci function:load-ADMS -ea 0)){
            
            function load-ADMS {
                <#
                .NOTES
                REVISIONS   :
                * 12:57 PM 8/22/2023 test before calling Add-PSTitleBar (for PRY/dep-less support)
                .INPUTS
                None.
                .OUTPUTS
                Outputs $True/False load-status
                .EXAMPLE
                PS> $ADMTLoaded = load-ADMS ; Write-Debug "`$ADMTLoaded: $ADMTLoaded" ;
                .EXAMPLE
                PS> $ADMTLoaded = load-ADMS -Cmdlet get-aduser,get-adcomputer ; Write-Debug "`$ADMTLoaded: $ADMTLoaded" ;
                Load solely the specified cmdlets from ADMS
                .EXAMPLE
                # load ADMS
                PS> $reqMods+="load-ADMS".split(";") ;
                PS> if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
                PS> write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):(loading ADMS...)" ;
                PS> load-ADMS | out-null ;
                #load-ADMS -cmdlet get-aduser,Set-ADUser,Get-ADGroupMember,Get-ADDomainController,Get-ADObject,get-adforest | out-null ;
                Demo a load from the verb-ADMS.ps1 module, with opt specific -Cmdlet set
                .EXAMPLE
                PS> if(connect-ad){write-host 'connected'}else {write-warning 'unable to connect'}  ;
                Variant capturing & testing returned (returns true|false), using the alias name (if don't cap|eat return, you'll get a 'True' in console
                #>
                [CmdletBinding()]
                [Alias('connect-AD')]
                PARAM(
                    [Parameter(HelpMessage="Specifies an array of cmdlets that this cmdlet imports from the module into the current session. Wildcard characters are permitted[-Cmdlet get-aduser]")]
                    [ValidateNotNullOrEmpty()]$Cmdlet
                ) ;
                $Verbose = ($VerbosePreference -eq 'Continue') ;
                # focus specific cmdlet loads to SPEED them UP!
                $tMod = "ActiveDirectory" ;
                $ModsReg=Get-Module -Name $tMod -ListAvailable ;
                $ModsLoad=Get-Module -name $tMod ;
                $pltAD=@{Name=$tMod ; ErrorAction="Stop"; Verbose = ($VerbosePreference -eq 'Continue') } ;
                if($Cmdlet){$pltAD.add('Cmdlet',$Cmdlet) } ;
                if ($ModsReg) {
                    if (!($ModsLoad)) {
                        $env:ADPS_LoadDefaultDrive = 0 ;
                        import-module @pltAD;
                        if(get-command Add-PSTitleBar -ea 0){
                            Add-PSTitleBar 'ADMS' -verbose:$($VerbosePreference -eq "Continue") ;
                        } ; 
                        return $TRUE;
                    } else {
                        return $TRUE;
                    } # if-E ;
                } else {
                    Write-Error {"$((get-date).ToString('HH:mm:ss')):($env:computername) does not have AD Mgmt Tools installed!";};
                    return $FALSE
                } # if-E ;
            } ;
                       
        } ; 
        #endregion load_ADMS ; #*----------^END Function load-ADMS ^---------- 

        #region GET_GCFAST ; #*------v get-GCFast v------
        if(-not(gci function:get-GCFast -ea 0)){
            function get-GCFast {
                <#
                .NOTES
                REVISIONS   :
                * 2:39 PM 1/23/2025 added -exclude (exclude array of dcs by name), -ServerPrefix (exclude on leading prefix of name) params, added expanded try/catch, swapped out w-h etc for wlt calls
                .PARAMETER  Domain
                Which AD Domain [Domain fqdn]
                .PARAMETER  Site
                DCs from which Site name (defaults to AD lookup against local computer's Site)
                .PARAMETER Exclude
                Array of Domain controller names in target site/domain to exclude from returns (work around temp access issues)
                .PARAMETER ServerPrefix
                Prefix string to filter for, in returns (e.g. 'ABC' would only return DCs with name starting 'ABC')
                .PARAMETER SpeedThreshold
                Threshold in ms, for AD Server response time(defaults to 100ms)
                .INPUTS
                None. Does not accepted piped input.
                .OUTPUTS
                Returns one DC object, .Name is name pointer
                .EXAMPLE
                PS> get-gcfast -domain dom.for.domain.com -site Site
                Lookup a Global domain gc, with Site specified (whether in Site or not, will return remote site dc's)
                .EXAMPLE
                PS> get-gcfast -domain dom.for.domain.com
                Lookup a Global domain gc, default to Site lookup from local server's perspective
                .EXAMPLE    
                PS> if($domaincontroller = get-gcfast -Exclude ServerBad -Verbose){
                PS>     write-warning "Changing DomainControler: Waiting 20seconds, for RelSync..." ;
                PS>     start-sleep -Seconds 20 ;
                PS> } ; 
                Demo acquireing a new DC, excluding a caught bad DC, and waiting before moving on, to permit ADRerplication from prior dc to attempt to ensure full sync of changes. 
                PS> get-gcfast -ServerPrefix ABC -verbose
                Demo use of -ServerPrefix to only return DCs with servernames that begin with the string 'ABC'
                .EXAMPLE
                PS> $adu=$null ;
                PS> $Exit = 0 ;
                PS> Do {
                PS>     TRY {
                PS>         $adu = get-aduser -id $rmbx.DistinguishedName -server $domainController -Properties $adprops -ea 0| select $adprops ;
                PS>         $Exit = $DoRetries ;
                PS>     }CATCH [System.Management.Automation.RuntimeException] {
                PS>         if ($_.Exception.Message -like "*ResourceUnavailable*") {
                PS>             $ErrorTrapped=$Error[0] ;
                PS>             $smsg = "Failed to exec cmd because: $($ErrorTrapped.Exception.Message )" ;
                PS>             if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                PS>             else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                PS>             # re-quire a new DC
                PS>             $badDC = $domaincontroller ; 
                PS>             $smsg = "PROBLEM CONTACTING $(domaincontroller)!:Resource unavailable: $($ErrorTrapped.Exception.Message)" ; 
                PS>             $smsg += "get-GCFast() an alterate DC" ; 
                PS>             if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                PS>             else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                PS>             if($domaincontroller = get-gcfast -Exclude $$badDC -Verbose){
                PS>                 write-warning "Changing DomainController:($($badDC)->$($domaincontroller)):Waiting 20seconds, for ReplSync..." ;
                PS>                 start-sleep -Seconds 20 ;
                PS>             } ;                             
                PS>         }else {
                PS>             throw $Error[0] ;
                PS>         } ; 
                PS>     } CATCH {
                PS>         $ErrorTrapped=$Error[0] ;
                PS>         Start-Sleep -Seconds $RetrySleep ;
                PS>         $Exit ++ ;
                PS>         $smsg = "Failed to exec cmd because: $($ErrorTrapped)" ;
                PS>         $smsg += "`nTry #: $Exit" ;
                PS>         if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                PS>         else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                PS>         If ($Exit -eq $DoRetries) {
                PS>             $smsg =  "Unable to exec cmd!" ;
                PS>             if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                PS>             else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                PS>         } ;
                PS>         Continue ;
                PS>     }  ;
                PS> } Until ($Exit -eq $DoRetries) ;
                Retry demo that includes aquisition of a new DC, excluding a caught bad DC, and waiting before moving on, to permit ADRerplication from prior dc to attempt to ensure full sync of changes. 
                #>
                [CmdletBinding()]
                PARAM(
                    [Parameter(Position = 0, Mandatory = $False, HelpMessage = "Optional: DCs from what Site name? (default=Discover)")]
                        [string]$Site,
                    [Parameter(HelpMessage = 'Target AD Domain')]
                        [string]$Domain,
                    [Parameter(HelpMessage = 'Array of Domain controller names in target site/domain to exclude from returns (work around temp access issues)')]
                        [string[]]$Exclude,
                    [Parameter(HelpMessage = "Prefix string to filter for, in returns (e.g. 'ABC' would only return DCs with name starting 'ABC')")]
                        [string]$ServerPrefix,
                    [Parameter(HelpMessage = 'Threshold in ms, for AD Server response time(defaults to 100ms)')]
                        $SpeedThreshold = 100
                ) ;
                $Verbose = $($PSBoundParameters['Verbose'] -eq $true)
                $SpeedThreshold = 100 ;
                $rgxSpbDCRgx = 'CN=EDCMS'
                $ErrorActionPreference = 'SilentlyContinue' ; # Set so we don't see errors for the connectivity test
                $env:ADPS_LoadDefaultDrive = 0 ; 
                $sName = "ActiveDirectory"; 
                TRY{
                    if ( -not(Get-Module | Where-Object { $_.Name -eq $sName }) ) {
                        $smsg = "Adding ActiveDirectory Module (`$script:ADPSS)" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $script:AdPSS = Import-Module $sName -PassThru -ea Stop ;
                    } ;
                    if (-not $Domain) {
                        $Domain = (get-addomain -ea Stop).DNSRoot ; # use local domain
                        $smsg = "Defaulting domain: $Domain";
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    }
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
    
                # Get all the local domain controllers
                if ((-not $Site)) {
                    # if no site, look the computer's Site Up in AD
                    TRY{
                        $Site = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name ;
                        $smsg = "Using local machine Site: $($Site)";
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 
                } ;

                # gc filter
                #$LocalDCs = Get-ADDomainController -filter { (isglobalcatalog -eq $true) -and (Site -eq $Site) } ;
                # ISSUE: ==3:26 pm 3/7/2024: NO LOCAL SITE DC'S IN SPB
                # os: LOGONSERVER=\\EDCMS8100
                TRY{
                    $LocalDCs = Get-ADDomainController -filter { (isglobalcatalog -eq $true) -and (Site -eq $Site) -and (Domain -eq $Domain) } -ErrorAction STOP
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
                if( $LocalDCs){
                    $smsg = "`Discovered `$LocalDCs:`n$(($LocalDCs|out-string).trim())" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } elseif($Site -eq 'Spellbrook'){
                    $smsg = "Get-ADDomainController -filter { (isglobalcatalog -eq `$true) -and (Site -eq $($Site)) -and (Domain -eq $($Domain)}"
                    $smsg += "`nFAILED to return DCs, and `$Site -eq Spellbrook:" 
                    $smsg += "`ndiverting to $($rgxSpbDCRgx) dcs in entire Domain:" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    TRY{
                        $LocalDCs = Get-ADDomainController -filter { (isglobalcatalog -eq $true) -and (Domain -eq $Domain) } -EA STOP | 
                            ?{$_.ComputerObjectDN -match $rgxSpbDCRgx } 
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 
                } ; 
  
                # any dc filter
                #$LocalDCs = Get-ADDomainController -filter {(Site -eq $Site)} ;

                $PotentialDCs = @() ;
                # Check connectivity to each DC against $SpeedThreshold
                if ($LocalDCs) {
                    foreach ($LocalDC in $LocalDCs) {
                        $TCPClient = New-Object System.Net.Sockets.TCPClient ;
                        $Connect = $TCPClient.BeginConnect($LocalDC.Name, 389, $null, $null) ;
                        $Wait = $Connect.AsyncWaitHandle.WaitOne($SpeedThreshold, $False) ;
                        if ($TCPClient.Connected) {
                            $PotentialDCs += $LocalDC.Name ;
                            $Null = $TCPClient.Close() ;
                        } # if-E
                    } ;
                    if($Exclude){
                        $smsg = "-Exclude specified:`n$((($exclude -join ',')|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        foreach($excl in $Exclude){
                            $PotentialDCs = $PotentialDCs |?{$_ -ne $excl} ; 
                        } ; 
                    } ; 
                    if($ServerPrefix){
                        $smsg = "-ServerPrefix specified: $($ServerPrefix)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        $PotentialDCs = $PotentialDCs |?{$_ -match "^$($ServerPrefix)" } ; 
            
                    }
                    write-host -foregroundcolor yellow  
                    $smsg = "`$PotentialDCs: $PotentialDCs";
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $DC = $PotentialDCs | Get-Random ;

                    $smsg = "(returning random domaincontroller from result to pipeline:$($DC)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $DC | write-output  ;
                } else {
                    write-host -foregroundcolor yellow  "NO DCS RETURNED BY GET-GCFAST()!";
                    write-output $false ;
                } ;
            }  ; 
        } ; 
        #endregion GET_GCFAST ; #*------^ END get-GCFast ^------

        #region RESOLVE_NETWORKLOCALTDO ; #*------v resolve-NetworkLocalTDO v------
        #*------v Function resolve-NetworkLocalTDO v------
        if(-not(gci function:resolve-NetworkLocalTDO -ea 0)){
            Function resolve-NetworkLocalTDO {
                <#
                .SYNOPSIS
                resolve-NetworkLocalTDO.ps1 - Retrieve local network interface descriptors, and resolved ip address PTR -> A Record FQDN
                .NOTES
                Version     : 0.0.1
                Author      : Todd Kadrie
                Website     : http://www.toddomation.com
                Twitter     : @tostka / http://twitter.com/tostka
                CreatedDate : 2025-04-28
                FileName    : resolve-NetworkLocalTDO.ps1
                License     : MIT License
                Copyright   : (c) 2025 Todd Kadrie
                Github      : https://github.com/tostka/verb-XXX
                Tags        : Powershell
                AddedCredit : REFERENCE
                AddedWebsite: URL
                AddedTwitter: URL
                REVISIONS
                .DESCRIPTION
                resolve-NetworkLocalTDO.ps1 - Retrieve local network interface descriptors, and resolved ip address PTR -> A Record FQDN                
                .INPUTS
                None. Does not accepted piped input.(.NET types, can add description)
                .OUTPUTS
                System.PsCustomObject summary of useful Nic descriptors                
                .EXAMPLE
                PS> $netsettings = resolve-NetworkLocalTDO ; 
                Demo run
                .LINK
                https://github.com/tostka/verb-Network
                #>                
                [CmdletBinding()]
                Param () ;
                BEGIN{
                    $rgxIP4Addr = "(?:\d{1,3}\.){3}\d{1,3}" ;
                    $rgxIP6Addr = "^((([0-9A-Fa-f]{1,4}:){1,6}:)|(([0-9A-Fa-f]{1,4}:){7}))([0-9A-Fa-f]{1,4})$" ; 
                    $rgxIP4AddrAuto = "169\.254\.\d{1,3}\.\d{1,3}" ;  
                    $prpNS = 'DNSHostName','ServiceName',@{N="DNSServerSearchOrder";E={"$($_.DNSServerSearchOrder)"}}, 
                        @{N='IPAddress';E={$_.IPAddress}},@{N='DefaultIPGateway';E={$_.DefaultIPGateway}} ;
                } ; 
                PROCESS {
                    $nicsettings = [ordered]@{ DNSHostName = $null ;  ServiceName = $null ;  DNSServerSearchOrder = $null ;  IPAddress = $null ;  DefaultIPGateway = $null ;  Fqdn = $null ;  }  ;                    
                    TRY{
                        $ret = Get-WMIObject Win32_NetworkAdapterConfiguration -Computername localhost -ea STOP|
                             ? {$_.IPEnabled -match "True"} | Select -property $prpNS ; 
                        $nicsettings.DNSHostName = $ret.DNSHostName; 
                        $nicsettings.ServiceName = $ret.ServiceName;  
                        $nicsettings.DNSServerSearchOrder = $ret.DNSServerSearchOrder;  
                        $nicsettings.IPAddress = $ret.IPAddress;  
                        $nicsettings.DefaultIPGateway = $ret.DefaultIPGateway;  
                        if($nicSettings.ipaddress | ?{$_ -MATCH $rgxIP4Addr -AND $_ -notmatch $rgxIP4AddrAuto} ){
                            $nicsettings.fqdn = (resolve-dnsname -name ($nicSettings.ipaddress | ?{$_ -MATCH $rgxIP4Addr -AND $_ -notmatch $rgxIP4AddrAuto} ) -type ptr).namehost | select -first 1 ;   
                        } ; 
                        [pscustomobject]$nicsettings | write-output ; 
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;                     
                } ; 
            }
        } ; 
        #*------^ END Function resolve-NetworkLocalTDO ^------
        #endregion RESOLVE_NETWORKLOCALTDO ; #*------^ END resolve-NetworkLocalTDO ^------

        #region OUT_CLIPBOARD ; #*------v out-Clipboard v------
        if(-not(gci function:out-Clipboard -ea 0)){
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
        } ; 
        #endregion OUT_CLIPBOARD ; #*------^ END out-Clipboard ^------

        #region CONVERTFROM_MARKDOWNTABLE ; #*------v convertFrom-MarkdownTable v------
        if(-not(gci function:convertFrom-MarkdownTable -ea 0)){
            Function convertFrom-MarkdownTable {
                <#
                .SYNOPSIS
                convertFrom-MarkdownTable.ps1 - Converts a Markdown table to a PowerShell object.
                .NOTES
                REVISION
                * 9:33 AM 4/11/2025 add alias: cfmdt (reflects standard verbalias)
                .PARAMETER markdowntext
                Markdown-formated table to be converted into an object [-markdowntext 'title text']
                .INPUTS
                Accepts piped input.
                .OUTPUTS
                System.Object[]
                .EXAMPLE
                PS> $svcs = Get-Service Bits,Winrm | select status,name,displayname | 
                    convertTo-MarkdownTable -border | ConvertFrom-MarkDownTable ;  
                Convert Service listing to and back from MD table, demo's working around border md table syntax (outter pipe-wrapped lines)
                .EXAMPLE
                PS> $mdtable = @"
                |EmailAddress|DisplayName|Groups|Ticket|
                |---|---|---|---|
                |da.pope@vatican.org||CardinalDL@vatican.org|999999|
                |bozo@clown.com|Bozo Clown|SillyDL;SmartDL|000001|
                "@ ; 
                    $of = ".\out-csv-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ; 
                    $mdtable | convertfrom-markdowntable | export-csv -path $of -notype ;
                    cat $of ;

                    "EmailAddress","DisplayName","Groups","Ticket"
                    "da.pope@vatican.org","","CardinalDL@vatican.org","999999"
                    "bozo@clown.com","Bozo Clown","SillyDL;SmartDL","000001"

                Example simpler method for building csv input files fr mdtable syntax, without PSCustomObjects, hashes, or invoked object creation.
                .EXAMPLE
                PS> $mdtable | convertFrom-MarkdownTable | convertTo-MarkdownTable -border ; 
                Example to expand and dress up a simple md table, leveraging both convertfrom-mtd and convertto-mtd (which performs space padding to align pipe columns)
                .LINK
                https://github.com/tostka/verb-IO
                #>                
                [CmdletBinding()]
                [alias('convertfrom-mdt','in-markdowntable','in-mdt','cfmdt')]    
                Param (
                    [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Markdown-formated table to be converted into an object [-markdowntext 'title text']")]
                        $markdowntext
                ) ;
                PROCESS {
                    $content = @() ; 
                    if(($markdowntext|measure).count -eq 1){$markdowntext  = $markdowntext -split '\n' } ;
                    $markdowntext  = $markdowntext -replace '\|\|','| |' ; 
                    $content = $markdowntext  | ?{$_ -notmatch "--" } ;
                } ;  
                END {
                    $PsObj = $content.trim('|').trimend('|')| where-object{$_} | ForEach-Object{ 
                        ($_.split('|') | where-object{$_} | foreach-object{$_.trim()} |where-object{$_} )  -join '|' ; 
                    } | ConvertFrom-Csv -Delimiter '|'; # convert to object
                    $PsObj | write-output ; 
                } ; 
            } ;             
        } ; 
        #endregion CONVERTFROM_MARKDOWNTABLE ; #*------^ END convertFrom-MarkdownTable ^------

        #region REMOVE_INVALIDVARIABLENAMECHARS ; #*------v Remove-InvalidVariableNameChars v------
        #*------v Function Remove-InvalidVariableNameChars v------
        if(-not (gcm Remove-InvalidVariableNameChars -ea 0)){
            Function Remove-InvalidVariableNameChars ([string]$Name) {
                ($Name.tochararray() -match '[A-Za-z0-9_]') -join '' | write-output ;
            };
        } ;
        #*------^ END Function Remove-InvalidVariableNameChars ^------
        #endregion REMOVE_INVALIDVARIABLENAMECHARS ; #*------^ END Remove-InvalidVariableNameChars ^------

        #region REMOVE_SMTPPLUSADDRESS ; #*------v remove-SmtpPlusAddress v------
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
        #endregion REMOVE_SMTPPLUSADDRESS ; #*------^ END remove-SmtpPlusAddress ^------

        #region INITIALIZE_XOPEVENTIDTABLE ; #*------v Initialize-xopEventIDTable v------
        if(-not(gci function:Initialize-xopEventIDTable -ea 0)){
            function Initialize-xopEventIDTable {
                <#
                .SYNOPSIS
                Initialize-xopEventIDTable - Builds an indexed hash tabl of Exchange Server Get-MessageTrackingLog EventIDs
                .NOTES
                Version     : 1.0.0
                Author      : Todd Kadrie
                Website     : http://www.toddomation.com
                Twitter     : @tostka / http://twitter.com/tostka
                CreatedDate : 2025-04-22
                FileName    : Initialize-xopEventIDTable
                License     : (none asserted)
                Copyright   : (none asserted)
                Github      : https://github.com/tostka/verb-Ex2010
                Tags        : Powershell,EmailAddress,Version
                AddedCredit : Bruno Lopes (brunokktro )
                AddedWebsite: https://www.linkedin.com/in/blopesinfo
                AddedTwitter: @brunokktro / https://twitter.com/brunokktro
                REVISIONS
                * 2;58 pm 4/28/2025 Updated table again, and found Ex2016/19 eventid specifications online, added. Did find that 
                the online doc doesn't document the edge SendExternal event id (added below, manually).             
                * 1:47 PM 7/9/2024 CBA github field correction
                * 1:22 PM 5/22/2024init
                .DESCRIPTION
                Initialize-xopEventIDTable - Builds an indexed hash tabl of Exchange Server Get-MessageTrackingLog EventIDs

                ## Exchange 2019 EventID reference:

                [Event types in the message tracking log | Microsoft Learn](https://learn.microsoft.com/en-us/exchange/mail-flow/transport-logs/message-tracking?view=exchserver-2019#event-types-in-the-message-tracking-log)

                Doesn't include Edge eventid: 
                "SENDEXTERNAL          | A message was sent by SMTP to sent to the SMTP server responsible to receive the email for the external email address."
            
                (needs to be manually spliced in below 'SEND' during updates from source MS documentation)
            
                .OUTPUT
                System.Collections.Hashtable returns an Indexed Hash of EventIDs EventName to Description
                .EXAMPLE
                PS> $eventIDLookupTbl = Initialize-EventIDTable ; 
                PS> $smsg = "`n`n## EventID Definitions:" ; 
                PS> $TrackMsgs | group eventid | select -expand Name | foreach-object{                   
                PS>     $smsg += "`n$(($eventIDLookupTbl[$_] | ft -hidetableheaders |out-string).trim())" ; 
                PS> } ; 
                PS> $smsg += "`n`n" ; 
                PS> if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
                PS> else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Demo resolving histogram eventid uniques, to MS documented meansings of each event id in the msgtrack.
                .EXAMPLE
                ps> Initialize-xopEventIDTable -EmailAddress 'monitoring+SolarWinds@toro.com;notanemailaddresstoro.com,todd+spam@kadrie.net' -verbose ;
                PS> 
                Demo with comma and semicolon delimiting, and an invalid address (to force a regex match fail error).
                .LINK
                https://github.com/brunokktro/EmailAddress/blob/master/Get-ExchangeEnvironmentReport.ps1
                .LINK
                https://github.com/tostka/verb-Ex2010
                #>
                [CmdletBinding()]
                #[Alias('rvExVers')]
                PARAM() ;
                BEGIN {
                    ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
                    $verbose = $($VerbosePreference -eq "Continue")
                    $rgxSMTPAddress = "([0-9a-zA-Z]+[-._+&='])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}" ; 
                    $sBnr="#*======v $($CmdletName): v======" ;
                    write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
                
                    $eventIDsMD = @"
EventName             | Description                                                                                                                                                                                                                                                                                                                                                                                                                                                                            
--------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
AGENTINFO             | This event is used by transport agents to log custom data.                                                                                                                                                                                                                                                                                                                                                                                                                             
BADMAIL               | A message submitted by the Pickup directory or the Replay directory that can't be delivered or returned.                                                                                                                                                                                                                                                                                                                                                                               
CLIENTSUBMISSION      | A message was submitted from the Outbox of a mailbox.                                                                                                                                                                                                                                                                                                                                                                                                                                  
DEFER                 | Message delivery was delayed.                                                                                                                                                                                                                                                                                                                                                                                                                                                          
DELIVER               | A message was delivered to a local mailbox.                                                                                                                                                                                                                                                                                                                                                                                                                                            
DELIVERFAIL           | An agent tried to deliver the message to a folder that doesn't exist in the mailbox.                                                                                                                                                                                                                                                                                                                                                                                                   
DROP                  | A message was dropped without a delivery status notification (also known as a DSN, bounce message, non-delivery report, or NDR). For example:<br/>- Completed moderation approval request messages.<br/>- Spam messages that were silently dropped without an NDR.                                                                                                                                                                                                                     
DSN                   | A delivery status notification (DSN) was generated.                                                                                                                                                                                                                                                                                                                                                                                                                                    
DUPLICATEDELIVER      | A duplicate message was delivered to the recipient. Duplication may occur if a recipient is a member of multiple nested distribution groups. Duplicate messages are detected and removed by the information store.                                                                                                                                                                                                                                                                     
DUPLICATEEXPAND       | During the expansion of the distribution group, a duplicate recipient was detected.                                                                                                                                                                                                                                                                                                                                                                                                    
DUPLICATEREDIRECT     | An alternate recipient for the message was already a recipient.                                                                                                                                                                                                                                                                                                                                                                                                                        
EXPAND                | A distribution group was expanded.                                                                                                                                                                                                                                                                                                                                                                                                                                                     
FAIL                  | Message delivery failed. Sources include SMTP, DNS, QUEUE, and ROUTING.                                                                                                                                                                                                                                                                                                                                                                                                                
HADISCARD             | A shadow message was discarded after the primary copy was delivered to the next hop. For more information, see Shadow redundancy in Exchange Server.                                                                                                                                                                                                                                                                                                                                   
HARECEIVE             | A shadow message was received by the server in the local database availability group (DAG) or Active Directory site.                                                                                                                                                                                                                                                                                                                                                                   
HAREDIRECT            | A shadow message was created.                                                                                                                                                                                                                                                                                                                                                                                                                                                          
HAREDIRECTFAIL        | A shadow message failed to be created. The details are stored in the source-context field.                                                                                                                                                                                                                                                                                                                                                                                             
INITMESSAGECREATED    | A message was sent to a moderated recipient, so the message was sent to the arbitration mailbox for approval. For more information, see Manage message approval.                                                                                                                                                                                                                                                                                                                       
LOAD                  | A message was successfully loaded at boot.                                                                                                                                                                                                                                                                                                                                                                                                                                             
MODERATIONEXPIRE      | A moderator for a moderated recipient never approved or rejected the message, so the message expired. For more information about moderated recipients, see Manage message approval.                                                                                                                                                                                                                                                                                                    
MODERATORAPPROVE      | A moderator for a moderated recipient approved the message, so the message was delivered to the moderated recipient.                                                                                                                                                                                                                                                                                                                                                                   
MODERATORREJECT       | A moderator for a moderated recipient rejected the message, so the message wasn't delivered to the moderated recipient.                                                                                                                                                                                                                                                                                                                                                                
MODERATORSALLNDR      | All approval requests sent to all moderators of a moderated recipient were undeliverable, and resulted in non-delivery reports (also known as NDRs or bounce messages).                                                                                                                                                                                                                                                                                                                
NOTIFYMAPI            | A message was detected in the Outbox of a mailbox on the local server.                                                                                                                                                                                                                                                                                                                                                                                                                 
NOTIFYSHADOW          | A message was detected in the Outbox of a mailbox on the local server, and a shadow copy of the message needs to be created.                                                                                                                                                                                                                                                                                                                                                           
POISONMESSAGE         | A message was put in the poison message queue or removed from the poison message queue.                                                                                                                                                                                                                                                                                                                                                                                                
PROCESS               | The message was successfully processed.                                                                                                                                                                                                                                                                                                                                                                                                                                                
PROCESSMEETINGMESSAGE | A meeting message was processed by the Mailbox Transport Delivery service.                                                                                                                                                                                                                                                                                                                                                                                                             
RECEIVE               | A message was received by the SMTP receive component of the transport service or from the Pickup or Replay directories (source: SMTP), or a message was submitted from a mailbox to the Mailbox Transport Submission service (source: STOREDRIVER).                                                                                                                                                                                                                                    
REDIRECT              | A message was redirected to an alternative recipient after an Active Directory lookup.                                                                                                                                                                                                                                                                                                                                                                                                 
RESOLVE               | A message's recipients were resolved to a different email address after an Active Directory lookup.                                                                                                                                                                                                                                                                                                                                                                                    
RESUBMIT              | A message was automatically resubmitted from Safety Net. For more information, see Safety Net in Exchange Server.                                                                                                                                                                                                                                                                                                                                                                      
RESUBMITDEFER         | A message resubmitted from Safety Net was deferred.                                                                                                                                                                                                                                                                                                                                                                                                                                    
RESUBMITFAIL          | A message resubmitted from Safety Net failed.                                                                                                                                                                                                                                                                                                                                                                                                                                          
SEND                  | A message was sent by SMTP between transport services.                                                                                                                                                                                                                                                                                                                                                                                                                                 
SENDEXTERNAL          | A message was sent by SMTP to sent to the SMTP server responsible to receive the email for the external email address.                                                                                                                                                                                                                                                                                                                                                                                                                           
SUBMIT                | The Mailbox Transport Submission service successfully transmitted the message to the Transport service. For SUBMIT events, the source-context property contains the following details:<br/>- MDB: The mailbox database GUID.<br/>- Mailbox: The mailbox GUID.<br/>- Event: The event sequence number.<br/>- MessageClass: The type of message. For example, IPM.Note.<br/>- CreationTime: Date-time of the message submission.<br/>- ClientType: For example, User, OWA, or ActiveSync.
SUBMITDEFER           | The message transmission from the Mailbox Transport Submission service to the Transport service was deferred.                                                                                                                                                                                                                                                                                                                                                                          
SUBMITFAIL            | The message transmission from the Mailbox Transport Submission service to the Transport service failed.                                                                                                                                                                                                                                                                                                                                                                                
SUPPRESSED            | The message transmission was suppressed.                                                                                                                                                                                                                                                                                                                                                                                                                                               
THROTTLE              | The message was throttled.                                                                                                                                                                                                                                                                                                                                                                                                                                                             
TRANSFER              | Recipients were moved to a forked message because of content conversion, message recipient limits, or agents. Sources include ROUTING or QUEUE.
"@ ; 
                    # UPDATE NOTE: MANUAL UNDOCUMENTED ADDITION: "SENDEXTERNAL          | A message was sent by SMTP to sent to the SMTP server responsible to receive the email for the external email address."
                    # (needs to be manually spliced in below 'SEND' during updates from source MS documentation)
                    $Object = $eventIDsMD | convertfrom-MarkdownTable ; 
                    $Key = 'EventName' ; 
                    $Hashtable = @{}
                }
                PROCESS {
                    Foreach ($Item in $Object){
                        $Procd++ ; 
                        $Hashtable[$Item.$Key.ToString()] = $Item ; 
                        if($ShowProgress -AND ($Procd -eq $Every)){
                            write-host -NoNewline '.' ; $Procd = 0 
                        } ; 
                    } ;                 
                } # PROC-E
                END{
                    $Hashtable | write-output ; 
                    write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
                }
            };
        } ;      
        #endregion INITIALIZE_XOPEVENTIDTABLE ; #*------^ Initialize-xopEventIDTable ^------

        #endregion FUNCTIONS_INTERNAL ; #*======^ END FUNCTIONS_INTERNAL ^======

        #region CONSTANTS_AND_ENVIRO ; #*======v CONSTANTS_AND_ENVIRO v======
        #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
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
        # splatted resolve-EnvironmentTDO CALL: 
        $pltRvEnv=[ordered]@{
            PSCmdletproxy = $rPSCmdlet ; 
            PSScriptRootproxy = $rPSScriptRoot ; 
            PSCommandPathproxy = $rPSCommandPath ; 
            MyInvocationproxy = $rMyInvocation ;
            PSBoundParametersproxy = $rPSBoundParameters
            verbose = [boolean]($PSBoundParameters['Verbose'] -eq $true) ; 
        } ;
        write-verbose "(Purge no value keys from splat)" ; 
        $mts = $pltRVEnv.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltRVEnv.remove($_.Name)} ; rv mts -ea 0 ; 
        $smsg = "resolve-EnvironmentTDO w`n$(($pltRVEnv|out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $rvEnv = resolve-EnvironmentTDO @pltRVEnv ; 
        $smsg = "`$rvEnv returned:`n$(($rvEnv |out-string).trim())" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
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
        #region TLS_LATEST_FORCE ; #*------v TLS_LATEST_FORCE v------
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
        #endregion TLS_LATEST_FORCE ; #*------^ END TLS_LATEST_FORCE ^------

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
        # - AADU Licensing group checks
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

        $DaysLimit = 30 ; # technically no specific limit to Get-MessageTrackingLog, but practical matter they're limited to 30d on the drive
        $rgxIsPlusAddrSmtpAddr = "[+].*@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}" ; 
        $prpMTLfta = 'Timestamp','EventId','Sender','Recipients','MessageSubject' ; 
        $prpXCsv = "Timestamp",@{N='TimestampLocal'; E={$_.Timestamp.ToLocalTime()}},"Source","EventId","RelatedRecipientAddress","Sender",@{N='Recipients'; E={$_.Recipients}},"RecipientCount",@{N='RecipientStatus'; E={$_.RecipientStatus}},"MessageSubject","TotalBytes",@{N='Reference'; E={$_.Reference}},"MessageLatency","MessageLatencyType","InternalMessageId","MessageId","ReturnPath","ClientIp","ClientHostname","ServerIp","ServerHostname","ConnectorId","SourceContext","MessageInfo",@{N='EventData'; E={$_.EventData}} ;
        $prpMTFailFL = 'Timestamp','ClientHostname','Source','EventId','Recipients','RecipientStatus','MessageSubject','ReturnPath' ;
        $s24HTimestamp = 'yyyyMMdd-HHmm'
        $sFiletimestamp =  $s24HTimestamp
        #region LOCAL_CONSTANTS ; #*------^ END LOCAL_CONSTANTS ^------

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

        # 5:03 PM 4/22/2025 MOVED FUNCTIONS UP TO TOP OF BEGIN, NOTHING IS AVAIL, UNLESS PREDECLARED, ON CURLY

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
            

        #region START_LOG_OPTIONS #*======v START_LOG_OPTIONS v======
        $useSLogHOl = $true ; # one or 
        $useSLogSimple = $false ; #... the other
        $useTransName = $false ; # TRANSCRIPTNAME
        $useTransPath = $false ; # TRANSCRIPTPATH
        $useTransRotate = $false ; # TRANSCRIPTPATHROTATE
        $useStartTrans = $false ; # STARTTRANS
        #region START_LOG_HOLISTIC #*------v START_LOG_HOLISTIC v------
        if($useSLogHOl){
            # Single log for script/function example that accomodates detect/redirect from AllUsers scope'd installed code, and hunts a series of drive letters to find an alternate logging dir (defers to profile variables)
            #${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ;
            if(-not (get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
            foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
            if(-not (get-variable rgxPSAllUsersScope -ea 0)){
                $rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps(((d|m))*)1|dll)$" ;
            } ;
            if(-not (get-variable rgxPSCurrUserScope -ea 0)){
                $rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps((d|m)*)1|dll)$" ;
            } ;
            $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($whatif) ;} ;
            # if using [CmdletBinding(SupportsShouldProcess)] + -WhatIf:$($WhatIfPreference):
            #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
            #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag="$($ticket)-$($TenOrg)-LASTPASS-" ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
            #$pltSL.Tag = $ModuleName ; 
            #$pltSL.Tag = "$($ticket)-$($usr)" ; 
            #$pltSL.Tag = $((@($ticket,$usr) |?{$_}) -join '-')
            if($ticket){$pltSL.Tag = $ticket} ;
            <#
            if($rPSBoundParameters.keys){ # alt: leverage $rPSBoundParameters hash
                $sTag = @() ; 
                #$pltSL.TAG = $((@($rPSBoundParameters.keys) |?{$_}) -join ','); # join all params
                if($rPSBoundParameters['Summary']){ $sTag+= @('Summary') } ; # build elements conditionally, string
                if($rPSBoundParameters['Number']){ $sTag+= @("Number$($rPSBoundParameters['Number'])") } ; # and keyname,value
                $pltSL.Tag = $sTag -join ',' ; 
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
                } ; 
            } ; 
            if(-not $rvEnv.isFunc){
                # under funcs, this is the scriptblock of the func, not a path
                if($rvEnv.MyInvocationproxy.MyCommand.Definition){$prxPath2 = $rvEnv.MyInvocationproxy.MyCommand.Definition }
                elseif($rvEnv.MyInvocationproxy.MyCommand.Definition){$prxPath2 = $rvEnv.MyInvocationproxy.MyCommand.Definition } ; 
            } ; 
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
                    $startResults = start-Transcript -path $transcript ;
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
            } ;
        } ; 
        #endregion START_LOG_HOLISTIC #*------^ END START_LOG_HOLISTIC ^------
        #...
        #endregion START_LOG_OPTIONS #*======^ START_LOG_OPTIONS ^======

        #region NETWORK_INFO ; #*======v NETWORK_INFO v======
        $netsettings = resolve-NetworkLocalTDO ; 
        #endregion NETWORK_INFO ; #*======^ END NETWORK_INFO ^======

        #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======
        # PRETUNE STEERING separately *before* pasting in balance of region
        # THIS BLOCK DEPS ON VERB-* FANCY CRED/AUTH HANDLING MODULES THAT *MUST* BE INSTALLED LOCALLY TO FUNCTION
        # NOTE: *DOES* INCLUDE *PARTIAL* DEP-LESS $useExopNoDep=$true OPT THAT LEVERAGES Connect-ExchangeServerTDO, VS connect-ex2010 & CREDS ARE ASSUMED INHERENT TO THE ACCOUNT) 
        # Connect-ExchangeServerTDO HAS SUBSTANTIAL BENEFIT, OF WORKING SEAMLESSLY ON EDGE SERVER AND RANGE OF DOMAIN-=CONNECTED EXOP ROLES
        $useO365 = $false ;
        $useEXO = $false ; 
        $UseOP=$true ; 
        $UseExOP=$true ;
        $useExopNoDep = $true ; # switch to use Connect-ExchangeServerTDO, vs connect-ex2010 (creds are assumed inherent to the account)
        $ExopVers = 'Ex2010' # 'Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000', Null for All versions
        if($Version){
            $ExopVers = $Version ; #defer to local script $version if set
        } ; 
        $useForestWide = $true ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
        $UseOPAD = $false ; 
        $UseMSOL = $false ; # should be hard disabled now in o365
        $UseAAD = $false  ; 
        if($env:userdomain -eq $env:computername){
            $isNonDomainServer = $true ; 
            $UseOPAD = $false ; 
        } 
        if(get-service MSExchangeTransport -ea 0){
            $UseExOP = $true ; 
            if( ($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup')) -or
                    ($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup')) -or $ByPassLocalExchangeServerTest)
                {
                    if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole') -or
                            (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole')) {
                        $smsg = "We are on Exchange Edge Transport Server"
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $IsEdgeTransport = $true ; 
                        if($IsEdgeTransport -AND $psise){
                            $smsg = "powershell_ISE UNDER Exchange Edge Transport role!" 
                            $smsg += "`nThis script is likely to fail the get-messagetrackingLog calls etc with Access Denied errors"
                            $smsg += "`nif run with this combo."
                            $smsg += "`nEXIT POWERSHELL ISE, AND RUN THIS DIRECTLY UNDER EMS FOR EDGE USE"; 
                            $smsg += "`n(bug appears to be a conflict in Remote EMS v EMS access permissions, not resolved yet)" ; 
                            write-warning $smsg ; 
                        } ; 
                    } else {
                        $IsEdgeTransport = $false ; 
                    } ; 
                } ; 
        }else{
            $isLocalExchangeServer = $false ; 
        } ; 
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
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            if((get-command get-TenantTag).Parameters.keys -contains 'silent'){
                $TenOrg = get-TenantTag -Credential $Credential -silent ;;
            }else {
                $TenOrg = get-TenantTag -Credential $Credential ;
            }
        } else { 
            # if not using Credentials or a TargetTenants/TenOrg loop, default the $TenOrg on the $env:USERDOMAIN
            $smsg = "(unconfigured `$TenOrg & *NO* `$Credential: fallback asserting from `$env:USERDOMAIN)" ; 
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            switch -regex ($env:USERDOMAIN){
                ([regex]('(' + (( @($TORMeta.legacyDomain,$CMWMeta.legacyDomain)  |foreach-object{[regex]::escape($_)}) -join '|') + ')')).tostring() {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
                $TOLMeta.legacyDomain {$TenOrg = 'TOL' }
                $env:COMPUTERNAME {
                    # non-domain-joined, no domain, but the $netsettings.fqdn has a dns suffix that can be steered.
                    if($netsettings.fqdn){
                    switch -regex (($netsettings.fqdn.split('.') | select -last 2 ) -join '.'){
                      'toro\.com$' {$tenorg = 'TOR' ; } ; 
                      '(charlesmachineworks\.com|cmw\.internal)$' { $TenOrg = 'CMW'} ; 
                      '(torolab\.com|snowthrower\.com)$'  { $TenOrg = 'TOL'} ; 
                      default {throw "UNRECOGNIZED DNS SUFFIX!:$(($netsettings.fqdn.split('.') | select -last 2 ) -join '.')" ; break ; } ;
                    } ; 
                    }else{
                        throw "NIC.ip $($netsettings.ipaddress) does not PTR resolve to a DNS A with a full fqdn!" ; 
                    } ; 
                } ; 
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
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
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
                if($UserRole){
                    $smsg = "(`$UserRole specified:$($UserRole -join ','))" ; 
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $pltGTCred.UserRole = $UserRole; 
                } else { 
                    $smsg = "(No `$UserRole found, defaulting to:'CSVC','SID' " ; 
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
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
                if((gv Credential) -AND $null -eq $Credential){
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
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatettus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
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
            # default connectivity cmds - force silent 
            $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$silent) ; 
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
            <#if($useExopNoDep){
                # Connect-ExchangeServerTDO use: creds are implied from the PSSession creds; assumed to have EXOP perms
                # 3:14 PM 1/9/2025 no they aren't, it still wants explicit creds to connect - I've just been doing rx10 and pre-initiating
            } else {
            #>
            # useExopNoDep: at this point creds are *not* implied from the PS context creds. So have to explicitly pass in $creds on the new-Pssession etc, 
            # so we always need the EXOP creds block, or at worst an explicit get-credential prompt to gather when can't find in enviro or profile. 
            #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
            if($TenOrg -ne 'CMW'){
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
                # defer cx10/rx10, until just before get-recipients qry
                # connect to ExOP X10
            } ; # skip above on CMW, the mods aren't installed
            if($useEXOP){
                if($useExopNoDep){ 
                    $smsg = "(Using ExOP:Connect-ExchangeServerTDO(), connect to local ComputerSite)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;           
                    TRY{
                        $Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name 
                    }CATCH{$Site=$env:COMPUTERNAME} ;
                    $pltCcX10=[ordered]@{
                        siteName = $Site ;
                        RoleNames = @('HUB','CAS') ;
                        verbose  = $($rPSBoundParameters['Verbose'] -eq $true)
                        Credential = $pltRX10.Credential ; 
                    } ;
                    if($ExopVers){
                        $pltCcX10.add('Version',$ExopVers) ; 
                        write-verbose "(Adding specified -Version:$($ExopVers) to `$pltCcX10)"
                    } ; 
                    # 5:15 PM 4/22/2025 on CMW, have to patch version to Ex2016
                    #if($env:userdomain -eq 'CMW'){
                    if($TenOrg -eq 'CMW'){
                        if($pltCcX10.keys -contains 'Version'){
                            $pltCcX10.version = 'Ex2016' ; 
                        } else { $pltCcX10.add('version','Ex2016') } ;
                    } ; 
                    $smsg = "Connect-ExchangeServerTDO w`n$(($pltCcX10|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #$PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
                    $PSSession = Connect-ExchangeServerTDO @pltCcX10 ; 
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
            if($isNonDomainServer){
                $smsg = "(non-Domain-connected server:Skipping GENERIC ADMS CONN) "  
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            }else {
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
                        if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                        if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #TK 9:44 AM 10/6/2022 need org wide for rolegrps in parent dom (only for onprem RBAC, not EXO)
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
            } ; 
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
        if($UseOP -AND -not $isNonDomainServer -AND -not (get-variable domaincontroller -ea 0)){
            #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((get-variable -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};
            # need to debug the above, credential issue?
            # just get it done
            $domaincontroller = get-GCFast
        }elseif($isNonDomainServer){
            $smsg = "(non-ADDomain-connected, skipping divert to EXO group resolution)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        }  else { 
            # have to defer to get-azuread, or use EXO's native cmds to poll grp members
            # TODO 1/15/2021
            $useEXOforGroups = $true ; 
            $smsg = "$($TenOrg):HAS NO ON-PREM ACTIVEDIRECTORY, DEFERRING ALL GROUP ACCESS & MGMT TO NATIVE EXO CMDS!" ; 
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        if(-not $isNonDomainServer -AND $useForestWide -AND -not $GcFwide){
            #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT: v------
            $smsg = "`$GcFwide = Get-ADDomainController -Discover -Service GlobalCatalog" ;
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;
            $smsg = "Discovered `$GcFwide:$($GcFwide)" ; 
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
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

        $eventIDLookupTbl = Initialize-xopEventIDTable ; 

        # SET DAYS=0 IF USING START/END (they only get used when days is non-0); $platIn.TAG is appended to ticketNO for output vari $vn, and $ofile
        if($Days -AND ($Start -OR $End)){
            write-warning "specified -Days with (-Start -OR -End); If using Start/End, specify -Days 0!" ; 
            Break ; 
        } ; 
        if($Start){[datetime]$Start = $Start } ; 
        if($End){[datetime]$End = $End} ; 
        if($Resultsize -eq 'unlimited' -OR $ResultSize -is [int]){}
        elseif( ($null -eq $ResultSize) -OR ('' -eq $ResultSize) ){
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

        TRY {
            $OrgName = (Get-OrganizationConfig -ea STOP).name ; 
            if($OrgName -eq 'First Organization'){
                if($TenOrg -AND $IsEdgeTransport){
                    $OrgTag = "$($TenOrg)edge" ; 
                } else {
                    $OrgTag = ($orgname  -replace '[^\w]','') ; 
                } ; 
            }elseif(($orgname.ToCharArray() | ?{$_ -cmatch '[A-Z]'}).count -gt 1){
                $OrgTag = ($orgname.ToCharArray() | ?{$_ -cmatch '[A-Z]'}) -join '' ; #Acroynymize the OrgName on caps
            } else { 
                $OrgTag = $orgname  -replace '[^\w]','' ; # use orgname with non-\W chars replaced
            } ; 
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

            if($Sender){
                if($Sender -match $rgxIsPlusAddrSmtpAddr){
                    $smsg = "WARNING! Sender $($Sender) HAS PLUS-ADDRESSING, WON'T WORK FOR EXOP RECIPIENTS!"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ; 
                $pltGMTL.add("Sender",$Sender) ;
                <# 2;18 pm 4/28/2025 drop the testing, the param is [string], let an err occur if a bad sender comes in
                if( ($Sender -split ' *, *') -is [array]){
                    # get-messagetrackinglog  -Sender is a [string] (not [string[]])!
                    $smsg = "-Sender specified - $($Sender) - resolves to an ARRAY, not a supported single-address string!" ; 
                    $smsg += "`nmultipole Senders are _not_ supported for get-messagetrackinglog!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    throw $smsg ; 
                    BREAK ; 
                }else {
                    #$pltGMTL.add('Sender',($Sender -split ' *, *')) ;
                    $pltGMTL.add('Sender',($Sender| out-string)) ; 
                    # this is resulting in an elipses trailing the sender, clearly borked.
                    #$ofile+=",From-$($pltGMTL.Sender.replace("*","ANY"))"  ; 
                } ; 
                #>
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
            $ofile+=if($OrgTag){
                "O_$($OrgTag)"
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

            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Running Get-MessageTrackingLog w`n$(($pltGMTL|out-string).trim())" ; 
            $Srvrs=Get-ExchangeServer -ea STOP ;
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

                    # $hReports.add('EventIDHisto'
                    $smsg = "`n`n## EventID Definitions:" ; 
                    $hReports.EventIDHisto | select -expand Name | foreach-object{                   
                        $smsg += "`n$(($eventIDLookupTbl[$_] | ft -HideTableHeaders |out-string).trim())" ; 
                    } ; 
                    $smsg += "`n`n"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success


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
                    if($mOOO = $mFails | ?{$_.Subject -match '^Automatic\sreply:\s'} ){
                        $smsg = $sBnr3="`n#*~~~~~~v EventID FAIL: Expected Policy Blocked External OutOfOffice (up to last $($DetailLimit))v~~~~~~" ;
                        $smsg += "`n$($mOOO| measure | select -expand count) msgs:Expected Out-of-Office Policy:(attempt to send externally)`n$(($mOOO | select -last $DetailLimit | ft -a $prpGXMTfta |out-string).trim())" ;
                        $smsg += "`n$($sBnr3.replace('~v','~^').replace('v~','^~'))`n"  ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    } ;
                    if($mRecl = $mFails | ?{$_.Subject -match '^Recall:\s'}){
                        $smsg = $sBnr3="`n#*~~~~~~v EventID FAIL: Expected: Recalled message (up to last $($DetailLimit)) v~~~~~~" ;
                        $smsg += "`n$($mRecl| measure | select -expand count) msgs:Expected Sender Recalled Message `n$(($mRecl| select -last $DetailLimit| ft -a $prpGXMTfta |out-string).trim())" ;
                        $smsg += "`n$($sBnr3.replace('~v','~^').replace('v~','^~'))`n"  ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    } ;

                    # reclass everything *not* one of the above as $mfails:
                    if($mFails = $mFails | ?{$_.Subject -notmatch '^Recall:\s' -AND $_.Subject -notmatch '^Automatic\sreply:\s'}){
                        <#$smsg = $sBnr3="`n#*~~~~~~v EventID FAIL: Other Failure message (up to last $($DetailLimit)) v~~~~~~" ;
                        $smsg += "`n$(($mFails| select -last $DetailLimit | ft -a |out-string).trim())" ;
                        $smsg += "`n$($sBnr3.replace('~v','~^').replace('v~','^~'))`n"  ;
                        write-host -foregroundcolor yellow $smsg ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        #>
                        if($mFailAddr = $mFails | ?{$_.EventID -eq 'FAIL' -AND $_.Source -eq 'ROUTING'}){
                            $smsg = $sBnr3="`n#*~~~~~~v EventID FAIL: BAD ADDRESS FAILS: (EventID:'FAIL' & Source:'ROUTING') ($(($mFailAddr|measure).count)msgs, echo up to last $($DetailLimit))) v~~~~~~" ;
                            $smsg += "`n$(($mFailAddr| select -last $DetailLimit | ft -a |out-string).trim())" ;
                            $smsg += "`n$($sBnr3.replace('~v','~^').replace('v~','^~'))`n"  ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        } ;
                        if($mFailRej = $mFails | ?{$_.EventID -eq 'FAIL' -AND $_.Source -eq 'SMTP'}){
                            $smsg = $sBnr3="`n#*~~~~~~v EventID FAIL: REJECTED BY RECIPIENT SERVER : (EventID:'FAIL' & Source:'SMTP') ($(($mFailRej|measure).count)msgs, echo up to last $($DetailLimit)) v~~~~~~" ;
                            $smsg += "`n$(($mFailRej| select -last $DetailLimit | ft -a |out-string).trim())" ;
                            $smsg += "`n$($sBnr3.replace('~v','~^').replace('v~','^~'))`n"  ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        } ;
                        if($mFailDom = $mFails | ?{$_.EventID -eq 'FAIL' -AND $_.Source -eq 'DNS'}){
                            $smsg = $sBnr3="`n#*~~~~~~v EventID FAIL: BAD DOMAIN : (EventID:'FAIL' & Source:'DNS') ($(($mFailDom|measure).count)msgs, echo up to last $($DetailLimit)) v~~~~~~" ;
                            $smsg += "`n$(($mFailDom| select -last $DetailLimit | ft -a |out-string).trim())" ;
                            $smsg += "`n$($sBnr3.replace('~v','~^').replace('v~','^~'))`n"  ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        } ;
                        # anything *not* filtered above
                        if($mFailsOther = $mFails | 
                            ?{$_.Subject -notmatch '^Recall:\s' -AND $_.Subject -notmatch '^Automatic\sreply:\s' -AND -not(
                                $_.EventID -eq 'FAIL' -AND $_.Source -eq 'ROUTING') -AND -NOT($_.EventID -eq 'FAIL' -AND $_.Source -eq 'SMTP') -AND -NOT(
                                $_.EventID -eq 'FAIL' -AND $_.Source -eq 'DNS') }){
                            $smsg = $sBnr3="`n#*~~~~~~v EventID FAIL: OTHER:) ($(($mFailsOther|measure).count)msgs, echo up to last $($DetailLimit)) v~~~~~~" ;
                            $smsg += "`n$(($mFailsOther| select -last $DetailLimit | ft -a |out-string).trim())" ;
                            $smsg += "`n$($sBnr3.replace('~v','~^').replace('v~','^~'))`n"  ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        } ;

                    } else {
                        $smsg = "(no non-RECALL,non-OOO block, FAILs logged)" 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
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
                    # defers are infra issue, not recipient/resolution, no sense dumping details, other than to let user know which delayed

                    $smsg = $sBnrS="`n#*------v DEFER's RETRY handling (up to last $($DetailLimit) messages) v------" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H2 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    # hash index the main msg list: (post filtering lags)
                    $swSEND = [Diagnostics.Stopwatch]::StartNew();
                    $mSends = $msgs|?{$_.EventID -eq 'SEND'}
                    $ndxMsgsSEND = @{} ;
                    $mSends | foreach-object{
                        $ndxMsgsSEND[$_.MessageId] = $_ ; 
                    } ; 
                    remove-variable mSends -force 
                    if($host.version.major -ge 3){
                        $smsg =  ("$($smsg) {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $swSEND.Elapsed) ; #outputs: Elapsed Time: 0d 0h 0m 10s 101ms
                    } else {
                        $smsg = "$($smsg) $($swSEND.Elapsed.ToString()))" ; # outputs: Elapsed Time: (HH:MM:SS.ms) 00:00:15.8418655
                    } ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    #

                    $midn = 20 ; $namn = 15 ; 
                    $dmprcd = 0 ; 
                    foreach($dmsg in ($mDefers | select -last $DetailLimit)){
                        $dmprcd++ ; 
                        #$smsg ="`n#*~~~~~~v " ; 
                        $smsg = "#$($dmprcd) " ; 
                        $smsg += "MsgID:$(($dmsg.MessageId -replace '[\<\>]','').substring(0,[System.Math]::Min(($dmsg.MessageId -replace '[\<\>]','').length,((($dmsg.MessageId -replace '[\<\>]','').ToCharArray() | select -first $midn) -join '').length)))..." ; 
                        $smsg += " Fr:$($dmsg.Sender.substring(0,[System.Math]::Min($dmsg.Sender.length,(($dmsg.Sender.ToCharArray() | select -first $namn) -join '').length)))..." ; 
                        #$smsg += " To:$($dmsg.Recipients.substring(0,[System.Math]::Min($dmsg.Recipients.length,(($dmsg.Recipients.ToCharArray() | select -first $namn) -join '').length)))..." ; 
                        #$smsg += " To:$( ($dmsg.Recipients -join ',').substring(0,[System.Math]::Min(($dmsg.Recipients -join ',').length,((($dmsg.Recipients -join ',').ToCharArray() | select -first $namn) -join '').length)))..." ; 
                        $sRcp = $dmsg.Recipients -join ',' ; 
                        $smsg += " To:$( $sRcp.substring(0,[System.Math]::Min($sRcp.length,(($sRcp.ToCharArray() | select -first $namn) -join '').length)))..." ; 
                        $smsg += " Sub:$($dmsg.MessageSubject.substring(0,[System.Math]::Min($dmsg.MessageSubject.length,(($dmsg.MessageSubject.ToCharArray() | select -first $namn) -join '').length)))..." ;         
                        #$smsg += " v~~~~~~" ; 
                        #$sBnr3 = $smsg ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;   
        
                        #$handling = $msgs |?{$_.messageid -eq $dmsg.MessageId -AND $_.eventid -match 'DEFER|SEND'} | sort TimestampLocal ;
                        # use $ndxMsgsSEND, postfilter too slow, aah, but msgid is repeated across evts pftftf
                        if($confirmedSEND = $ndxMsgsSEND[$dmsg.MessageId]){
                            $aHndl = @() ; 
                            $aHndl += @('DEFER') ; 
                            #$aHndl += @($handling[0].Timestamp) ; 
                            $aHndl += @($dmsg.Timestamp) ; 
                            $aHndl += @('SEND') ; 
                            #$aHndl += @($handling[1].Timestamp) ; 
                            $aHndl += @($confirmedSEND.Timestamp) ; 
                    

                            $aHndl += @('Retried Sent') ; 
                            #$aHndl += @("{0:dd}d:{0:hh}h:{0:mm}m:{0:ss}s:{0:fff}ms" -f (new-timespan -start $handling[0].Timestamp -end $handling[-1].Timestamp)) ; 
                            $aHndl += @("{0:dd}d:{0:hh}h:{0:mm}m:{0:ss}s:{0:fff}ms" -f (new-timespan -start $dmsg.Timestamp -end $confirmedSEND.Timestamp)) ; 
                            $smsg = "=>$($ahndl -join ' : ')`n`n" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level SUCCESS } else{ write-host -foregroundcolor green $smsg } ;                         

                        } else {
                            $smsg = "UNABLE TO LOCATE A SUCCESFUL SEND eventID for:" ; 
                            $smsg += "`n$(($dmsg | fl |out-string).trim())" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        }


                        #$smsg = "$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;
                        #if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H3 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;   
                    } ; 
    
                    $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H2 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                } else {
                    $smsg = "(no DEFERs logged)" 
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
} ;
#*------^ Get-MessageTrackingLogTDO ^------