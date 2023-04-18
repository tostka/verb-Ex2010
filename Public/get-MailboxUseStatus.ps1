#*------v get-MailboxUseStatus.ps1 v------
function get-MailboxUseStatus {
<#
    .SYNOPSIS
    get-MailboxUseStatus - Analyze and summarize a specified array of Exchange OnPrem mailbox objects to determine 'in-use' status, and export summary statistics to CSV file
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2022-01-28
    FileName    : get-MailboxUseStatus.ps1
    License     : MIT License
    Copyright   : (c) 2023 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell,Exchange,ExchangeOnline,Mailbox,Verification,DataGathering
    AddedCredit : REFERENCE
    AddedWebsite:	URL
    AddedTwitter:	URL
    REVISIONS
    * 2:02 PM 4/17/2023 rev: $MinNoWinRMVersion from 2.0.6 => 3.0.0.
    * 1:36 PM 4/13/2023 fix: add err suppr: missing -ea 0's on gcm tests, backward compat -silent param test before use in $pltrxo or $pltrx10. Ren use of xow alias with full invoke-xowrapper calls
    * 9:57 AM 4/12/2023 add: emit the outputfiles array to the pipeline, to capture and reuse it for post-procesesing.
    * 4:59 PM 4/11/2023 finally got through a full pass to export on JB; 
    added code to check for dn break (moved/renamed adu), and failback to $oprcp or discovered $adu.dn, upn etc. ; added tests logic for rmbx/migrated user combos; 
    and retrieval of rmbx when found. patch over get-AzureAdUser to user the adu.upn (unless it's a broken user); 
    rem'd out Catch Continues in loops, along with Throw's (both were exiting the loop, advancing to next mbx). Point is to capture even semi-broken user info.
    defer's gmbxstat if no local Usermailbox in the get-recipient.
    * 10:29 AM 4/10/2023 found some get-aduser failing because someone's moved the user, which breaks get-aduser -id [DN]: add: targeted catch & retry on get-aduser $mbx.alias.
    * 3:45 PM 4/7/2023 updated $prpExportCSV, wasn't outputting TagYellow, etc and field order was unupdated; plus side- ran wo issues on JB using EOM205p6 (confirmed works both EOM310 and the older pre WinRM vers). 
    Add: try/catch around get-mailusestatus data calls to exo & exop (replaced single t/c on whole loop); wrote fail-through data code for data gather fails
    * 2:35 PM 4/5/2023 add: params: TagYellow, TagGreen, TagPurple; fix: had broken -LicensedMail isExoLicensed lookups, by loop-top blanking $LicensedMail (which is the driving param for the func)
    * 3:47 PM 4/4/2023 added equate $EOMMinNoWinRMVersion = $MinNoWinRMVersion ; ported over eom module ipmo/test, and version check ; 
    * 10:00 AM 3/30/2023 updated prpExportCSV to include all hsummary fields, including split brains; added SB#: banner view of ttl split-brains so far in pass.
    * 4:26 PM 3/29/2023 some splitbrains are false; nulling all working varis loop top, to exclude holdover ; 
     REN'D $modname => $EOMModName ; $EOMmodname = 'ExchangeOnlineManagement' ;
    Added code and params to pre-detect native EOMv3 connections and skip the invoke-XOWWrapper use altogether.
    REN: $MinNoWinRMVersion => $EOMMinNoWinRMVersion
    * 3:54 PM 3/28/2023 workstation installing latest: ExchangeOnlineManagement 3.1.0: (3.2.0-Preview2 is out, 9d old, 757 dl's, p1 has 1857 30d ago; 3.1.0 has 851k dl's.
        - [About the Exchange Online PowerShell V2 module and V3 module | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#updates-for-version-300-the-exo-v3-module)
        - [PowerShell Gallery | ExchangeOnlineManagement 3.1.0](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.1.0)
        PS> Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.1.0 -scope CurrentUser 
        # update:
        PS> Get-InstalledModule ExchangeOnlineManagement | Format-List Name,Version,InstalledLocation ; 
        PS> Update-Module -Name ExchangeOnlineManagement -Scope CurrentUser ; Import-Module ExchangeOnlineManagement; Get-Module ExchangeOnlineManagement ; 
        This is a General Availability (GA) release of the Exchange Online Powershell V3 module. Exchange Online cmdlets in this module are REST-backed and do not require Basic Authentication to be enabled in WinRM.
        Please check the documentation here - https://aka.ms/exov3-module.
        Bug reporting, run log:
        Connect-ExchangeOnline -EnableErrorReporting -LogDirectoryPath <Path to store log file> -LogLevel All
        >  Note
        > Frequent use of the Connect-ExchangeOnline and Disconnect-ExchangeOnline 
        > cmdlets in a single PowerShell session or script might lead to a memory leak. 
        > The best way to avoid this issue is to use the CommandName parameter on the 
        > Connect-ExchangeOnline cmdlet to limit the cmdlets that are used in the session.
        -> Mem leaks in GA! whatta load of BS!
        All versions of the module are supported in Windows PowerShell 5.1.
            PowerShell 7 on Windows requires version 2.0.4 or later.
            Version 2.0.5 or later of the module requires the Microsoft .NET Framework 
            4.7.1 or later to connect. Otherwise, you'll get an 
            System.Runtime.InteropServices.OSPlatform error. This requirement shouldn't be 
            an issue in current versions of Windows. For more information about versions of 
            Windows that support the .NET Framework 4.7.1, see this article. 
        ... added split-brain support (which entailed bringing lagging rxo2/exov2 code into function, and adding xo support to this); add: -useExov2 ; spliced in debugged/latest import-xoW() ;
        new exported properties in hSummary:MbxExchangeGUID, XoMbxExchangeGUID, XoMbxTotalItemSizeGB, XoMbxLastLogonTime,xoMbxWhenChanged, xoMbxWhenCreated, SplitBrain; 
        code to get-xomailbox & get-xomailboxstats & identify split-brain state ; 
    * 4:20 PM 3/22/2023 typo: $hSummary 'ADLastLogon' -> ADLastLogonTime ;  was drawing blank adu.lastlogon, so added failback to lastlogontimestamp use;  updated adu.lastlogon, compare resolved for blank value; added -ADLastLogon to the script, which enables the get-aduser.lastlogon 
        collection - only stored on the local DC, so it's going to be spotty, but where 
        there's no mbxstat.lastlogon, it at least might provide evidence user is logging 
        into AD, and is active at that level (to avoid deleting unused legit mailboxes)
    * 1:17 PM 3/21/2023 reworked field order in $prpExportCSV (useful xlsx order)
    * 3:07 PM 11/28/2022 working. CBH example #3 still has issues with example that post-exports csv - not 
    collapsing objects; but native export csv & xml works fine. -outputobject works 
    fine as well, as long as you massage the exports and ensure they properly 
    expand the objects going into csv ;  ;fixed export props, csvs were still 
    objects;  ADD: expanded/joined objects for csv exports, both internally, and in 
    the CBH example -outputobject csv export ; added $OutputFiles array and END reporting (pita to dig them out 
    retroactively); CBH example, added export for last that uses -ooutputobject, 
    and skips csv/xml in the script 
    * 4:18 PM 11/23/2022 defaulted mbxstats value to an error (if user never logged in, there are no statrtys to return) ;  spliced over updated SvcConn block ;  add: xoWP{} and support, refactored the svc conn patterns to try to approx xow order wo running xow itself.
    # 2:49 PM 3/8/2022 pull Requires -modules ...verb-ex2010 ref - it's generating nested errors, when ex2010 requires exo requires ex2010 == loop.
    * 4:12 PM 3/7/2022 moved the isExoLicensed test below the lic loop ; fixed a dangling w-v ; 
    * 4:09 PM 3/3/2022 coded in test for adu.memberof matching the $xxxmeta.rgx , to validate that a licensure-grp is in play, to explain the 44% of existing prev profiled users that haven't got an EXO-supporting lic
    * 3:50 PM 3/1/2022 add ADMemberof and parse for lic grp; fixed non-default quota typos ; updated CBH properties returned list; updated the function return tests from switch to simpler [type] tests.
    * 4:28 PM 2/28/2022 debugged, to full pass, added conversion to gb decimal for sizes, and formatted dates on timestmps, added test for EXO-usermailbox-supporting license;
        implemented external verb-AAD:get-ExoMailboxLicenses() & verb-EX2010:get-MailboxDatabaseQuotas() & verb-exo:get-ExoMailboxLicenses() to provide the content  ;
        validated pipeline -mailboxes functioning find. Probably should implement an xml export, along with csv.
    * 4:57 PM 2/25/2022 WIP: added services mgmt, & cred handling, pulling AADU licenses etc (need to parse end eval if 'Mailbox'-supporting lic assigned), added local EXOP: quotas, server, db, totoalitemsize for mbx, etc.
    * 1:48 PM 1/28/2022 hit a series of mbxs that were onprem in AM, but migrated in PM; also  they've got 2 david.smith@toro.com's onboarded, both with same UPN, shift gmbxstat to DN it's more specific ; expanded added broad range of ADUser.geopoliticals; added calculated SiteOU as well; working
    .DESCRIPTION
    get-MailboxUseStatus - Analyze and summarize a specified array of Exchange OnPrem mailbox objects to determine 'in-use' status, and export summary statistics to CSV file

    Collects & exports to CSV (or outputs to pipeline, where -outputobject specified), the following information per Mailbox/ADUser
        AADUSMTPProxyAddresses
        AADUAssignedLicenses
        AADUDirSyncEnabled
        AADULastDirSyncTime
        AADUserPrincipalName
        ADCity
        ADCompany
        ADCountry
        ADcountryCode
        ADcreateTimeStamp
        ADDepartment
        ADDivision
        ADEmployeenumber
        ADemployeeType
        ADEnabled
        ADGivenName
        ADmailNickname
        ADMemberof
        ADMobilePhone
        ADmodifyTimeStamp
        ADOffice
        ADOfficePhone
        ADOrganization
        ADphysicalDeliveryOfficeName
        ADPOBox
        ADPostalCode
        ADSMTPProxyAddresses
        ADState
        ADStreetAddress
        ADSurname
        ADTitle
        DistinguishedName
        IsExoLicensed
        LicGrouppDN
        MbxDatabase
        MbxIssueWarningQuotaGB
        MbxLastLogonTime
        MbxProhibitSendQuotaGB
        MbxProhibitSendReceiveQuotaGB
        MbxRetentionPolicy
        MbxServer
        MbxTotalItemSizeGB
        MbxUseDatabaseQuotaDefaults
        name
        ParentOU
        samaccountname
        SiteOU
        UserPrincipalName
        WhenChanged
        WhenCreated
        WhenMailboxCreated

        For CSV exports, the 'system-object' AADUSMTPProxyAddresses, ADSMTPProxyAddresses, & ADMemberof are condensed into a semicolon(;)-concatonated string. 
        The XML contains the full nested object tree (to the extend the depth is specified).

    .PARAMETER Mailboxes
    Array of Exchange OnPrem Mailbox Objects[-Mailboxes `$mailboxes]
    .PARAMETER Ticket
    Ticket number[-Ticket 123456]
    .PARAMETER SiteOUNestingLevel
    Number of levels down the SiteOU name appears in the DistinguishedName (Used to calculate SiteOU: counting from right; defaults to 5)[-SiteOUNestingLevel 3]
    .PARAMETER ADLastLogon
    Switch to query for and include - broadly inaccruate (stored single dc logged to) - ADUser.LastLogon spec
    .PARAMETER TagYellow
    Array of UPNs to be tagged with 'TagYellow:Y' in output fields[-TagYellow 'Aleksandra.Kotas@toro.com','Daniel.Hughes@toro.com']
    .PARAMETER TagGreen
    Array of UPNs to be tagged with 'TagGreen:Y' in output fields[-TagGreen 'Anita.Stahl@toro.com','Boguslaw.Drozd@toro.com']
    .PARAMETER TagPurple
    Array of UPNs to be tagged with 'TagPurple:Y' in output fields[-TagPurple 'Hayter.VideoConference@toro.com','Juergen.Hoffmann@toro.com']
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER outputObject
    Object output switch [-outputObject]
    .EXAMPLE
    PS> get-MailboxUseStatus -ticket 665437 -mailboxes $NonTermUmbxs -verbose  ;
    Example processing the specified array, and writing report to CSV, with -verbose output
    .EXAMPLE
    PS> (get-mailbox -id USER) | get-mailboxusestatus -ticket 999999 -verbose ;
    Pipeline example
    .EXAMPLE
    PS>  $ticket = '123456' ; 
    PS>  write-host "Profile specified list of users (pre-filtered for recipienttypedetails & not stored in term-related OUs), and below Users OUs)Then postfilter and count the number actually ADEnabled" ; 
    PS>  write-host "get-mailbox -ResultSize unlimited..." ; 
    PS>  $sw = [Diagnostics.Stopwatch]::StartNew();
    PS>  $allExopmbxs = get-mailbox -ResultSize unlimited ;
    PS>  $sw.Stop() ; write-host ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
    PS>  $sw = [Diagnostics.Stopwatch]::StartNew();
    PS>  $of = ".\logs\allExopmbxs-$(get-date -format 'yyyyMMdd-HHmmtt').xml" ; 
    PS>  touch $of ; $of = (resolve-path $of).path ; 
    PS>  write-host "export/import xml:$($of)" ; 
    PS>  write-host "`$allExopmbxs | export-clixml:$($of)" ;
    PS>  $allExopmbxs | export-clixml -path $of -depth 100 -f;
    PS>  $allExopmbxs = import-clixml -path $of ;
    PS>  write-host "`$allExopmbxs.count:$(($allExopmbxs|  measure | select -expand count |out-string).trim())" ; 
    PS>  $sw.Stop() ;write-host ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
    PS>  write-host "post filter for NonTermUMbxs..." ; 
    PS>  $sw = [Diagnostics.Stopwatch]::StartNew();
    PS>  $NonTermUmbxs = $allExopmbxs | ?{$_.recipienttypedetails -eq 'UserMailbox' -AND $_.distinguishedname -notmatch ',OU=(Disabled|TERM),' -AND $_.distinguishedname -match ',OU=Users,'} ;
    PS>  $sw.Stop() ; write-host ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
    PS>  $of = $of.replace('allExopmbxs-','NonTermUmbxs-') ; 
    PS>  touch $of ; $of = (resolve-path $of).path ; 
    PS>  write-host "`$NonTermUmbxs | export-clixml:$(resolve-path $of.replace('allExopmbxs-','NonTermUmbxs-'))" ; 
    PS>  $NonTermUmbxs | export-clixml -path $of -depth 100 -f ;
    PS>  $NonTermUmbxs = import-clixml -path $of ; 
    PS>  write-host "`$NonTermUmbxs.count:$(($NonTermUmbxs|  measure | select -expand count |out-string).trim())" ; 
    PS>  write-host "Run `$NonTermUmbxs through get-MailboxUseStatus" ; 
    PS>  $sw = [Diagnostics.Stopwatch]::StartNew();
    PS>  $Results = get-MailboxUseStatus -ticket $ticket -mailboxes $NonTermUmbxs -outputObject ;
    PS>  $sw.Stop() ;write-host ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
    PS>  $sw = [Diagnostics.Stopwatch]::StartNew();
    PS>  write-host "Measure ADEnabled users:" ; 
    PS>  $results |?{$_.ADEnabled} |  measure | select -expand count  ;
    PS>  if($results){
    PS>      $prpExportCSV = @{name="AADUAssignedLicenses";expression={($_.AADUAssignedLicenses) -join ";"}},'AADUDirSyncEnabled','AADULastDirSyncTime','AADUserPrincipalName',@{name="AADUSMTPProxyAddresses";expression={$_.AADUSMTPProxyAddresses.SmtpProxyAddresses -join ";"}},'ADCity','ADCompany','ADCountry','ADcountryCode','ADcreateTimeStamp','ADDepartment','ADDivision','ADEmployeenumber','ADemployeeType','ADEnabled','ADGivenName','ADmailNickname',@{name="ADMemberof";expression={$_.ADMemberof -join ";"}},'ADMobilePhone','ADmodifyTimeStamp','ADOffice','ADOfficePhone','ADOrganization','ADphysicalDeliveryOfficeName','ADPOBox','ADPostalCode',@{name="ADSMTPProxyAddresses";expression={$_.ADSMTPProxyAddresses.SmtpProxyAddresses -join ";"}},'ADState','ADStreetAddress','ADSurname','ADTitle','DistinguishedName','IsExoLicensed','LicGrouppDN','MbxDatabase','MbxIssueWarningQuotaGB','MbxLastLogonTime','MbxProhibitSendQuotaGB','MbxProhibitSendReceiveQuotaGB','MbxRetentionPolicy','MbxServer','MbxTotalItemSizeGB','MbxUseDatabaseQuotaDefaults','Name','ParentOU','samaccountname','SiteOU','UserPrincipalName','WhenChanged','WhenCreated','WhenMailboxCreated' ; 
    PS>      $ofile = ".\logs\$($ticket)-get-MailboxUseStatus-Summary-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
    PS>      write-host "`$ofile:$($ofile)" ;
    PS>      $results | select $propsExportCSV | export-csv -NoTypeInformation -path $ofile ;
    PS>      $ofile = $ofile.replace('.csv','.xml') ; 
    PS>      write-host "`$ofile:$($ofile)" ;
    PS>      $results | export-clixml -depth 100 -path $ofile ;
    PS>  } ;
    PS>  $sw.Stop() ;write-host ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
    Profile specified list of users (pre-filtered for recipienttypedetails & not stored in term-related OUs, and below Users OUs)
    Capture returned output object, and postfilter and count the number actually ADEnabled. Then manually export the results to csv & xml.
    Issue above: function works fine, but the post object export above still coming through as objects in csv (unusable). Native non-'-outputobject' csv exports just fine...
    .EXAMPLE
    PS>  $ticket = '123456' ;
    PS>  write-host "Profile specified list of users (pre-filtered for recipienttypedetails & not stored in term-related OUs), and below Users OUs)Then postfilter and count the number actually ADEnabled" ;
    PS>  write-host "get-mailbox -ResultSize unlimited..." ;
    PS>  $sw = [Diagnostics.Stopwatch]::StartNew();
    PS>  $allExopmbxs = get-mailbox -ResultSize unlimited ;
    PS>  $sw.Stop() ;
    PS>   write-host ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ;
    PS>  $sw = [Diagnostics.Stopwatch]::StartNew();
    PS>  $of = ".\logs\allExopmbxs-$(get-date -format 'yyyyMMdd-HHmmtt').xml" ;
    PS>  touch $of ;
    PS>   $of = (resolve-path $of).path ;
    PS>  write-host "export/import xml:$($of)" ;
    PS>  write-host "$allExopmbxs | export-clixml:$($of)" ;
    PS>  $allExopmbxs | export-clixml -path $of -depth 100 -f;
    PS>  $allExopmbxs = import-clixml -path $of ;
    PS>  write-host "$allExopmbxs.count:$(($allExopmbxs|  measure | select -expand count |out-string).trim())" ;
    PS>  $sw.Stop() ;
    PS>  write-host ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ;
    PS>  write-host "post filter for NonTermUMbxs..." ;
    PS>  $sw = [Diagnostics.Stopwatch]::StartNew();
    PS>  $NonTermUmbxs = $allExopmbxs | ?{$_.recipienttypedetails -eq 'UserMailbox' -AND $_.distinguishedname -notmatch ',OU=(Disabled|TERM),' -AND $_.distinguishedname -match ',OU=Users,'} ;
    PS>  $sw.Stop() ;
    PS>   write-host ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ;
    PS>  $of = $of.replace('allExopmbxs-','NonTermUmbxs-') ;
    PS>  touch $of ;
    PS>   $of = (resolve-path $of).path ;
    PS>  write-host "$NonTermUmbxs | export-clixml:$(resolve-path $of.replace('allExopmbxs-','NonTermUmbxs-'))" ;
    PS>  $NonTermUmbxs | export-clixml -path $of -depth 100 -f ;
    PS>  $NonTermUmbxs = import-clixml -path $of ;
    PS>  write-host "$NonTermUmbxs.count:$(($NonTermUmbxs|  measure | select -expand count |out-string).trim())" ;
    PS>  write-host "Run $NonTermUmbxs through get-MailboxUseStatus" ;
    PS>  $sw = [Diagnostics.Stopwatch]::StartNew();
    PS>  get-MailboxUseStatus -ticket $ticket -mailboxes $NonTermUmbxs  ;
    PS>  $sw.Stop() ;
    PS>  write-host ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ;
    PS>  $sw = [Diagnostics.Stopwatch]::StartNew();
    PS>  write-host "Measure ADEnabled users:" ;
    PS>  $Results = import-clixml -path '.\path-to\123456-get-MailboxUseStatus-EXEC-yyyymmdd-hhmmPM.XML ;
    PS>  $results |?{$_.ADEnabled} |  measure | select -expand count  ;
    PS>  $sw.Stop() ;
    PS>  write-host ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ;
    Above - working - Profile specified list of users (pre-filtered for recipienttypedetails & not stored in term-related OUs, and below Users OUs)
    does *not* use -outputobject, which has the script output native .csv & .xml files (full functioning), then count the number actually ADEnabled.
    .EXAMPLE
    PS>  $results = 'USER1@DOMAIN.com','USER2@DOMAIN.com' | get-mailbox | get-MailboxUseStatus -ticket 123456 -outputObject ;
    Feed a list of UPNs through get-mailbox and then through the script, via pipeline
    .EXAMPLE
    PS>  $oxml = 'c:\usr\work\incid\allExopmbxs-20230320-1600PM.xml' ;
    PS>  $TagYellow = 'fname.lname@toro.com','fname.lname2@toro.com' ; 
    PS>  $TagGreen = 'fname.lname3@toro.com','fname.lname4@toro.com' ; 
    PS>  $TagPurple = 'fname.lname5@toro.com','fname.lname6@toro.com' ; 
    PS>  $allExopmbxs = import-clixml $oxml  ;
    PS>  $NonTermUmbxs = $allExopmbxs | ?{$_.recipienttypedetails -eq 'UserMailbox' -AND $_.distinguishedname -notmatch ',OU=(Disabled|TERM),' -AND $_.distinguishedname -match',OU=Users,'} ;
    PS>  $NonTermUmbxs |  measure | select -expand count ;
    PS>  $Results = get-MailboxUseStatus -ticket 755280 -mailboxes $NonTermUmbxs -adlastlogon -TagYellow $TagYellow -TagGreen $TagGreen -TagPurple $TagPurple -verbose ; 

        [trimmed]
        15:44:53: INFO:  Exporting summary for 205 mailboxes to CSV:
        c:\scripts\logs\755280-get-MailboxUseStatus-EXEC-20230404-1455PM.csv
        15:44:57: INFO:  Exporting summary for 205 mailboxes to XML:
        c:\scripts\logs\755280-get-MailboxUseStatus-EXEC-20230404-1455PM.XML
        15:44:58: INFO:  $OutputFiles:
        c:\scripts\logs\755280-get-MailboxUseStatus-Transcript-BATCH-EXEC-20230404-1455PM-trans-log.txt
        c:\scripts\logs\755280-get-MailboxUseStatus-LOG-BATCH-EXEC-20230404-1455PM-log.txt
        c:\scripts\logs\755280-get-MailboxUseStatus-EXEC-20230404-1455PM.csv
        c:\scripts\logs\755280-get-MailboxUseStatus-EXEC-20230404-1455PM.XML
        15:44:58: INFO:  Transcript stopped, output file is C:\scripts\logs\755280-get-MailboxUseStatus-Transcript-BATCH-EXEC-20230404-1455PM-trans-log.txt

    Demo use of the -Tagyellow, -TagGreen & -TagPurple params (with arrays of UPNs). 
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    #Requires -Version 3
    #requires -PSEdition Desktop
    #Requires -Modules ActiveDirectory, verb-ADMS, verb-IO, verb-logging, verb-Network, verb-Text, verb-AAD
    #Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("US","GB","AU")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)]#positiveInt:[ValidateRange(0,[int]::MaxValue)]#negativeInt:[ValidateRange([int]::MinValue,0)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    ###[Alias('Alias','Alias2')]
    PARAM(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        [string]$TenOrg = 'TOR',
        [Parameter(Position=1,Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of Exchange OnPrem Mailbox Objects[-Mailboxes `$mailboxes]")]
        $Mailboxes,
        [Parameter(Mandatory=$true,HelpMessage="Ticket number[-Ticket 123456]")]
        [string]$Ticket,
        [Parameter(HelpMessage="Switch to confirm Mail-related license assigned on mailbox(es)[-LicensedMail]")]
        [switch] $LicensedMail = $true,
        [Parameter(HelpMessage="Switch to query for and include - broadly inaccruate (stored single dc logged to) - ADUser.LastLogon spec[-ADLastLogon]")]
        [switch] $ADLastLogon,
        [Parameter(HelpMessage="Array of UPNs to be tagged with 'TagYellow:Y' in output fields[-TagYellow 'Aleksandra.Kotas@toro.com','Daniel.Hughes@toro.com']")]
        [string[]]$TagYellow,
        [Parameter(HelpMessage="Array of UPNs to be tagged with 'TagGreen:Y' in output fields[-TagGreen 'Anita.Stahl@toro.com','Boguslaw.Drozd@toro.com']")]
        [string[]]$TagGreen,
        [Parameter(HelpMessage="Array of UPNs to be tagged with 'TagPurple:Y' in output fields[-TagPurple 'Hayter.VideoConference@toro.com','Juergen.Hoffmann@toro.com']")]
        [string[]]$TagPurple,
        [Parameter(HelpMessage="Number of levels down the SiteOU name appears in the DistinguishedName (Used to calculate SiteOU: counting from right; defaults to 5)[-SiteOUNestingLevel 3]")]
        [int]$SiteOUNestingLevel=5,
        [Parameter(HelpMessage="Object output switch [-outputObject]")]
        [switch] $outputObject,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2=$true
    ) # PARAM BLOCK END

    BEGIN {
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;

        # 3:47 PM 3/1/2022 added memberof, need to track down that unlic'd aren't even members of lic grps
        # 3:24 PM 3/22/2023 coming through lastlogon blank, but lastLogonTimestamp present, pull both and fail back
        $prpADU = 'employeenumber','createTimeStamp','modifyTimeStamp','City','Company','Country','countryCode','Department',
            'Division','EmployeeNumber','employeeType','GivenName','Office','OfficePhone','Organization','MobilePhone',
            'physicalDeliveryOfficeName','POBox','PostalCode','State','StreetAddress','Surname','Title','proxyAddresses','memberof',
            'LastLogon','lastLogonTimestamp'  | select -unique ;
        # ,'lastLogonTimestamp' ; worthless, only updated every 9-14d, and then only on local dc - is converting to 1600 as year
        # 12:55 PM 3/22/2023 need to try somethign on ad.lastlogon
        # this isn't used, disabled, we're working with raw values assigned into summary
        <#$selectADU = 'DistinguishedName','Enabled','GivenName','Name','ObjectClass','ObjectGUID','SamAccountName','SID',
            'Surname','UserPrincipalName','employeenumber','createTimeStamp','modifyTimeStamp', 
            @{n='LastLogon';e={[DateTime]::FromFileTime($_.LastLogon)}} ; 
        #>
        $prpAadu = 'UserPrincipalName','GivenName','Surname','DisplayName','AccountEnabled','Description','PhysicalDeliveryOfficeName',
            'JobTitle','AssignedLicenses','Department','City','State','Mail','MailNickName','LastDirSyncTime','OtherMails','ProxyAddresses' ;
        # keep the smtp prefix to tell prim/alias addreses
        #$propsAxDUserSmtpProxyAddr = @{Name="SmtpProxyAddresses";Expression={ ($_.ProxyAddresses.tolower() |?{$_ -match 'smtp:'})  -replace ('smtp:','') } } ;
        $prpAxDUserSmtpProxyAddr = @{Name="SmtpProxyAddresses";Expression={ ($_.ProxyAddresses.tolower() |?{$_ -match 'smtp:'}) } } ;
        <#$prpExportCSV = @{name="AADUAssignedLicenses";expression={($_.AADUAssignedLicenses) -join ";"}},'AADUDirSyncEnabled','AADULastDirSyncTime',
            'AADUserPrincipalName',@{name="AADUSMTPProxyAddresses";expression={$_.AADUSMTPProxyAddresses.SmtpProxyAddresses -join ";"}},
            'ADCity','ADCompany','ADCountry','ADcountryCode','ADcreateTimeStamp','ADDepartment','ADDivision','ADEmployeenumber','ADemployeeType',
            'ADEnabled','ADGivenName','ADmailNickname',@{name="ADMemberof";expression={$_.ADMemberof -join ";"}},'ADMobilePhone','ADmodifyTimeStamp',
            'ADOffice','ADOfficePhone','ADOrganization','ADphysicalDeliveryOfficeName','ADPOBox','ADPostalCode',@{name="ADSMTPProxyAddresses";expression={$_.ADSMTPProxyAddresses.SmtpProxyAddresses -join ";"}},
            'ADState','ADStreetAddress','ADSurname','ADTitle','DistinguishedName','IsExoLicensed','LicGrouppDN','MbxDatabase',
            'MbxIssueWarningQuotaGB','MbxLastLogonTime','MbxProhibitSendQuotaGB','MbxProhibitSendReceiveQuotaGB','MbxRetentionPolicy',
            'MbxServer','MbxTotalItemSizeGB','MbxUseDatabaseQuotaDefaults','Name','ParentOU','samaccountname','SiteOU','UserPrincipalName',
            'WhenChanged','WhenCreated','WhenMailboxCreated' ; 
        #>
        # 1:10 PM 3/21/2023 rework field order to put usefuls on left/first:
        # 10:05 AM 3/22/2023 add back ADLastLogon (driven by -ADLastLogon switch)
        # 4:56 PM 3/29/2023 need split brain etc: output the hsummary and rearranged to pref'd csv field order.
        # 2:41 PM 4/7/2023  add the fields to export '$TagYellow,','$TagGreen,','$TagPurple,'; rearrange order for key up front
        $prpExportCSV = 'AADUserPrincipalName','TagYellow','TagGreen','TagPurple','DistinguishedName',
            @{name="AADUAssignedLicenses";expression={($_.AADUAssignedLicenses) -join ";"}},
            'ADEnabled','IsExoLicensed','MbxRetentionPolicy','MbxTotalItemSizeGB','AADUDirSyncEnabled',
            'OPRecipientType','OPRecipientTypeDetails','AADULastDirSyncTime','MbxLastLogonTime','ADLastLogonTime',
            'ParentOU','SiteOU','samaccountname',
            @{name="AADUSMTPProxyAddresses";expression={$_.AADUSMTPProxyAddresses.SmtpProxyAddresses -join ";"}},
            'ADCity','ADCompany','ADCountry','ADcountryCode','ADcreateTimeStamp','ADDepartment','ADDivision',
            'ADEmployeenumber','ADemployeeType','ADGivenName','ADmailNickname','ADMemberof','ADMobilePhone',
            'ADmodifyTimeStamp','ADOffice','ADOfficePhone','ADOrganization','ADphysicalDeliveryOfficeName',
            'ADPOBox','ADPostalCode','ADSMTPProxyAddresses','ADState','ADStreetAddress','ADSurname','ADTitle',
            'LicGrouppDN','MbxDatabase','MbxIssueWarningQuotaGB','MbxProhibitSendQuotaGB',
            'MbxProhibitSendReceiveQuotaGB','MbxServer',
            'MbxUseDatabaseQuotaDefaults','MbxExchangeGUID','Name','UserPrincipalName','WhenChanged',
            'WhenCreated','WhenMailboxCreated','XoMbxExchangeGUID','XoMbxTotalItemSizeGB','XoMbxLastLogonTime',
            'xoMbxWhenChanged','xoMbxWhenCreated','SplitBrain' ; 

        # EXO V2/3 steering params
        $EOMModName =  'ExchangeOnlineManagement' ;
        $EOMMinNoWinRMVersion = $MinNoWinRMVersion = '3.0.0' ; # support both names

        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
        write-verbose -verbose:$verbose "`$PSBoundParameters:`n$(($PSBoundParameters|out-string).trim())" ;
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        #if ($PSScriptRoot -eq "") {
        if( -not (get-variable -name PSScriptRoot -ea 0) -OR ($PSScriptRoot -eq '')){
            if ($psISE) { $ScriptName = $psISE.CurrentFile.FullPath } 
            elseif($psEditor){
                if ($context = $psEditor.GetEditorContext()) {$ScriptName = $context.CurrentFile.Path } 
            } elseif ($host.version.major -lt 3) {
                $ScriptName = $MyInvocation.MyCommand.Path ;
                $PSScriptRoot = Split-Path $ScriptName -Parent ;
                $PSCommandPath = $ScriptName ;
            } else {
                if ($MyInvocation.MyCommand.Path) {
                    $ScriptName = $MyInvocation.MyCommand.Path ;
                    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
                } else {throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$MyInvocation IS BLANK!" } ;
            };
            if($ScriptName){
                $ScriptDir = Split-Path -Parent $ScriptName ;
                $ScriptBaseName = split-path -leaf $ScriptName ;
                $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($ScriptName) ;
            } ; 
        } else {
            if($PSScriptRoot){$ScriptDir = $PSScriptRoot ;}
            else{
                write-warning "Unpopulated `$PSScriptRoot!" ; 
                $ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
            }
            if ($PSCommandPath) {$ScriptName = $PSCommandPath } 
            else {
                $ScriptName = $myInvocation.ScriptName
                $PSCommandPath = $ScriptName ;
            } ;
            $ScriptBaseName = (Split-Path -Leaf ((& { $myInvocation }).ScriptName))  ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
        } ;
        if(!$ScriptDir){
            write-host "Failed `$ScriptDir resolution on PSv$($host.version.major): Falling back to $MyInvocation parsing..." ; 
            $ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
            $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ; 
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;     
        } else {
            if(-not $PSCommandPath ){
                $PSCommandPath  = $ScriptName ; 
                if($PSCommandPath){ write-host "(Derived missing `$PSCommandPath from `$ScriptName)" ; } ;
            } ; 
            if(-not $PSScriptRoot  ){
                $PSScriptRoot   = $ScriptDir ; 
                if($PSScriptRoot){ write-host "(Derived missing `$PSScriptRoot from `$ScriptDir)" ; } ;
            } ; 
        } ; 
        if(-not ($ScriptDir -AND $ScriptBaseName -AND $ScriptNameNoExt)){ 
            throw "Invalid Invocation. Blank `$ScriptDir/`$ScriptBaseName/`ScriptNameNoExt" ; 
            BREAK ; 
        } ; 

        $smsg = "`$ScriptDir:$($ScriptDir)" ;
        $smsg += "`n`$ScriptBaseName:$($ScriptBaseName)" ;
        $smsg += "`n`$ScriptNameNoExt:$($ScriptNameNoExt)" ;
        $smsg += "`n`$PSScriptRoot:$($PSScriptRoot)" ;
        $smsg += "`n`$PSCommandPath:$($PSCommandPath)" ;  ;
        write-host $smsg ; 

        #*======v FUNCTIONS v======
        
        if(-not(get-command invoke-XoWrapper -EA 0 )){
            write-verbose "need the _func.ps1 to target, gcm doesn't do substrings, wo a wildcard" ; 
            if(-not($lmod = get-command import-XoW_func.ps1 -ea 0)){
                write-verbose "found local $($lmod.source), deferring to..." ; 
                ipmo -fo -verb $lmod ; 
            } else {
                #*------v import-XoW v------
                function import-XoW_func {
                    <#
                    .SYNOPSIS
                    import-XoW - import freestanding local invoke-XOWrapper_func.ps1 (back fill lack of xow support in verb-exo mod)
                    .NOTES
                    Version     : 1.0.0.
                    Author      : Todd Kadrie
                    Website     : http://www.toddomation.com
                    Twitter     : @tostka / http://twitter.com/tostka
                    CreatedDate : 2021-07-13
                    FileName    : import-XoW_func.ps1
                    License     : MIT License
                    Copyright   : (c) 2021 Todd Kadrie
                    Github      : https://github.com/tostka/verb-XXX
                    Tags        : Powershell
                    AddedCredit : REFERENCE
                    AddedWebsite: URL
                    AddedTwitter: URL
                    REVISIONS
                    * 10:32 AM 3/24/2023 flip wee lxoW into full function call
                    .DESCRIPTION
                    import-XoW - import freestanding local invoke-XOWrapper_func.ps1 (back fill lack of xow support in verb-exo mod)
                    .INPUTS
                    None. Does not accepted piped input.
                    .OUTPUTS
                    None.
                    .EXAMPLE
                    PS> import-XoW_func -users 'Test@domain.com','Test2@domain.com' -verbose  ;
                    Process an array of users, with default 'hunting' -LicenseSkuIds array.
                    .LINK
                    https://github.com/tostka/verb-exo
                    #>
                    [CmdletBinding()]
                    [Alias('lxoW')]
                    PARAM(
                        [Parameter(Mandatory=$false,HelpMessage="Tenant Tag (3-letter abbrebiation)[-TenOrg 'XYZ']")]
                        #[ValidateNotNullOrEmpty()]
                        [string]$ModuleName = 'invoke-XOWrapper_func.ps1'
                    ) ;
                    write-verbose "ipmo invoke-XOWrapper/xOW function" ;
                    if($iflpath = get-command $ModuleName | select -expand source){ 
                        if(test-path $iflpath){
                            $tMod = $iflpath ; 
                        }elseif(test-path (join-path -path 'C:\usr\work\o365\scripts\' -childpath $ModuleName)){
                            $tMod = (join-path -path 'C:\usr\work\o365\scripts\' -childpath $ModuleName) ;  
                        } else {throw 'Unable to locate xoW_func.ps1!' ;
                            break ;
                        } ;
                        if($tmod){
                            write-verbose 'Check for preloaded target function' ; 
                            if(-not(get-command (split-path $tmod -leaf).replace('_func.ps1','') -ea 0)){ 
                                write-verbose "`$tMod:$($tMod)" ;
                                Import-Module -force -verbose $tMod ;
                            } else { write-host "($tmod already loaded)" } ;
                        } else { write-warning "unable to resolve `$tmod!" } ;
                    } else { 
                        throw "Unable to locate $()" ; 
                    } ;  
                 }
                 #*------^ import-XoW ^------
            } ; ;
            import-XoW_func -verbose ;
        } ; 
                    
        #*======^ END FUNCTIONS ^======

        if(!(get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
        foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
        if(!(get-variable rgxPSAllUsersScope -ea 0)){
            $rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps(((d|m))*)1|dll)$" ;
        } ;
        if(!(get-variable rgxPSCurrUserScope -ea 0)){
            $rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps((d|m)*)1|dll)$" ;
        } ;
        #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($whatif) ;} ;
        # implement -TagFirst to lead with the TicketNumber (easier to group/sort ticket outputs if all named with ticket prefix)
        $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;TagFirst=$null; showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($whatif) ;} ;
        $pltSL.Tag = $Ticket ;
        $pltSL.TagFirst = $true ;
        if($script:PSCommandPath){
            if(($script:PSCommandPath -match $rgxPSAllUsersScope) -OR ($script:PSCommandPath -match $rgxPSCurrUserScope)){
                $bDivertLog = $true ;
                switch -regex ($script:PSCommandPath){
                    $rgxPSAllUsersScope{$smsg = "AllUsers"}
                    $rgxPSCurrUserScope{$smsg = "CurrentUser"}
                } ;
                $smsg += " context script/module, divert logging into [$budrv]:\scripts"
                write-verbose $smsg  ;
                if($bDivertLog){
                    if((split-path $script:PSCommandPath -leaf) -ne $cmdletname){
                        # function in a module/script installed to allusers|cu - defer name to Cmdlet/Function name
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
                    } else {
                        # installed allusers|CU script, use the hosting script name
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $script:PSCommandPath -leaf)) ;
                    }
                } ;
            } else {
                $pltSL.Path = $script:PSCommandPath ;
            } ;
        } else {
            if(($MyInvocation.MyCommand.Definition -match $rgxPSAllUsersScope) -OR ($MyInvocation.MyCommand.Definition -match $rgxPSCurrUserScope) ){
                 $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $script:PSCommandPath -leaf)) ;
            } elseif(test-path $MyInvocation.MyCommand.Definition) {
                $pltSL.Path = $MyInvocation.MyCommand.Definition ;
            } elseif($cmdletname){
                $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
            } else {
                $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$CMDLETNAME, FROM WHICH TO BUILD A START-LOG.PATH!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                BREAK ;
            } ;
        } ;
        write-verbose "start-Log w`n$(($pltSL|out-string).trim())" ;
        $logspec = start-Log @pltSL ;
        $error.clear() ;
        [array]$OutputFiles = @() ; 
        TRY {
            if($logspec){
                $logging=$logspec.logging ;
                $logfile=$logspec.logfile ;
                $transcript=$logspec.transcript ;
                $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                $startResults = start-Transcript -path $transcript ;
                $OutputFiles += $transcript ; 
                $OutputFiles += $logfile ; 
            } else {throw "Unable to configure logging!" } ;
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
        if ($PSCmdlet.MyInvocation.ExpectingInput) {
            $smsg = "Data received from pipeline input: '$($InputObject)'" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } else {
            #$smsg = "Data received from parameter input: '$($InputObject)'" ;
            $smsg = "(non-pipeline - param - input)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;

        #region BANNER ; #*------v BANNER v------
        $sBnr="#*======v $(${CmdletName}): v======" ;
        $smsg = $sBnr ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #endregion BANNER ; #*------^ END BANNER ^------
    
        $1stConn = $true ;

        if($TagYellow){
            $smsg = "-TagYellow specified: Adding 'TagYellow' to output fields" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
        } ; 
        if($TagGreen){
            $smsg = "-TagGreen specified: Adding 'TagGreen' to output fields" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        if($TagPurple){
            $smsg = "-TagPurple specified: Adding 'TagPurple' to output fields" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 

        # calc the rgxLicGrpName fr the existing $xxxmeta.rgxLicGrpDN: (gv tormeta).value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
        $rgxLicGrpName = (gv -name "$($tenorg)meta").value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
        # use the dn vers LicGrouppDN = $null ; # | ?{$_ -match $tormeta.rgxLicGrpDN}
        $rgxLicGrpDN = (gv -name "$($tenorg)meta").value.rgxLicGrpDN
        <# prestock licSkus
        $pltConn=[ordered]@{verbose=$false ; silent=$false} ;
        if($1stConn){
            $pltConn.silent = $false ;
        } else {
            $pltConn.silent = $true ;
        } ;
        rx10 @pltConn ; rxo @pltConn  ; cmsol @pltConn ;
        connect-ad -verbose:$false | out-null ;
        $1stConn = $false ;
        #>

        #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======
        #*------v STEERING VARIS v------
        $useO365 = $true ;
        $useEXO = $true ; 
        $UseOP=$true ; 
        $UseExOP=$true ;
        $useForestWide = $true ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
        $UseOPAD = $true ; 
        $UseMSOL = $false ; # should be hard disabled now in o365
        $UseAAD = $true  ; 
        $useO365 = [boolean]($useO365 -OR $useEXO -OR $UseMSOL -OR $UseAAD)
        $UseOP = [boolean]($UseOP -OR $UseExOP -OR $UseOPAD) ;
        #*------^ END STEERING VARIS ^------
        #region useO365 ; #*------v useO365 v------
        #$useO365 = $false ; # non-dyn setting, drives variant EXO reconnect & query code
        #if($CloudFirst){ $useO365 = $true } ; # expl: steering on a parameter
        if($useO365){
            #region GENERIC_EXO_CREDS_&_SVC_CONN #*------v GENERIC EXO CREDS & SVC CONN BP v------
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
            # -UserRole 'CSVC','SID'
            #$pltGTCred=@{TenOrg=$TenOrg ;userrole='CSVC','SID'; verbose=$($verbose)} ;
            # -UserRole 'SID','CSVC'
            $pltGTCred=@{TenOrg=$TenOrg ;userrole='SID','CSVC'; verbose=$($verbose)} ;
            $smsg = "get-TenantCredentials w`n$(($pltGTCred|out-string).trim())" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            if($o365Cred=(get-TenantCredentials @pltGTCred)){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name cred$($tenorg) -scope Script -ea 0 ){ remove-Variable -Name cred$($tenorg) -scope Script } ;
                New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
                $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatettus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ;
            <### CALLS ARE IN FORM: (cred$($tenorg))
            $pltRXO = @{Credential = $Credential ; verbose = $($verbose) ; }
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; Silent = $true ;} ;
            if ($script:useO365v2) { reconnect-eXO2 @pltRXO }
            else { reconnect-EXO @pltRXO } ;
            # or with Tenant-specific cred($Tenorg) lookup
            #$pltRXO creds & .username can also be used for AzureAD connections
            Connect-AAD @pltRXO ;
            ###>
            # configure splat for connections: (see above useage)
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; 
            } ;
            if((gcm Reconnect-EXO2).Parameters.keys -contains 'silent'){
                $pltRxo.add('Silent',$false) ;
            } ; 
            #region EOMREV ; #*------v EOMREV Check v------
            #$EOMmodname = 'ExchangeOnlineManagement' ;
            $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
            if($xmod = Get-Module $EOMmodname -ErrorAction Stop){ } else {
                $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                Try {
                    Import-Module @pltIMod | out-null ;
                    $xmod = Get-Module $EOMmodname -ErrorAction Stop ;
                } Catch {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = $ErrTrapd.Exception.Message ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Break ;
                } ;
            } ; # IsImported
            if([version]$xmod.version -ge $MinNoWinRMVersion){$MinNoWinRMVersion = $xmod.version.tostring() ;}
            [boolean]$UseConnEXO = [boolean]([version]$xmod.version -ge $MinNoWinRMVersion) ; 
            #endregion EOMREV ; #*------^ END EOMREV Check  ^------
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
            if ($script:useEXOv2 -OR $useEXOv2) { reconnect-eXO2 @pltRXO }
            else { reconnect-EXO @pltRXO } ;
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
                $statusdelta = ";ERROR";
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
            if((gcm Reconnect-Ex2010).Parameters.keys -contains 'silent'){
                $pltRX10.add('Silent',$false) ;
            } ;
            Reconnect-Ex2010 @pltRX10 ; # local org conns
            #$pltRx10 creds & .username can also be used for local ADMS connections
            ###>
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; 
            } ;
            if((gcm Reconnect-Ex2010).Parameters.keys -contains 'silent'){
                $pltRX10.add('Silent',$false) ;
            } ;

            # defer cx10/rx10, until just before get-recipients qry
            #endregion GENERIC_EXOP_CREDS_&_SRVR_CONN #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
            # connect to ExOP X10
            if($useEXOP){
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
                    write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                    $smsg += $ErrTrapd.Exception.Message ;
                    if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    CONTINUE ;
                } ;
            } else { 
        
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


        #region UseOPAD #*------v UseOPAD v------
        if($UseOP -OR $UseOPAD){
            #region GENERIC_ADMS_CONN_&_XO #*------v GENERIC ADMS CONN & XO  v------
            $smsg = "(loading ADMS...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # always capture load-adms return, it outputs a $true to pipeline on success
            $ADMTLoaded = load-ADMS -Verbose:$FALSE ;
            <# disabled/fw-borked cross-org code
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
                #-=-record a STATUS=-=-=-=-=-=-=
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                #-=-=-=-=-=-=-=-=
                $smsg = "Unable to detect POPULATED `$global:ADPsDriveNames!`n(should have multiple values, resolved to $()"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ;
            #>
            TRY {
                if(-not(Get-ADDomain).DNSRoot){
                    $smsg = "Missing AD Connection! (no (Get-ADDomain).DNSRoot returned)" ; 
                    throw $smsg ; 
                    $smsg | write-warning  ; 
                } ; 
                if($useForestWide){
                    #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE AD GC QRY SUPPORT v------
                    $smsg = "(`$useForestWide:$($useForestWide)):Enabling AD Forestwide)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #TK 9:44 AM 10/6/2022 need org wide for rolegrps in parent dom (only for onprem RBAC, not EXO)
                    $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;        
                    #endregion  ; #*------^ END  OPTIONAL CODE TO ENABLE FOREST-WIDE AD GC QRY SUPPORT  ^------
                } ;    
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = $ErrTrapd ;
                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                $smsg += $ErrTrapd.Exception.Message ;
                if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                CONTINUE ;
            } ;        
            #endregion GENERIC_ADMS_CONN_&_XO #*------^ END GENERIC ADMS CONN & XO ^------
        } else {
            $smsg = "(`$UseOP:$($UseOP)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        }  ;
        #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
        #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller=get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
        # use new get-GCFastXO cross-org dc finde
        # default to Op_ExADRoot forest from $TenOrg Meta
        #if($UseOP -AND -not $domaincontroller){
        if($UseOP -AND -not (get-variable domaincontroller -ea 0)){
            #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((gv -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};
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
            connect-msol @pltRXO ;
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
            Connect-AAD @pltRXO ;
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
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
        else { reconnect-EXO @pltRXO } ;
        # reenable VerbosePreference:Continue, if set, during mod loads
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        #>
        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        #endregion SERVICE_CONNECTIONS #*======^ END SERVICE_CONNECTIONS ^======

        if($useEXO){
            # splice in xow support
            #function lxoW { write-verbose "ipmo xOW function" ; if(test-path D:\scripts\xoW_func.ps1){$tMod = 'D:\scripts\xoW_func.ps1'}elseif(test-path C:\usr\work\o365\scripts\xoW_func.ps1){$tMod = 'C:\usr\work\o365\scripts\xoW_func.ps1' } else {throw 'Unable to locate xoW_func.ps1!' ;break ;} ; if($tmod){ if(!(gcm $tmod)){ write-verbose "`$tMod:$($tMod)" ; ipmo -fo -verb $tMod ; } else { write-host "($tmod already loaded)" } ;  } else { \write-warning "unable to resolve `$tmod!" } ;  } ; lxoW -verbose ; 
            # defers to import-XoW up in functions block
        } ; 

        $1stConn = $false ;
        $rgxHashTableTypeName = "(System.Collections.Hashtable|System.Collections.Specialized.OrderedDictionary)" ;
        #-=-=-=-=-=-=-=-=
        $pltGLPList=[ordered]@{
            TenOrg= $TenOrg;
            verbose=$($VerbosePreference -eq "Continue") ;
            credential= $pltRXO.credential ;
            #(Get-Variable -name cred$($tenorg) ).value ;
        } ;
        $smsg = "$($tenorg):get-AADlicensePlanList w`n$(($pltGLPList|out-string).trim())" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $objRet = $null ;
        $objRet = get-AADlicensePlanList @pltGLPList ;
        if( ($objRet|Measure-Object).count -AND $objRet.GetType().FullName -match $rgxHashTableTypeName ){
            $smsg = "get-AADlicensePlanList:$($tenorg):returned populated LicensePlanList" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $licensePlanListHash = $objRet ;
        } else {
            $smsg = "get-AADlicensePlanList:$($tenorg)FAILED TO RETURN populated [hashtable] LicensePlanList" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            THROW $SMSG ;
            break ;
        } ;
        #-=-=-=-=-=-=-=-=
        #-=-=-=-=-=-=-=-=
        $smsg = "get-MailboxDatabaseQuotas:Qry onprem org hashtable of mailboxquotas per mailboxdatabase" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        $pltGMDQ=[ordered]@{
            TenOrg= $TenOrg;
            verbose=$($VerbosePreference -eq "Continue") ;
            #credential= $pltRXO.credential ;
            credential= $pltRX10.credential ;
            #(Get-Variable -name cred$($tenorg) ).value ;
        } ;
        $smsg = "$($tenorg):get-MailboxDatabaseQuotas w`n$(($pltGMDQ|out-string).trim())" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $objRet = $null ;
        $objRet = get-MailboxDatabaseQuotas @pltGMDQ ;
        if( ($objRet|Measure-Object).count -AND $objRet.GetType().FullName -match $rgxHashTableTypeName ){
            $smsg = "get-MailboxDatabaseQuotas:$($tenorg):returned populated [hashtable].MailboxDatabaseQuotas" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $mdbquotas = $objRet ;
        } else {
            $smsg = "get-MailboxDatabaseQuotas:$($tenorg):FAILED TO RETURN populated [hashtable] MailboxDatabaseQuotas of " ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            THROW $SMSG ;
            break ;
        } ;
        #-=-=-=-=-=-=-=-=
        #-=-=-=-=-=-=-=-=
        $pltGXML=[ordered]@{
            #TenOrg= $TenOrg;
            verbose=$($VerbosePreference -eq "Continue") ;
            #credential= $pltRXO.credential ;
            #(Get-Variable -name cred$($tenorg) ).value ;
        } ;
        $smsg = "$($tenorg):get-ExoMailboxLicenses w`n$(($pltGXML|out-string).trim())" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $objRet = $null ;
        $objRet = get-ExoMailboxLicenses @pltGXML ;
        #$objRet = invoke-XOWrapper {get-ExoMailboxLicenses @pltGXML} -credential $pltRXO.Credential -credentialOP $pltRX10.Credential ; ;
        # ^ not needed get-EXOMailboxLicenses is static text parse, canned material. No queries to XO at all.
        if( ($objRet|Measure-Object).count -AND $objRet.GetType().FullName -match $rgxHashTableTypeName ){
            $smsg = "get-ExoMailboxLicenses:$($tenorg):returned populated ExMbxLicenses" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $ExMbxLicenses = $objRet ;
        } else {
            $smsg = "get-ExoMailboxLicenses:$($tenorg):FAILED TO RETURN populated [hashtable] ExMbxLicenses" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            THROW $SMSG ;
            break ;
        } ;
        $smsg = "$(($ExMbxLicenses.Values|measure).count) EXO UserMailbox-supporting License summaries returned)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #-=-=-=-=-=-=-=-=

        $Rpt = @() ;

        # check if using Pipeline input or explicit params:
        if ($PSCmdlet.MyInvocation.ExpectingInput) {
            write-verbose "Data received from pipeline input: '$($InputObject)'" ;
        } else {
            # doesn't actually return an obj in the echo
            #write-verbose "Data received from parameter input: '$($InputObject)'" ;
        } ;

    } ;  # BEGIN-E
    PROCESS {
        $Error.Clear() ;
        # call func with $PSBoundParameters and an extra (includes Verbose)
        #call-somefunc @PSBoundParameters -anotherParam

        # - Pipeline support will iterate the entire PROCESS{} BLOCK, with the bound - $array -
        #   param, iterated as $array=[pipe element n] through the entire inbound stack.
        # $_ within PROCESS{}  is also the pipeline element (though it's safer to declare and foreach a bound $array param).

        # - foreach() below alternatively handles _named parameter_ calls: -array $objectArray
        # which, when a pipeline input is in use, means the foreach only iterates *once* per
        #   Process{} iteration (as process only brings in a single element of the pipe per pass)

        #$1stConn = $true ;
        $ttl = ($Mailboxes|measure).count ; $Procd = $SplitBs = 0 ;
        foreach ($mbx in $Mailboxes){
            # null all working varis
            $adu = $ombx = $mbxstat = $AADUser = $xoMbx = $xmbxstat = $tLastLogon = $userList = $sku= $null;
            $isInvalid=$false ;
            switch ($mbx.GetType().fullname){
                'System.String' {
                    # BaseType: System.Object
                    $smsg = "$($mbx) specified does not appear to be a proper Exchange OnPrem Mailbox object"
                    $smsg+= "`ndetected type:`n$(($mbx.GetType() | ft -a fullname,basetype|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $isInvalid=$true;
                    CONTINUE ;
                }
                'System.Management.Automation.PSObject' {
                    # BaseType: System.Object
                    $smsg = "(valid 'System.Management.Automation.PSObject)'" ;
                    if($verbose){
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                }
                'System.Object[]'{
                    # BaseType: System.Array
                    $smsg = "$($mbx) specified does not appear to be a proper Exchange OnPrem Mailbox object"
                    $smsg+= "`ndetected type:`n$(($mbx.GetType() | ft -a fullname,basetype|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $isInvalid=$true;
                    CONTINUE ;
                }
                default {
                    $smsg = "Unrecognized object type! "
                    $smsg+= "`ndetected type:`n$(($mbx.GetType() | ft -a fullname,basetype|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $isInvalid=$true;
                    CONTINUE ;
                }
            } ;
            $Procd ++ ;

            if(-not $isInvalid){
                $sBnrS="`n#*------v PROCESSING : ($($Procd)/$($ttl)) $($mbx.UserPrincipalName) (SB#:$($SplitBs)) v------" ;
                $smsg = "$((get-date).ToString('HH:mm:ss')):$($sBnrS)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                #rx10 ;
                #$pltConn=[ordered]@{verbose=$false ; silent=$false} ;
                if($1stConn){
                    $pltRX10.silent = $pltRXO.silent = $false ;
                } else {
                    $pltRX10.silent = $pltRXO.silent =$true ;
                } ;
                if($pltRX10){ReConnect-Ex2010 @pltRX10 }
                else {ReConnect-Ex2010 }
                if($useAAD){
                    if($pltRXO){Connect-AAD @pltRXO}
                    else {Connect-AAD } ; 
                } ;

                $error.clear() ;
                
                # 12:50 PM 4/7/2023 rem, going to bracket each call, and recover wo abandoning  entire user process get-aduser fails (deleted users) are dropping entire user reporting
                    $hSummary=[ordered]@{
                        AADUAssignedLicenses = $null ;
                        AADUDirSyncEnabled = $null ;
                        AADULastDirSyncTime = $null ;
                        AADUserPrincipalName = $null ;
                        AADUSMTPProxyAddresses = $null ;
                        ADCity = $null ;
                        ADCompany = $null ;
                        ADCountry = $null ;
                        ADcountryCode = $null ;
                        ADcreateTimeStamp = $null ;
                        ADDepartment = $null ;
                        ADDivision = $null ;
                        ADEmployeenumber = $null ;
                        ADemployeeType = $null ;
                        ADEnabled = $null ;
                        ADGivenName = $null ;
                        ADmailNickname = $null ;
                        ADMemberof = $null ;
                        ADMobilePhone = $null ;
                        ADmodifyTimeStamp = $null ;
                        ADOffice = $null ;
                        ADOfficePhone = $null ;
                        ADOrganization = $null ;
                        ADphysicalDeliveryOfficeName = $null ;
                        ADPOBox = $null ;
                        ADPostalCode = $null ;
                        ADSMTPProxyAddresses = $null ;
                        ADState = $null ;
                        ADStreetAddress = $null ;
                        ADSurname = $null ;
                        ADTitle = $null ;
                        ADLastLogonTime = $null ; 
                        DistinguishedName = $mbx.DistinguishedName;
                        IsExoLicensed = $null ;
                        LicGrouppDN = $null ; # | ?{$_ -match $xxxmeta.rgxLicGrpDN}
                        MbxDatabase = $null ;
                        MbxIssueWarningQuotaGB = $null ;
                        MbxLastLogonTime = $null ;
                        MbxProhibitSendQuotaGB = $null ;
                        MbxProhibitSendReceiveQuotaGB = $null ;
                        MbxRetentionPolicy = $null ;
                        MbxServer = $null ;
                        MbxTotalItemSizeGB = $null ;
                        MbxUseDatabaseQuotaDefaults = $null ;
                        MbxExchangeGUID = $null ;
                        OPRecipientType = $null ; 
                        OPRecipientTypeDetails = $null ; 
                        Name = $mbx.name;
                        ParentOU = (($mbx.distinguishedname.tostring().split(',')) |select -skip 1) -join ',' ;
                        samaccountname = $mbx.samaccountname;
                        SiteOU = ($mbx.distinguishedname.tostring().split(','))[(-1*$SiteOUNestingLevel)..-1] -join ',' ;
                        UserPrincipalName = $mbx.UserPrincipalName;
                        WhenChanged = $mbx.WhenChanged ;
                        WhenCreated  = $mbx.WhenCreated ;
                        WhenMailboxCreated = $null ;
                        XoMbxExchangeGUID = $null ; 
                        XoMbxTotalItemSizeGB = $null ; 
                        XoMbxLastLogonTime = $null ;
                        xoMbxWhenChanged = $mbx.WhenChanged ;
                        xoMbxWhenCreated  = $mbx.WhenCreated ;
                        SplitBrain = $null ; 
                    } ;

                    # 9:38 AM 4/5/2023 ADD TagYellow, TagGreen, TagPurple param support
                    if($TagYellow){
                        $hSummary.add('TagYellow',$null) ; 
                    } ; 
                    if($TagGreen){
                        $hSummary.add('TagGreen',$null) ; 
                    } ; 
                    if($TagPurple){
                        $hSummary.add('TagPurple',$null) ; 
                    } ; 

                    # 12:12 PM 4/10/2023 need an onprem recipient first - only way to spot migrat4ed rmbx, or even DN shifts, which break gmbxstat and gadu on DN
                   $pltGrcp=[ordered]@{
                        identity = $mbx.UserPrincipalName ;
                        ErrorAction = 'STOP' ;
                        verbose = ($VerbosePreference -eq "Continue") ;
                    } ;
                    $smsg = "get-recipient w`n$(($pltGrcp|out-string).trim())" ;
                    if($verbose){
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;

                    #$error.clear() ;
                    # 2:16 PM 4/10/2023: after all these years, turns out: https://www.michev.info/blog/post/1415/error-handling-in-exchange-remote-powershell-sessions
                    <# can't get get-recipient to fire the catch, it just dumps the error to pipeline, and skips the catch c\ode.
                    TRY{ 
                        $OpRcp = get-recipient @pltGrcp
                    # 'missing' CATCH

                    } CATCH {
                        # or just do idiotproof: Write-Warning -Message $_.Exception.Message ;
                        $ErrTrapd=$Error[0] ;
                        $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $smsg = $ErrTrapd.Exception.Message ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #-=-record a STATUSWARN=-=-=-=-=-=-=
                        $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                        if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                        if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                        #-=-=-=-=-=-=-=-=
                        $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        CONTINUE #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                    } ; 
                    #>

                    $OpRcp = $null ; 
                    if($OpRcp = get-recipient @pltGrcp){
                        # $objRet = $null ;
                        $smsg = "found matching recipient:`n$(($OpRcp | ft -a name,recipienttype,recipienttypedetails|out-string).trim())" ; 

                        switch($OpRcp.recipienttypedetails){
                            'UserMailbox' {
                                $smsg += "`nOf expected RecipientTypeDetails:$($OpRcp.recipienttypedetails)" ; 
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            } ; 
                            'RemoteUserMailbox' {
                                $smsg += "`nOf UNEXPECTED RecipientTypeDetails:$($OpRcp.recipienttypedetails)"
                                $smsg += "`nUSER APPEARS TO BE A *MIGRATED* CLOUD MAILBOX, SINCE ORIGINAL POLL!" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                                $pltRMbx=[ordered]@{
                                    identity = $mbx.UserPrincipalName ;
                                    ErrorAction = 'STOP' ;
                                    verbose = ($VerbosePreference -eq "Continue") ;
                                } ;
                                $smsg = "get-Remotemailbox w`n$(($pltRMbx|out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                                $objRet = $null ;
                                TRY {
                                    $objRet = get-remotemailbox @pltRMbx ; 
                                } CATCH {
                                    # or just do idiotproof: Write-Warning -Message $_.Exception.Message ;
                                    $ErrTrapd=$Error[0] ;
                                    $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    $smsg = $ErrTrapd.Exception.Message ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    #-=-record a STATUSWARN=-=-=-=-=-=-=
                                    $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                                    if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                                    if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                                    #-=-=-=-=-=-=-=-=
                                    $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    # 8:48 AM 4/11/2023 don't continue: advances loop, and this is one missing component - on a migrated user - we want the rest reported on!
                                    #CONTINUE #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                                } ; 

                                if( ($objRet|Measure-Object).count -AND $objRet.GetType().FullName -eq 'System.Management.Automation.PSObject' ){
                                    $smsg = "Get-RemoteMailbox:$($tenorg):returned populated RemoteMailbox" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                
                                    $Rmbx = $objRet ; 

                                } else {G2084

                                    $smsg = "Get-RemoteMailbox:$($tenorg):FAILED TO RETURN populated 'System.Management.Automation.PSObject' mbx!" ;
                                    if($OpRcp.recipienttypedetails -eq 'RemoteUserMailbox'){
                                        $smsg += "`nGiven `$OpRcp.recipienttypedetails:'RemoteUserMailbox', this is an expected result: Someone has migrated the mailbox to cloud" ; 
                                        $smsg += "`n(if both a Mailbox & xoMailbox existed: That would reflect a split-brain condition)" ; 
                                    } else { 
                                        $smsg += "`nNo Op Mbx, and `$OpRcp.recipienttypedetails:$($OpRcp.recipienttypedetails)" ; 
                                        $smsg += "`nThis is a combo this code is not currently written to accomodate" ; 
                                    } ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }
                                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    # 12:32 PM 4/11/2023 rem throw's they're exiting script entirely (vs just echoing an error to console)
                                    #THROW $SMSG ;
                                } ;

                            } ; 
                            'Mailuser' {
                                $smsg += "`nOf UNEXPECTED RecipientTypeDetails:$($OpRcp.recipienttypedetails)"
                                $smsg += "`nUSER APPEARS TO HAVE BECOME AN ORPHANED POINTER TO CLOUD, SINCE ORIGINAL POLL!" ; 
                                $SMSG += "`n(as a RemoteMailbox should not flip to MailUser, and this user had MailboxUser type in past)" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            } 
                            default {
                                $smsg += "`nOf UNEXPECTED RecipientTypeDetails:$($OpRcp.recipienttypedetails)"
                                $smsg += "`nThis code is not written to adadequately handle this combo!" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            }
                        } ; 

                        # test DN hasn't changed
                        if($mbx.DistinguishedName -eq $OpRcp.DistinguishedName){
                            $smsg = "(confirmed DistName still matches original spec)" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        } else { 
                            $smsg = "RECIPIENT DISTINGUISHEDNAME - $($OpRcp.DistinguishedName) - " ;
                            $smsg += "`n- NO LONGER MATCHES ORIGINAL DN - $($mbx.DistinguishedName)!" ;
                            $SMSG += "`n ADUser object has been MOVED in the hierarchy!" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            
                            $smsg = "RECHECKING FOR CURRENT MAILBOX AND SPECIFICATIONS!" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                            $pltGMbx=[ordered]@{
                                identity = $mbx.UserPrincipalName ;
                                ErrorAction = 'STOP' ;
                                verbose = ($VerbosePreference -eq "Continue") ;
                            } ;
                            $smsg = "get-mailbox w`n$(($pltGMbx|out-string).trim())" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            $objRet = $null ;
                            TRY {
                                $objRet = get-mailbox @pltGMbx ; 
                            } CATCH {
                                # or just do idiotproof: Write-Warning -Message $_.Exception.Message ;
                                $ErrTrapd=$Error[0] ;
                                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $smsg = $ErrTrapd.Exception.Message ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #-=-record a STATUSWARN=-=-=-=-=-=-=
                                $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                                if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                                if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                                #-=-=-=-=-=-=-=-=
                                $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                # 8:48 AM 4/11/2023 don't continue: advances loop, and this is one missing component - on a migrated user - we want the rest reported on!
                                #CONTINUE #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                            } ; 

                            if( ($objRet|Measure-Object).count -AND $objRet.GetType().FullName -eq 'System.Management.Automation.PSObject' ){
                                $smsg = "Get-Mailbox:$($tenorg):returned populated mbx, where there is *no* ADUser!" ;
                                $smsg += "`nShould not be a possible combo!" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                
                                # assign inbound $mbx to $ombx, update $mbx to the fresh info
                                $ombx = $mbx ; 
                                $mbx = $objRet ; 

                            } else {
                                $smsg = "Get-Mailbox:$($tenorg):FAILED TO RETURN populated 'System.Management.Automation.PSObject' mbx!" ;
                                if($OpRcp.recipienttypedetails -eq 'RemoteUserMailbox'){
                                    $smsg += "`nGiven `$OpRcp.recipienttypedetails:'RemoteUserMailbox', this is an expected result: Someone has migrated the mailbox to cloud" ; 
                                    $smsg += "`n(if both a Mailbox & xoMailbox existed: That would reflect a split-brain condition)" ; 
                                } else { 
                                    $smsg += "`nNo Op Mbx, and `$OpRcp.recipienttypedetails:$($OpRcp.recipienttypedetails)" ; 
                                    $smsg += "`nThis is a combo this code is not currently written to accomodate" ; 
                                } ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                # 12:32 PM 4/11/2023 rem throw's they're exiting script entirely (vs just echoing an error to console)
                                #THROW $SMSG ;
                            } ;

                        } ; 
                    } else { 
                        $smsg = "NO MATCHING GET-RECIPIENT -id $($pltGrcpidentity)!" ; 
                        $smsg += "`nUSER MAY HAVE BEEN DELETED OR OTHER WISE REMOVED FROM AD!" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                    } ; 


                    <# $prpAadu = 'UserPrincipalName','GivenName','Surname','DisplayName','AccountEnabled','Description','PhysicalDeliveryOfficeName','JobTitle','AssignedLicenses','Department','City','State','Mail','MailNickName','LastDirSyncTime','OtherMails','ProxyAddresses' ;
                    #>
                    $pltGadu=[ordered]@{
                        identity = $mbx.DistinguishedName ;
                        ErrorAction='STOP' ;
                        properties=$prpADU;
                        verbose = ($VerbosePreference -eq "Continue") ;
                    } ;
                    if($OpRcp.RecipientType -ne 'UserMailbox'){
                        $smsg = "Missing local Mailbox: `$OpRcp.RecipientType -ne 'UserMailbox',`n resetting to use DN of `$OpRcp, over original `$mbx.dn" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        $pltGadu.identity = $OpRcp.DistinguishedName ; 
                    } ; 
                    $smsg = "get-aduser w`n$(($pltGadu|out-string).trim())" ;
                    if($verbose){
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;

                    $error.clear() ;
                    TRY{ 
                        $adu = get-aduser @pltGadu
                    } CATCH [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{
                        # not found err - get-aduser is on dn; found some moved in tree, breaking dn, retry on another attrib; the trailing .xxx of the modern fname.lname.nnnnn is the SAMACCTNAME (pirnr?)
                        # retry on mbx.Alias, should be the samaccountname
                        $pltGadu=[ordered]@{
                            identity = $mbx.alias ;
                            ErrorAction='STOP' ;
                            properties=$prpADU;
                            verbose = ($VerbosePreference -eq "Continue") ;
                        } ;
                        $smsg = "get-aduser on DN failed: Retrying on `$mbx.Alias:$($mbx.alias)`n(against reloc of user, breaking DN)" ; 
                        $smsg += "`nget-aduser w`n$(($pltGadu|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                        TRY{ 
                            $adu = get-aduser @pltGadu
                        } CATCH {
                            # or just do idiotproof: Write-Warning -Message $_.Exception.Message ;
                            $ErrTrapd=$Error[0] ;
                            $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $smsg = $ErrTrapd.Exception.Message ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #-=-record a STATUSWARN=-=-=-=-=-=-=
                            $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                            if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                            if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                            #-=-=-=-=-=-=-=-=
                            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            # 8:48 AM 4/11/2023 don't continue: advances loop, and this is one missing component - on a migrated user - we want the rest reported on!
                            #CONTINUE #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                        } ; 

                    } CATCH {
                        # or just do idiotproof: Write-Warning -Message $_.Exception.Message ;
                        $ErrTrapd=$Error[0] ;
                        $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $smsg = $ErrTrapd.Exception.Message ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #-=-record a STATUSWARN=-=-=-=-=-=-=
                        $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                        if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                        if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                        #-=-=-=-=-=-=-=-=
                        $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        # 8:48 AM 4/11/2023 don't continue: advances loop, and this is one missing component - on a migrated user - we want the rest reported on!
                        #CONTINUE #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                    } ; 

                    #| select $selectADU ;
                    $pltGMStat=[ordered]@{
                        #identity = $mbx.UserPrincipalName ;
                        # they've got 2 david.smith@toro.com's onboarded, both with same UPN, shift to DN it's more specific
                        identity = $mbx.DistinguishedName ;
                        ErrorAction='STOP' ;
                        verbose = ($VerbosePreference -eq "Continue") ;
                    } ;
                    if($OpRcp.RecipientType -ne 'UserMailbox'){
                        #$smsg = "Missing local Mailbox: `$OpRcp.RecipientType -ne 'UserMailbox',`n resetting to use DN of `$OpRcp, over original `$mbx.dn" ; 
                        $smsg = "Missing local Mailbox: `$OpRcp.RecipientType -ne 'UserMailbox',`nSKIPPING ATTEMPT TO RETRIEVE ONPREM:Get-MailboxStatistics " ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        $pltGMStat.identity = $OpRcp.DistinguishedName ; 
                    } else {
                    
                        $smsg = "Get-MailboxStatistics  w`n$(($pltGMStat|out-string).trim())" ;
                        if($verbose){
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        Reconnect-Ex2010 @pltRX10  ;  
                        $error.clear() ;
                        TRY{
                            $mbxstat = Get-MailboxStatistics @pltGMStat ;
                        } CATCH {
                            # or just do idiotproof: Write-Warning -Message $_.Exception.Message ;
                            $ErrTrapd=$Error[0] ;
                            $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $smsg = $ErrTrapd.Exception.Message ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #-=-record a STATUSWARN=-=-=-=-=-=-=
                            $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                            if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                            if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                            #-=-=-=-=-=-=-=-=
                            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            # 8:48 AM 4/11/2023 don't continue: advances loop, and this is one missing component - on a migrated user - we want the rest reported on!
                            #CONTINUE #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                        } ; 
                    } ; 

                    if($ADLastLogon){
                        # blank comes through as: Sunday, December 31, 1600 6:00:00 PM
                        # g format (short date), outputs: 10/15/2012 3:13 PM
                        #$tLastLogon = [datetime]::FromFileTime($adu.LastLogon).ToString('g') ; 
                        if($adu.LastLogon){
                            $tLastLogon = [datetime]::FromFileTime($adu.LastLogon) ; 
                        } elseif($adu.lastLogonTimestamp){
                            $tLastLogon = [datetime]::FromFileTime($adu.lastLogonTimestamp) ; 
                        } else { 
                            $smsg = "(neither adu.LastLogon nor adu.lastLogonTimestamp was populated, to determing ADU.LastLogon)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ; 
                        if($tLastLogon -gt (get-date '12/31/1600')){
                            $hSummary.ADLastLogonTime =  $tLastLogon.ToString('g');
                        } else {
                            $hSummary.ADLastLogonTime = $null ;
                            write-verbose "(1600 invalid LastLogon resolved date)" ; 
                        } ;
                        
                    } ; 

                    # do direct lookup of AADU on specified eml (assumed to be UPN, if it came out of ADC error log)
                    $pltGAADU=[ordered]@{
                        ObjectId = $mbx.UserPrincipalName ;
                        ErrorAction = 'STOP' ;
                        verbose = ($VerbosePreference -eq "Continue") ;
                    } ;

                    if($mbx.UserPrincipalName -ne $adu.UserPrincipalName){
                        $smsg = "`$mbx.UserPrincipalName:$($mbx.UserPrincipalName)"  ; 
                        $smsg += "`nDOES NOT EQUAL DISCOVERED:" ; 
                        $smsg += "`$adu.UserPrincipalName:$($adu.UserPrincipalName)" ;
                        $smsg += "`ndeferring get-AzureAdUser to use of the `$adu.userprincipalname" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        $pltGAADU.ObjectId = $adu.UserPrincipalName ;
                    }; 

                    $smsg = "Get-AzureADUser on UPN:`n$(($pltGAADU|out-string).trim())" ;
                    $smsg = $recursetag,$smsg -join '' ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    # check for name changes since $mbx dumped:
                    if ($mbx.PrimarySmtpAddress -ne $oprcp.PrimarySmtpAddress){
                        $smsg = "original `$mbx.PrimarySmtpAddress:`n$($mbx.PrimarySmtpAddress)" ; 
                        $smsg += "`nDOES NOT match discovered `$OpRcp.PrimarySmtpAddress:`n$($OpRcp.PrimarySmtpAddress)" ; 
                        $smsg += "`nTHE USER APPEARS TO AHVE UNDERGONE A NAME OR DOMAIN CHANGE!" ; 
                        $smsg +="`nThis *will* impede Get-AzureAdUser lookups!" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                        
                    } ; 
                    $error.clear() ;
                    TRY {
                        $AADUser = Get-AzureADUser @pltGAADU
                    } CATCH {
                        # or just do idiotproof: Write-Warning -Message $_.Exception.Message ;
                        $ErrTrapd=$Error[0] ;
                        $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $smsg = $ErrTrapd.Exception.Message ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #-=-record a STATUSWARN=-=-=-=-=-=-=
                        $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                        if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                        if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                        #-=-=-=-=-=-=-=-=
                        $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        # 8:48 AM 4/11/2023 don't continue: advances loop, and this is one missing component - on a migrated user - we want the rest reported on!
                        #CONTINUE #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                    } ;

                    if($AADUser){
                        if(($AADUser|measure).count -gt 1){
                            $smsg = "MULTIPLE AZUREADUSERS **SAME USEPRINCIPALNAME**!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        foreach($aadu in $AADUser){
                            $smsg = "`n===`n$(($aadu|fl $prpAadu  | out-string).trim())" ;
                            # select smtpproxyaddresses out:
                            $smsg +="`nSMTPProxyAddresses:`n$(($aadu | select $prpAxDUserSmtpProxyAddr | select -expand SMTPProxyAddresses| sort |out-string).trim())" ;
                            $smsg += "`nProvisioningErrors :`n$(($aadu|select -expand provisioningerrors | out-string).trim())" ;
                            $smsg = $recursetag,$smsg -join '' ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            $error.clear() ;
                            TRY {
                                $userList = $aadu | Select -ExpandProperty AssignedLicenses | Select SkuID  ;
                                $userLicenses=@() ;
                                $userList | ForEach {
                                    $sku=$_.SkuId ;
                                    $userLicenses+=$licensePlanListHash[$sku].SkuPartNumber ;
                                } ;
                            } CATCH {
                                # or just do idiotproof: Write-Warning -Message $_.Exception.Message ;
                                $ErrTrapd=$Error[0] ;
                                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $smsg = $ErrTrapd.Exception.Message ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #-=-record a STATUSWARN=-=-=-=-=-=-=
                                $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                                if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                                if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                                #-=-=-=-=-=-=-=-=
                                $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                # 8:48 AM 4/11/2023 don't continue: advances loop, and this is one missing component - on a migrated user - we want the rest reported on!
                                #CONTINUE #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                            } ; 
                            $hSummary.AADUAssignedLicenses = $userLicenses ;

                            if($LicensedMail){
                                $IsExoLicensed = $false ;
                                # test for presence of a common mailbox-supporting lic, (or (Shared|Room|Equipment)Mailbox recipienttypedetail)
                                foreach($pLic in $hSummary.AADUAssignedLicenses){
                                    $smsg = "--(LicSku:$($plic): checking EXO UserMailboxSupport)" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    # array contans chk
                                    #if($ExMbxLicenses.SKU -contains $pLic){
                                    # indexed hash lookup:
                                    if($ExMbxLicenses[$plic]){
                                        $hSummary.IsExoLicensed = $true ;
                                        $smsg = "$($mbx.userprincipalname) HAS EXO UserMailbox-supporting License:$($ExMbxLicenses[$sku].SKU)|$($ExMbxLicenses[$sku].Label)" ;
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        BREAK ; # no sense running whole set, break on 1st mbx-support match
                                    } ;
                                } ;
                                # confirm at least one of the assigned lic's chkd is EXO mbx-supporting
                                if(-not $hSummary.IsExoLicensed){
                                    $smsg = "$($mbx.userprincipalname) WAS FOUND TO HAVE *NO* EXO UserMailbox-supporting License!" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                } ;

                                <# for unlicensed - or any, nail down if it has a memberof that should license
                                # use the dn vers LicGrouppDN = $null ; # | ?{$_ -match $tormeta.rgxLicGrpDN}
                                $rgxLicGrpDN = (gv -name "$($tenorg)meta").value.rgxLicGrpDN
                                #>
                                if($adu.memberof  | ?{$_ -match $rgxLicGrpDN}){
                                    $hSummary.LicGrouppDN = $adu.memberof  | ?{$_ -match $rgxLicGrpDN };
                                    $smsg = "Onprem memberof LicGrp matches:$($hSummary.LicGrouppDN)" ;
                                    if(-not $hSummary.IsExoLicensed){
                                        $smsg = "NON-EXO-LICENSE:" + $SMSG ;
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    } else {
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    } ;
                                } else {
                                    if($hSummary.IsExoLicensed){
                                        $smsg = "(no standard lic-grp memberof, appears to be direct-grant (E3?))" ;
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                        else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    } else {
                                        $smsg = "NO EVIDENCE $($ADU.userprincipalname) HAS FUNCTIONAL EXO MBX LIC, CHECK FOR EXO.MBX!" ;
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    } ;
                                };
                            } ;
                        } ;
                    } else {
                        $smsg = "=>Get-AzureADUser NOMATCH" ;
                        $smsg = $recursetag,$smsg -join '' ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;

                    # check splitbrain (op & XO mbx):
                    $pltGxMbx=[ordered]@{
                        identity = $mbx.UserPrincipalName ;
                        ErrorAction = 'SilentlyContinue' ;
                        verbose = ($VerbosePreference -eq "Continue") ;
                    } ;
                    $smsg = "get-xomailbox on UPN:`n$(($pltGxMbx|out-string).trim())" ;
                    $smsg = $recursetag,$smsg -join '' ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $objRet = $null ;
                    if($UseConnEXO){
                        $error.clear() ;
                        TRY {
                            $objRet = get-xomailbox @pltGxMbx ; 
                        } CATCH {
                            # or just do idiotproof: Write-Warning -Message $_.Exception.Message ;
                            $ErrTrapd=$Error[0] ;
                            $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $smsg = $ErrTrapd.Exception.Message ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #-=-record a STATUSWARN=-=-=-=-=-=-=
                            $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                            if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                            if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                            #-=-=-=-=-=-=-=-=
                            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            # 8:48 AM 4/11/2023 don't continue: advances loop, and this is one missing component - on a migrated user - we want the rest reported on!
                            #CONTINUE #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                        } ; 
                    } else { 
                        $objRet = invoke-XOWrapper {get-xomailbox @pltGxMbx} -credential $pltRXO.Credential -credentialOP $pltRX10.Credential ;
                    } ; 
                    
                    if( ($objRet|Measure-Object).count -AND $objRet.GetType().FullName -eq 'System.Management.Automation.PSObject' ){
                        $smsg = "get-xomailbox:$($tenorg):returned populated ExMbx!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        $xoMbx = $objRet ;

                        $pltGxMStat=[ordered]@{
                            #identity = $mbx.UserPrincipalName ;
                            # they've got 2 david.smith@toro.com's onboarded, both with same UPN, shift to DN it's more specific
                            identity = $xoMbx.UserPrincipalName 
                            ErrorAction='STOP' ;
                            verbose = ($VerbosePreference -eq "Continue") ;
                        } ;
                        $smsg = "Get-xoMailboxStatistics  w`n$(($pltGxMStat|out-string).trim())" ;
                        if($verbose){
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        <# 3:10 PM 3/27/2023 anything using xow *doesn't need pre-rxo2, xow tests and does automatically
                        if ($script:useEXOv2 -OR $useEXOv2) { reconnect-eXO2 @pltRXO }
                        else { reconnect-EXO @pltRXO } ;
                        #>

                        $objRet = $null ;
                        if($UseConnEXO){
                            $error.clear() ;
                            TRY {
                                $objRet = Get-xoMailboxStatistics @pltGxMStat ;  
                            } CATCH {
                                # or just do idiotproof: Write-Warning -Message $_.Exception.Message ;
                                $ErrTrapd=$Error[0] ;
                                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $smsg = $ErrTrapd.Exception.Message ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #-=-record a STATUSWARN=-=-=-=-=-=-=
                                $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                                if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                                if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
                                #-=-=-=-=-=-=-=-=
                                $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                # 8:48 AM 4/11/2023 don't continue: advances loop, and this is one missing component - on a migrated user - we want the rest reported on!
                                #CONTINUE #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                            } ; 
                        } else { 
                            $objRet = invoke-XOWrapper {Get-xoMailboxStatistics @pltGxMStat} -credential $pltRXO.Credential -credentialOP $pltRX10.Credential ;
                        } ; 
                        if( ($objRet|Measure-Object).count -AND $objRet.GetType().FullName -eq 'System.Management.Automation.PSObject' ){
                            $smsg = "Get-xoMailboxStatistics:$($tenorg):returned populated mbxstat" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $xmbxstat = $objRet ;
                        } else {
                            $smsg = "Get-xoMailboxStatistics:$($tenorg):FAILED TO RETURN populated 'System.Management.Automation.PSObject' xmbxstat!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            # 12:32 PM 4/11/2023 rem throw's they're exiting script entirely (vs just echoing an error to console)
                            #THROW $SMSG ;
                            # 8:48 AM 4/11/2023 don't continue: advances loop, and this is one missing component - on a migrated user - we want the rest reported on!
                            #CONTINUE ;
                        } ;

                    } else {
                        # no issue, no split-brain, nothing returned, and no error
                        $smsg = "(no xoMbx found matching:$($mbx.userprincipalname)" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    } ;

                    if($TagYellow -AND ($TagYellow -contains $mbx.UserPrincipalName)){
                        $hSummary.TagYellow = 'Y' ; 
                    } ; 
                    if($TagGreen -AND ($TagGreen -contains $mbx.UserPrincipalName)){
                        $hSummary.TagGreen = 'Y' ; 
                    } ; 
                    if($TagPurple -AND ($TagPurple -contains $mbx.UserPrincipalName)){
                        $hSummary.TagPurple = 'Y' ; 
                    } ; 

                    #OPRecipientType = $null ; 
                    #OPRecipientTypeDetails = $null ; 
                    if($OpRcp){
                        $hSummary.OPRecipientType = $OpRcp.RecipientType ; 
                        $hSummary.OPRecipientTypeDetails = $OpRcp.RecipientTypeDetails ; 
                    } else { 
                        $hSummary.OPRecipientType = "MISSING ONPREM RECIPIENT OBJECT!" ; 
                        $hSummary.OPRecipientTypeDetails = "MISSING ONPREM RECIPIENT OBJECT!" ; 
                    }; 

                    #$hSummary.MbxTotalItemSizeGB = $mbxstat.TotalItemSize ; # dehydraed dbl value, foramt it v
                    if($mbxstat){
                        $hSummary.MbxTotalItemSizeGB = [decimal]("{0:N2}" -f ($mbxstat.TotalItemSize.tostring().split('(')[1].split(' ')[0].replace(',','')/1GB)) ;
                    } elseif($OpRcp.recipienttype -ne 'UserMailbox'){
                        $hSummary.MbxTotalItemSizeGB = "(No Stats returned: No OP Mailbox found)" ; 
                    } else {
                        $hSummary.MbxTotalItemSizeGB = "(No Stats returned: Never logged in?)" ; 
                    } ;
                    if($adu){
                        $hSummary.ADEmployeenumber = $adu.Employeenumber ;
                        $hSummary.ADEnabled = [boolean]($adu.enabled) ;
                        $hSummary.ADCity = $adu.City ;
                        $hSummary.ADCompany = $adu.Company ;
                        $hSummary.ADCountry = $adu.Country ;
                        $hSummary.ADcountryCode = $adu.countryCode ;
                        $hSummary.ADDepartment = $adu.Department ;
                        $hSummary.ADDivision = $adu.Division ;
                        $hSummary.ADemployeeType = $adu.employeeType ;
                        $hSummary.ADGivenName = $adu.GivenName ;
                        $hSummary.ADmailNickname = $adu.mailNickname ;
                        $hSummary.ADMobilePhone = $adu.MobilePhone ;
                        $hSummary.ADOffice = $adu.Office ;
                        $hSummary.ADOfficePhone = $adu.OfficePhone ;
                        $hSummary.ADOrganization = $adu.Organization ;
                        $hSummary.ADphysicalDeliveryOfficeName = $adu.physicalDeliveryOfficeName ;
                        $hSummary.ADPOBox = $adu.POBox ;
                        $hSummary.ADPostalCode = $adu.PostalCode ;
                        $hSummary.ADState = $adu.State ;
                        $hSummary.ADStreetAddress = $adu.StreetAddress ;
                        $hSummary.ADSurname = $adu.Surname ;
                        $hSummary.ADTitle = $adu.Title ;
                        #$prpAxDUserSmtpProxyAddr = @{Name="SmtpProxyAddresses";Expression={ ($_.ProxyAddresses.tolower() |?{$_ -match 'smtp:'}) } } ;
                        $hSummary.ADSMTPProxyAddresses = $adu | select $prpAxDUserSmtpProxyAddr  ;
                        $hSummary.ADMemberof = $adu.memberof ;
                        # for pipeline items, don't process unless there's a value... (err suppress)
                        if($adu.createTimeStamp){
                            $hSummary.ADcreateTimeStamp = (get-date $adu.createTimeStamp -format 'MM/dd/yyyy hh:mm tt');
                        } else {
                            $hSummary.ADcreateTimeStamp = $null ;
                        } ;
                        if($adu.modifyTimeStamp){
                            $hSummary.ADmodifyTimeStamp = (get-date $adu.modifyTimeStamp -format 'MM/dd/yyyy hh:mm tt');
                        } else {
                           $hSummary.ADmodifyTimeStamp = $null ;
                        } ;
                    } else { 
                        $hSummary.ADEmployeenumber = 'MISSING ADUSER!' ;
                        $hSummary.ADEnabled = $FALSE ;
                        $hSummary.ADCity = $null ;
                        $hSummary.ADCompany = $null ;
                        $hSummary.ADCountry = $null ;
                        $hSummary.ADcountryCode = $null ;
                        $hSummary.ADDepartment = $null ;
                        $hSummary.ADDivision = $null ;
                        $hSummary.ADemployeeType = $null ;
                        $hSummary.ADGivenName = $null ;
                        $hSummary.ADmailNickname = $null ;
                        $hSummary.ADMobilePhone = $null ;
                        $hSummary.ADOffice = $null ;
                        $hSummary.ADOfficePhone = $null ;
                        $hSummary.ADOrganization = $null ;
                        $hSummary.ADphysicalDeliveryOfficeName = $null ;
                        $hSummary.ADPOBox = $null ;;
                        $hSummary.ADPostalCode = $null ;
                        $hSummary.ADState = $null ;
                        $hSummary.ADStreetAddress = $null ;
                        $hSummary.ADSurname = $null ;
                        $hSummary.ADTitle = $null ;
                        $hSummary.ADSMTPProxyAddresses = $null ;
                        $hSummary.ADMemberof = $null ;
                        $hSummary.ADcreateTimeStamp = $null ;
                        $hSummary.ADmodifyTimeStamp = $null ;
                    } ; 

                    if($AadUser){
                        $hsummary.AADUDirSyncEnabled = $AADUser.DirSyncEnabled ;
                        $hSummary.AADUSMTPProxyAddresses = $AADUser | select $prpAxDUserSmtpProxyAddr  ;
                        $hSummary.AADUserPrincipalName = $AADUser.UserPrincipalName ;
                    
                        if($AADUser.LastDirSyncTime){
                            $hSummary.AADULastDirSyncTime = (get-date $AADUser.LastDirSyncTime -format 'MM/dd/yyyy hh:mm tt');
                        } else {
                            $hSummary.AADULastDirSyncTime = $null ;
                        } ;
                    } else { 
                        $hsummary.AADUDirSyncEnabled = 'MISSING AADUSER!' ;
                        $hSummary.AADUSMTPProxyAddresses = $null ;
                        $hSummary.AADUserPrincipalName = $null ;
                        $hSummary.AADULastDirSyncTime = $null ;
                    } ; 
                    
                    if($mbx){
                        $hsummary.MbxServer = $mbx.ServerName ;
                        $hsummary.MbxDatabase = $mbx.database ;
                        $hSummary.MbxRetentionPolicy = $mbx.RetentionPolicy ;

                        if($mbx.WhenMailboxCreated){
                            $hSummary.WhenMailboxCreated = (get-date $mbx.WhenMailboxCreated -format 'MM/dd/yyyy hh:mm tt');
                        } else {
                            $hSummary.WhenMailboxCreated = $null ;
                        } ;
                        $hSummary.MbxUseDatabaseQuotaDefaults = $mbx.MbxUseDatabaseQuotaDefaults ;

                        if($mbx.UseDatabaseQuotaDefaults){
                            $hSummary.MbxProhibitSendQuotaGB = $mdbquotas[$mbx.database].ProhibitSendQuotaGB ;
                            $hSummary.MbxProhibitSendReceiveQuotaGB = $mdbquotas[$mbx.database].ProhibitSendReceiveQuotaGB ;
                            $hSummary.MbxIssueWarningQuotaGB = $mdbquotas[$mbx.database].IssueWarningQuotaGB ;
                        } else {
                            $smsg = "(Custom Mbx Quotas configured...)" ;
                            if($verbose){
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            } ; 
                            if($mbx.ProhibitSendQuota -eq 'unlimited'){
                                $hSummary.MbxProhibitSendQuotaGB = $mbx.ProhibitSendQuota ;
                            } else { $hSummary.MbxProhibitSendQuotaGB = $mbx.ProhibitSendQuota | convert-DehydratedBytesToGB }
                            if($mbx.ProhibitSendReceiveQuota  -eq 'unlimited'){
                                $hSummary.MbxProhibitSendReceiveQuotaGB = $mbx.ProhibitSendReceiveQuota ;
                            } else { $hSummary.MbxProhibitSendReceiveQuotaGB = $mbx.ProhibitSendReceiveQuota | convert-DehydratedBytesToGB}
                            if($mbx.IssueWarningQuota -eq 'unlimited'){
                                $hSummary.MbxIssueWarningQuotaGB = $mbx.IssueWarningQuota;
                            } else { $hSummary.MbxIssueWarningQuotaGB = $mbx.IssueWarningQuota | convert-DehydratedBytesToGB }
                        } ;
                    } else { 
                        #'MISSING OP MAILBOX!' ;
                        $hsummary.MbxServer = 'MISSING OP MAILBOX!' ;
                        $hsummary.MbxDatabase = $null ;
                        $hSummary.MbxRetentionPolicy = $null ;
                        $hSummary.WhenMailboxCreated = $null ;
                        $hSummary.MbxUseDatabaseQuotaDefaults = $null ;
                        $hSummary.MbxProhibitSendQuotaGB = $null ;
                        $hSummary.MbxProhibitSendReceiveQuotaGB = $null ;
                        $hSummary.MbxIssueWarningQuotaGB = $null ;
                    } ; 

                    if($mbxstat){
                        if($mbxstat.LastLogonTime){
                            $hSummary.MbxLastLogonTime =  (get-date $mbxstat.LastLogonTime -format 'MM/dd/yyyy hh:mm tt');
                        } else {
                            $hSummary.MbxLastLogonTime = $null ;
                        } ;
                        if($mbxstat.TotalItemSize){
                            $hSummary.MbxTotalItemSizeGB = $mbxstat.TotalItemSize | convert-DehydratedBytesToGB ;
                        } else {
                            $hSummary.MbxTotalItemSizeGB = $null ;
                        } ;
                    } elseif($OpRcp.recipienttype -ne 'UserMailbox'){
                        $hSummary.MbxLastLogonTime = "(No Stats returned: No OP Mailbox found)" ;                         
                        $hSummary.MbxTotalItemSizeGB = "(No Stats returned: No OP Mailbox found)" ; 
                    } else { 
                        $hSummary.MbxLastLogonTime = $null ;
                        $hSummary.MbxTotalItemSizeGB = $null ;
                    } ; 

                    <#if( (-not $ADU) -OR (-not $mbxstat)){
                        # if ADUser is missing, mailbox & xoMbx should be missing as well! or mbxstat is empty, could indicate a repaired splitbrain!
                        write-warning "gotcha! unhandled else" ; 

                    }  ; 
                    #>
                    

                    #if($xoMbx -AND $mbx){
                    if ($xoMbx -AND $mbx -AND $OpRcp.RecipientType -eq 'UserMailbox'){
                        #if( (-not $mbxstat) -AND 
                        $smsg = "$($Mbx.UserPrincipalName) HAS SPLIT-BRAIN!" ; 
                        $smsg +="`nOnPremMbx with ExchangeGuid:$($Mbx.ExchangeGuid)" ; 
                        $smsg +="`n*AND* XOMbx with ExchangeGuid:$($xoMbx.ExchangeGuid)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        $SplitBs ++ ; 
                        $hSummary.XoMbxExchangeGUID = $xoMbx.ExchangeGuid ; 
                        $hSummary.SplitBrain = $true ; 
                        if($Xmbxstat){
                            #$hSummary.add('XoMbxTotalItemSizeGB',[decimal]("{0:N2}" -f ($Xmbxstat.TotalItemSize.tostring().split('(')[1].split(' ')[0].replace(',','')/1GB))) ;
                            $hSummary.XoMbxTotalItemSizeGB = [decimal]("{0:N2}" -f ($Xmbxstat.TotalItemSize.tostring().split('(')[1].split(' ')[0].replace(',','')/1GB)) ;
                        } else {
                            $hSummary.XoMbxTotalItemSizeGB = "(No Stats returned: Never logged in?)" ; 
                        } ;
                        if($Xmbxstat.LastLogonTime){
                            $hSummary.XoMbxLastLogonTime = (get-date $Xmbxstat.LastLogonTime -format 'MM/dd/yyyy hh:mm tt');
                        } else {
                            $hSummary.XoMbxLastLogonTime,$null ;
                        } ;
                        #xoMbxWhenChanged = $XoMbx.WhenChanged ;
                        if($Xombx.WhenChanged){
                            $hSummary.XoMbxWhenChanged = (get-date $Xombx.WhenChanged -format 'MM/dd/yyyy hh:mm tt');
                        } else {
                            $hSummary.XoMbxWhenChanged = $null ;
                        } ;
                        #xoMbxWhenCreated  = $XoMbx.WhenCreated ;
                        if($Xombx.WhenCreated){
                            $hSummary.xoMbxWhenCreated = (get-date $Xombx.WhenCreated -format 'MM/dd/yyyy hh:mm tt');
                        } else {
                            $hSummary.xoMbxWhenCreated = $null ;
                        } ;
                    } elseif($xoMbx -AND $OpRcp.RecipientType -eq 'MailUser' -AND $OpRcp.RecipientTypeDetails -like 'Remote*'){
                        # migrated Rmbx 
                        $smsg = "$($Mbx.UserPrincipalName) was PREVIOUSLY AN ONPREM MAILBOX" ; 
                        $smsg +="`nTHAT HAS SINCE BEEN MIGRATED TO O365!" ; 
                        $smsg +="`nOnPrem RemoteMailbox with ExchangeGuid:`n$($xoMbx.ExchangeGuid)" ; 
                        $smsg +="`n*AND* XOMbx with ExchangeGuid:`n$($RMbx.ExchangeGuid)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    } ; 

                    #$Rpt += [psobject]$hSummary ;
                    # convert the hashtable to object for output to pipeline
                    $Rpt += New-Object PSObject -Property $hSummary ;

                
                $smsg = "$((get-date).ToString('HH:mm:ss')):$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $smsg = "Invalid Object Type: Skipping" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level warn } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            };
        } ;  # loop-E

    } ;  # PROC-E
    END {
        if($Rpt){
            if($outputObject){
                $smsg = "(Returning summary objects to pipeline)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $Rpt | Write-Output ;
            } else {
                $ofile = $logfile.replace('-LOG-BATCH','').replace('-log.txt','.csv') ;
                $smsg = "Exporting summary for $(($Rpt|measure).count) mailboxes to CSV:`n$($ofile)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                TRY {
                    # question? Do we need to expand summarize - probly:
                    # $prpExportCSV
                    $Rpt | select $prpExportCSV | export-csv -NoTypeInformation -path $ofile -ErrorAction STOP;
                    $OutputFiles += @($ofile) ; 
                    $ofile = $logfile.replace('-LOG-BATCH','').replace('-log.txt','.XML') ;
                    $smsg = "Exporting summary for $(($Rpt|measure).count) mailboxes to XML:`n$($ofile)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $Rpt | Export-Clixml -Depth 100 -path $ofile -ErrorAction STOP;
                    $OutputFiles += @($ofile) ; 
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
                    # 8:48 AM 4/11/2023 don't continue: advances loop, and this is one missing component - on a migrated user - we want the rest reported on!
                    #CONTINUE #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                } ;
            }
        } else {
            $smsg = "(empty aggregator, nothing successfully processed)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level warn } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        };
        if($OutputFiles){
            $smsg = "`$OutputFiles:`n$(($OutputFiles|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # emit the outputfiles to the pipeline, for capture and reuse.
            $OutputFiles | write-output ; 
        } else {

        }; 
        $stopResults = Stop-transcript  ;
        $smsg = $stopResults ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    } ;  # END-E
}

#*------^ get-MailboxUseStatus.ps1 ^------