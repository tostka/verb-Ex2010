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
    CreatedDate : 2020-
    FileName    : get-MailboxUseStatus.ps1
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite:	URL
    AddedTwitter:	URL
    REVISIONS
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
        [Parameter(HelpMessage="Number of levels down the SiteOU name appears in the DistinguishedName (Used to calculate SiteOU: counting from right; defaults to 5)[-SiteOUNestingLevel 3]")]
        [int]$SiteOUNestingLevel=5,
        [Parameter(HelpMessage="Object output switch [-outputObject]")]
        [switch] $outputObject
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
        $prpExportCSV = 'AADUserPrincipalName','DistinguishedName',@{name="AADUAssignedLicenses";expression={($_.AADUAssignedLicenses) -join ";"}},
            'ADEnabled','IsExoLicensed','AADUDirSyncEnabled','AADULastDirSyncTime','MbxLastLogonTime','ADLastLogonTime','ParentOU','samaccountname',
            'SiteOU',@{name="AADUSMTPProxyAddresses";expression={$_.AADUSMTPProxyAddresses.SmtpProxyAddresses -join ";"}},
            'ADCity','ADCompany','ADCountry','ADcountryCode','ADcreateTimeStamp','ADDepartment','ADDivision','ADEmployeenumber','ADemployeeType',
            'ADGivenName','ADmailNickname',@{name="ADMemberof";expression={$_.ADMemberof -join ";"}},'ADMobilePhone','ADmodifyTimeStamp',
            'ADOffice','ADOfficePhone','ADOrganization','ADphysicalDeliveryOfficeName','ADPOBox','ADPostalCode',
            @{name="ADSMTPProxyAddresses";expression={$_.ADSMTPProxyAddresses.SmtpProxyAddresses -join ";"}},'ADState','ADStreetAddress',
            'ADSurname','ADTitle','LicGrouppDN','MbxDatabase','MbxIssueWarningQuotaGB','MbxProhibitSendQuotaGB',
            'MbxProhibitSendReceiveQuotaGB','MbxRetentionPolicy','MbxServer','MbxTotalItemSizeGB','MbxUseDatabaseQuotaDefaults',
            'Name','UserPrincipalName','WhenChanged','WhenCreated','WhenMailboxCreated' ; 
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
        $useEXO = $false ; 
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
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
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
                Silent = $true ;
            } ;
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
            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
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
                Verbose = $FALSE ; Silent = $true ; } ;
            Reconnect-Ex2010 @pltRX10 ; # local org conns
            #$pltRx10 creds & .username can also be used for local ADMS connections
            ###>
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; 
                Silent = $true ; 
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
        #$objRet = xoW {get-ExoMailboxLicenses @pltGXML} -credential $pltRXO.Credential -credentialOP $pltRX10.Credential ; ;
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
        $ttl = ($Mailboxes|measure).count ; $Procd = 0 ;
        foreach ($mbx in $Mailboxes){
            $adu = $mbxstat = $AADUser = $null;
            $isInvalid=$false ;
            switch ($mbx.GetType().fullname){
                'System.String' {
                    # BaseType: System.Object
                    $smsg = "$($mbx) specified does not appear to be a proper Exchange OnPrem Mailbox object"
                    $smsg+= "`ndetected type:`n$(($mbx.GetType() | ft -a fullname,basetype|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $isInvalid=$true;
                    BREAK ;
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
                    BREAK ;
                }
                default {
                    $smsg = "Unrecognized object type! "
                    $smsg+= "`ndetected type:`n$(($mbx.GetType() | ft -a fullname,basetype|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $isInvalid=$true;
                    BREAK ;
                }
            } ;
            $Procd ++ ;

            if(-not $isInvalid){
                $sBnrS="`n#*------v PROCESSING : ($($Procd)/$($ttl)) $($mbx.UserPrincipalName) v------" ;
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
                TRY {
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
                        Name = $mbx.name;
                        ParentOU = (($mbx.distinguishedname.tostring().split(',')) |select -skip 1) -join ',' ;
                        samaccountname = $mbx.samaccountname;
                        SiteOU = ($mbx.distinguishedname.tostring().split(','))[(-1*$SiteOUNestingLevel)..-1] -join ',' ;
                        UserPrincipalName = $mbx.UserPrincipalName;
                        WhenChanged = $mbx.WhenChanged ;
                        WhenCreated  = $mbx.WhenCreated ;
                        WhenMailboxCreated = $null ;
                    } ;
                    <# $prpAadu = 'UserPrincipalName','GivenName','Surname','DisplayName','AccountEnabled','Description','PhysicalDeliveryOfficeName','JobTitle','AssignedLicenses','Department','City','State','Mail','MailNickName','LastDirSyncTime','OtherMails','ProxyAddresses' ;
                    #>
                    $pltGadu=[ordered]@{
                        identity = $mbx.DistinguishedName ;
                        ErrorAction='STOP' ;
                        properties=$prpADU;
                        verbose = ($VerbosePreference -eq "Continue") ;
                    } ;
                    $smsg = "get-aduser w`n$(($pltGadu|out-string).trim())" ;
                    if($verbose){
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $adu = get-aduser @pltGadu
                    #| select $selectADU ;
                    $pltGMStat=[ordered]@{
                        #identity = $mbx.UserPrincipalName ;
                        # they've got 2 david.smith@toro.com's onboarded, both with same UPN, shift to DN it's more specific
                        identity = $mbx.DistinguishedName ;
                        ErrorAction='STOP' ;
                        verbose = ($VerbosePreference -eq "Continue") ;
                    } ;
                    $smsg = "Get-MailboxStatistics  w`n$(($pltGMStat|out-string).trim())" ;
                    if($verbose){
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    Reconnect-Ex2010 @pltRX10  ;  
                    $mbxstat = Get-MailboxStatistics @pltGMStat ;
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
                    $smsg = "Get-AzureADUser on UPN:`n$(($pltGAADU|out-string).trim())" ;
                    $smsg = $recursetag,$smsg -join '' ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    if($AADUser = Get-AzureADUser @pltGAADU){
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

                            $userList = $aadu | Select -ExpandProperty AssignedLicenses | Select SkuID  ;
                            $userLicenses=@() ;
                            $userList | ForEach {
                                $sku=$_.SkuId ;
                                $userLicenses+=$licensePlanListHash[$sku].SkuPartNumber ;
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
                                        break ; # no sense running whole set, break on 1st mbx-support match
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

                    #$hSummary.MbxTotalItemSizeGB = $mbxstat.TotalItemSize ; # dehydraed dbl value, foramt it v
                    if($mbxstat){
                        $hSummary.MbxTotalItemSizeGB = [decimal]("{0:N2}" -f ($mbxstat.TotalItemSize.tostring().split('(')[1].split(' ')[0].replace(',','')/1GB)) ;
                    } else {
                            $hSummary.MbxTotalItemSizeGB = "(No Stats returned: Never logged in?)" ; 
                    } ;
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
                    $hsummary.AADUDirSyncEnabled = $AADUser.DirSyncEnabled ;
                    $hSummary.AADUSMTPProxyAddresses = $AADUser | select $prpAxDUserSmtpProxyAddr  ;
                    $hSummary.AADUserPrincipalName = $AADUser.UserPrincipalName ;

                    $hsummary.MbxServer = $mbx.ServerName ;
                    $hsummary.MbxDatabase = $mbx.database ;
                    $hSummary.MbxRetentionPolicy = $mbx.RetentionPolicy ;

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
                    if($AADUser.LastDirSyncTime){
                        $hSummary.AADULastDirSyncTime = (get-date $AADUser.LastDirSyncTime -format 'MM/dd/yyyy hh:mm tt');
                    } else {
                        $hSummary.AADULastDirSyncTime = $null ;
                    } ;
                    if($mbx.WhenMailboxCreated){
                        $hSummary.WhenMailboxCreated = (get-date $mbx.WhenMailboxCreated -format 'MM/dd/yyyy hh:mm tt');
                    } else {
                        $hSummary.WhenMailboxCreated = $null ;
                    } ;
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

                    #$Rpt += [psobject]$hSummary ;
                    # convert the hashtable to object for output to pipeline
                    $Rpt += New-Object PSObject -Property $hSummary ;
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
                    CONTINUE #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                } ;
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
                    CONTINUE #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
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
        } else {

        }; 
        $stopResults = Stop-transcript  ;
        $smsg = $stopResults ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    } ;  # END-E
}

#*------^ get-MailboxUseStatus.ps1 ^------