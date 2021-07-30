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
    * 12:07 PM 7/30/2021 added CustAttribs to the props dumped ; pulled 'verb-Ex2010' from #requires (nesting limit) ; init
    .DESCRIPTION
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
    https://github.com/tostka/verb-exo
    .LINK
    #>
    ###Requires -Version 5
    #Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS
    #Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.toro\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
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
        $aggreg =@() ; $contacts =@() ;
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
                    } Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
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
                
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
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
            } ; 
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):outputing $(($Report.Enabled|measure).count) Enabled User summaries,`nand $(($Report.Disabled|measure).count) ADUser.Disabled or Disabled/TERM-OU account summaries`nand $(($Report.Contacts|measure).count) users resolved to MailContacts" ; 
        } else { 
            $Report = [ordered]@{
                Enabled = $Aggreg|?{($_.ADEnabled -eq $true) -AND -not($_.distinguishedname -match $rgxDisabledOUs) } ;#?{$_.adDisabled -ne $true -AND -not($_.distinguishedname -match $rgxDisabledOUs)}
                Disabled = $Aggreg|?{($_.ADEnabled -eq $False) -OR ($_.distinguishedname -match $rgxDisabledOUs) } ; 
                Contacts = $contacts ; 
            } ; 
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):outputing $(($Report.Enabled|measure).count) Enabled User summaries,`nand $(($Report.Disabled|measure).count) ADUser.Disabled`nand $(($Report.Contacts|measure).count) users resolved to MailContacts" ; 
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):-ADDisabledOnly specified: 'Disabled' output are *solely* ADUser.Disabled (no  Disabled/TERM-OU account filtering applied)`nand $(($Report.Contacts|measure).count) users resolved to MailContacts" ; 
        } ; 
        New-Object PSObject -Property $Report | write-output ;
        
     } ; 
 } ; 
#*------^ get-UserMailADSummary.ps1 ^------