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
    FileName    : 
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite:	URL
    AddedTwitter:	URL
    REVISIONS
    * 4:57 PM 2/25/2022 WIP: added services mgmt, & cred handling, pulling AADU licenses etc (need to parse end eval if 'Mailbox'-supporting lic assigned), added local EXOP: quotas, server, db, totoalitemsize for mbx, etc.
    * 1:48 PM 1/28/2022 hit a series of mbxs that were onprem in AM, but migrated in PM; also  they've got 2 david.smith@toro.com's onboarded, both with same UPN, shift gmbxstat to DN it's more specific ; expanded added broad range of ADUser.geopoliticals; added calculated SiteOU as well; working
    .DESCRIPTION
    get-MailboxUseStatus - Analyze and summarize a specified array of Exchange OnPrem mailbox objects to determine 'in-use' status, and export summary statistics to CSV file 
    
    Collects & exports to CSV (or outputs to pipeline, where -outputobject specified), the following information per Mailbox/ADUser
        DistinguishedName
        name        
        MbxLastLogonTime
        MbxTotalItemSizeGB
        ParentOU (calculated from DN)
        SiteOU (calculated from DN)
        samaccountname
        UserPrincipalName
        WhenChanged
        WhenCreated
        WhenMailboxCreated
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
        ADMobilePhone
        ADmodifyTimeStamp
        ADOffice
        ADOfficePhone
        ADOrganization
        ADphysicalDeliveryOfficeName
        ADPOBox
        ADPostalCode
        ADState
        ADStreetAddress
        ADSurname
        ADTitle

    .PARAMETER Mailboxes
    Array of Exchange OnPrem Mailbox Objects[-Mailboxes `$mailboxes]
    .PARAMETER Ticket
    Ticket number[-Ticket 123456]
    .PARAMETER SiteOUNestingLevel
    Number of levels down the SiteOU name appears in the DistinguishedName (Used to calculate SiteOU: counting from right; defaults to 5)[-SiteOUNestingLevel 3]
    .PARAMETER outputObject
    Object output switch [-outputObject]
    .EXAMPLE
    PS> get-MailboxUseStatus -ticket 665437 -mailboxes $NonTermUmbxs -verbose  ; 
    Example processing the specified array, and writing report to CSV, with -verbose output
    .EXAMPLE
    PS> $allExopmbxs | export-clixml .\allExopmbxs-20220128-0945AM.xml ; 
        $allExopmbxs = import-clixml .\allExopmbxs-20220128-0945AM.xml ; 
        $NonTermUmbxs = $allExopmbxs | ?{$_.recipienttypedetails -eq 'UserMailbox' -AND $_.distinguishedname -notmatch ',OU=(Disabled|TERM),' -AND $_.distinguishedname -match ',OU=Users,'} ;
        $Results = get-MailboxUseStatus -ticket 665437 -mailboxes $NonTermUmbxs -outputObject ; 
        $results |?{$_.ADEnabled} |  measure | select -expand count  ; 
    Profile specified list of users (pre-filtered for recipienttypedetails & not stored in term-related OUs), and below Users OUs)
    Then postfilter and count the number actually ADEnabled. 
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    ##Requires -Version 2.0
    #Requires -Version 3
    #requires -PSEdition Desktop
    ##requires -PSEdition Core
    ##Requires -PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
    ##Requires -Modules ActiveDirectory, AzureAD, MSOnline, ExchangeOnlineManagement, MicrosoftTeams, SkypeOnlineConnector, Lync,  verb-AAD, verb-ADMS, verb-Auth, verb-Azure, VERB-CCMS, verb-Desktop, verb-dev, verb-Ex2010, verb-EXO, verb-IO, verb-logging, verb-Mods, verb-Network, verb-L13, verb-SOL, verb-Teams, verb-Text, verb-logging
    ##Requires -Modules ActiveDirectory, AzureAD, MSOnline, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Auth, verb-Ex2010, verb-EXO, verb-IO, verb-logging, verb-Network, verb-Text
    ##Requires -Modules MSOnline, verb-AAD, ActiveDirectory, verb-ADMS, verb-Ex2010, verb-EXO, verb-IO, verb-logging, verb-Network, verb-Text
    #Requires -Modules ActiveDirectory, verb-ADMS, verb-IO, verb-logging, verb-Network, verb-Text
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
        [switch] $LicensedMail,
        [Parameter(HelpMessage="Number of levels down the SiteOU name appears in the DistinguishedName (Used to calculate SiteOU: counting from right; defaults to 5)[-SiteOUNestingLevel 3]")]
        [int]$SiteOUNestingLevel=5,
        [Parameter(HelpMessage="Object output switch [-outputObject]")]
        [switch] $outputObject
    ) # PARAM BLOCK END

    BEGIN { 
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        
        $propsADU = 'employeenumber','createTimeStamp','modifyTimeStamp','City','Company','Country','countryCode','Department','Division','EmployeeNumber','employeeType','GivenName','Office','OfficePhone','Organization','MobilePhone','physicalDeliveryOfficeName','POBox','PostalCode','State','StreetAddress','Surname','Title'  | select -unique ;
        # ,'lastLogonTimestamp' ; worthless, only updated every 9-14d, and then only on local dc - is converting to 1600 as year
        $selectADU = 'DistinguishedName','Enabled','GivenName','Name','ObjectClass','ObjectGUID','SamAccountName','SID',
            'Surname','UserPrincipalName','employeenumber','createTimeStamp','modifyTimeStamp' ;
            #, @{n='LastLogon';e={[DateTime]::FromFileTime($_.LastLogon)}}
        $propsAadu = 'UserPrincipalName','GivenName','Surname','DisplayName','AccountEnabled','Description','PhysicalDeliveryOfficeName','JobTitle','AssignedLicenses','Department','City','State','Mail','MailNickName','LastDirSyncTime','OtherMails','ProxyAddresses' ; 
        # keep the smtp prefix to tell prim/alias addreses
        #$propsAxDUserSmtpProxyAddr = @{Name="SmtpProxyAddresses";Expression={ ($_.ProxyAddresses.tolower() |?{$_ -match 'smtp:'})  -replace ('smtp:','') } } ;
        $propsAxDUserSmtpProxyAddr = @{Name="SmtpProxyAddresses";Expression={ ($_.ProxyAddresses.tolower() |?{$_ -match 'smtp:'}) } } ;

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
        TRY {
            if($logspec){
                $logging=$logspec.logging ;
                $logfile=$logspec.logfile ;
                $transcript=$logspec.transcript ;
                $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                $startResults = start-Transcript -path $transcript ;
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

        $1stConn = $true ; 
        
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
        #region useEXO ; #*------v useEXO v------
        $useEXO = $true ; # non-dyn setting, drives variant EXO reconnect & query code
        #if($CloudFirst){ $useEXO = $true } ; # expl: steering on a parameter
        if($useEXO){
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
            if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -UserRole 'CSVC','SID' -verbose:$($verbose))){
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
            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
            else { reconnect-EXO @pltRXO } ;
            # or with Tenant-specific cred($Tenorg) lookup
            #$pltRXO creds & .username can also be used for AzureAD connections
            Connect-AAD @pltRXO ;
            ###>
            # configure splat for connections: (see above useage)
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; Silent = $false ; } ;
            if($1stConn){
                $pltRXO.silent = $false ; 
            } else { 
                $pltRXO.silent = $true ; 
            } ; 

            #endregion GENERIC_EXO_CREDS_&_SVC_CONN #*------^ END GENERIC EXO CREDS & SVC CONN BP ^------
        } # if-E $useEXO
        #endregion useEXO ; #*------^ END useEXO ^------

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
            $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
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


        #region UseOPAD #*------v UseOPAD v------
        if($UseExOP){
            #region GENERIC_ADMS_CONN_&_XO #*------v GENERIC ADMS CONN & XO  v------
            $smsg = "(loading ADMS...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # always capture load-adms return, it outputs a $true to pipeline on success
            $ADMTLoaded = load-ADMS -Verbose:$FALSE ;
            <#
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
            #endregion GENERIC_ADMS_CONN_&_XO #*------^ END GENERIC ADMS CONN & XO ^------
        } ;
        #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
        #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller=get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
        # use new get-GCFastXO cross-org dc finde
        # default to Op_ExADRoot forest from $TenOrg Meta
        if($UseExOP -AND -not $domaincontroller){
            #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((gv -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};
            # need to debug the above, credential issue?
            # just get it done
            $domaincontroller = get-GCFast
        } ;
        #endregion UseOPAD #*------^ END UseOPAD ^------

        #region MSOL_CONNECTION ; #*------v  MSOL CONNECTION v------
        $reqMods += "connect-msol".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        $smsg = "(loading AAD...)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #connect-msol ;
        connect-msol @pltRXO ;
        #endregion MSOL_CONNECTION ; #*------^  MSOL CONNECTION ^------
        #

        #
        #region AZUREAD_CONNECTION ; #*------v AZUREAD CONNECTION v------
        $reqMods += "Connect-AAD".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        $smsg = "(loading AAD...)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #connect-msol ;
        Connect-AAD @pltRXO ;
        #region AZUREAD_CONNECTION ; #*------^ AZUREAD CONNECTION ^------
        #

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

        $pltGLPList=[ordered]@{
            TenOrg= $TenOrg;
            verbose=$($verbose) ;
            credential= $pltRXO.credential ;
            #(Get-Variable -name cred$($tenorg) ).value ;
        } ;
        $smsg = "$($tenorg):get-AADlicensePlanList w`n$(($pltGLPList|out-string).trim())" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $objRet = $null ;
        $objRet = get-AADlicensePlanList @pltGLPList ;
        switch ($objRet.GetType().FullName){
            "System.Collections.Hashtable" {
                if( ($objRet|Measure-Object).count ){
                    $smsg = "get-AADlicensePlanList:$($tenorg):returned populated LicensePlanList" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $licensePlanListHash = $objRet ;
                } else {
                    $smsg = "get-AADlicensePlanList:$($tenorg):FAILED TO RETURN populated LicensePlanList" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            }
            default {
                $smsg = "get-AADlicensePlanList:$($tenorg):RETURNED UNDEFINED OBJECT TYPE!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Exit ;
            } ;
        } ;  # SWITCH-E

        $smsg = "get-MailboxDatabaseQuotas:Qry onprem org hashtable of mailboxquotas per mailboxdatabase" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $mdbquotas = get-MailboxDatabaseQuotas -TenOrg $TenOrg -verbose:$($VerbosePreference -eq "Continue") ;
        $smsg = "$(($mdbquotas|measure).count) quota summaries returned)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

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
        
        $1stConn = $true ; 
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
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnrS)" ;


                #rx10 ; 
                #$pltConn=[ordered]@{verbose=$false ; silent=$false} ; 
                if($1stConn){
                    $pltRX10.silent = $pltRXO.silent = $false ; 
                } else { 
                    $pltRX10.silent = $pltRXO.silent =$true ; 
                } ; 
                ReConnect-Ex2010 @pltRX10 ;
                #rxo @pltConn  ; 
                if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                else { reconnect-EXO @pltRXO } ;
                #cmsol @pltConn ;
                connect-msol @pltRXO ;
                Connect-AAD @pltRXO ;
                connect-ad -verbose:$false | out-null ; 
                $1stConn = $false ; 


                $error.clear() ;
                TRY {
                    $hSummary=[ordered]@{
                        name = $mbx.name; 
                        UserPrincipalName = $mbx.UserPrincipalName; 
                        DistinguishedName = $mbx.DistinguishedName; 
                        ParentOU = (($mbx.distinguishedname.tostring().split(',')) |select -skip 1) -join ',' ;
                        #SiteOU = ($mbx.distinguishedname.tostring().split(','))[-5,-4,-3,-2,-1] -join ',' ;
                        # ((get-mailbox TARGET).distinguishedname.tostring().split(','))[-5..-1] -join ',' ;
                        SiteOU = ($mbx.distinguishedname.tostring().split(','))[(-1*$SiteOUNestingLevel)..-1] -join ',' ;
                        samaccountname = $mbx.samaccountname;
                        MbxServer = $null ; 
                        MbxDatabase = $null ;
                        MbxProhibitSendQuota = $null ;
                        MbxProhibitSendReceiveQuota = $null ;
                        MbxUseDatabaseQuotaDefaults = $null ;
                        MbxIssueWarningQuota = $null ;                        
                        MbxLastLogonTime = $null ;
                        MbxTotalItemSizeGB = $null ; 
                        MbxRetentionPolicy = $null ;
                        WhenMailboxCreated = $null ;
                        WhenChanged = $mbx.WhenChanged ;
                        WhenCreated  = $mbx.WhenCreated ;
                        ADEnabled = $null ; 
                        ADcreateTimeStamp = $null ; 
                        ADmodifyTimeStamp = $null ; 
                        ADCity = $null ; 
                        ADCompany = $null ; 
                        ADCountry = $null ; 
                        ADcountryCode = $null ; 
                        ADDepartment = $null ; 
                        ADDivision = $null ; 
                        ADEmployeenumber = $null ;                         
                        ADemployeeType = $null ; 
                        ADGivenName = $null ; 
                        ADmailNickname = $null ;  
                        ADMobilePhone = $null ; 
                        ADOffice = $null ; 
                        ADOfficePhone = $null ; 
                        ADOrganization = $null ; 
                        ADphysicalDeliveryOfficeName = $null ; 
                        ADPOBox = $null ; 
                        ADPostalCode = $null ; 
                        ADState = $null ; 
                        ADStreetAddress = $null ; 
                        ADSurname = $null ; 
                        ADTitle = $null ; 
                        ADSMTPProxyAddresses = $null ; 
                        AADUAssignedLicenses = $null ; 
                        AADUDirSyncEnabled = $null ; 
                        AADULastDirSyncTime = $null ; 
                        AADSMTPProxyAddresses = $null ;
                        AADUserPrincipalName = $null ; 
                         
                    } ; 
                    <# $propsAadu = 'UserPrincipalName','GivenName','Surname','DisplayName','AccountEnabled','Description','PhysicalDeliveryOfficeName','JobTitle','AssignedLicenses','Department','City','State','Mail','MailNickName','LastDirSyncTime','OtherMails','ProxyAddresses' ; 
                    #>
                    $pltGadu=[ordered]@{
                        identity = $mbx.DistinguishedName ;
                        ErrorAction='STOP' ;
                        properties=$propsADU;
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
                    $mbxstat = Get-MailboxStatistics @pltGMStat ; 
                    <#if($adu.LastLogon){
                        $hSummary.ADLastLogonTime =  (get-date $adu.LastLogon -format 'MM/dd/yyyy hh:mm tt'); 
                    } else { 
                        $hSummary.ADLastLogonTime = $null ; 
                    } ; 
                    #>

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
                            $smsg = "`n===`n$(($aadu|fl $propsAAdu  | out-string).trim())" ; 
                            # select smtpproxyaddresses out: 
                            $smsg +="`nSMTPProxyAddresses:`n$(($aadu | select $propsAxDUserSmtpProxyAddr | select -expand SMTPProxyAddresses| sort |out-string).trim())" ; 
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
                                # test for presence of a common mailbox-supporting lic, (or (Shared|Room|Equipment)Mailbox recipienttypedetail)

                            } ; 
                        } ; 
                    } else { 
                        $smsg = "=>Get-AzureADUser NOMATCH" ; 
                        $smsg = $recursetag,$smsg -join '' ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 

                    #$hSummary.MbxTotalItemSizeGB = $mbxstat.TotalItemSize ; # dehydraed dbl value, foramt it v
                    $hSummary.MbxTotalItemSizeGB = [decimal]("{0:N2}" -f ($mbxstat.TotalItemSize.tostring().split('(')[1].split(' ')[0].replace(',','')/1GB)) ; 
                    $hSummary.ADEmployeenumber = $adu.Employeenumber ; 
                    $hSummary.ADcreateTimeStamp = $adu.createTimeStamp ; 
                    $hSummary.ADmodifyTimeStamp = $adu.modifyTimeStamp ; 
                    $hSummary.ADEnabled = [boolean]($adu.enabled) ; 
                    $hSummary.ADCity = $adu.City ; 
                    $hSummary.ADCompany = $adu.Company ; 
                    $hSummary.ADCountry = $adu.Country ; 
                    $hSummary.ADcountryCode = $adu.countryCode ; 
                    $hSummary.ADcreateTimeStamp = $adu.createTimeStamp ; 
                    $hSummary.ADDepartment = $adu.Department ; 
                    $hSummary.ADDivision = $adu.Division ; 
                    $hSummary.ADemployeeType = $adu.employeeType ; 
                    $hSummary.ADGivenName = $adu.GivenName ; 
                    $hSummary.ADmailNickname = $adu.mailNickname ;
                    $hSummary.ADMobilePhone = $adu.MobilePhone ; 
                    $hSummary.ADmodifyTimeStamp = $adu.modifyTimeStamp ; 
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
                    #$propsAxDUserSmtpProxyAddr = @{Name="SmtpProxyAddresses";Expression={ ($_.ProxyAddresses.tolower() |?{$_ -match 'smtp:'}) } } ;
                    $hSummary.ADSMTPProxyAddresses = $adu.proxyaddresses | select $propsAxDUserSmtpProxyAddr  ;

                    #$hSummary.AADUAssignedLicenses = $null ; 
                    $hsummary.AADUDirSyncEnabled = = $AADUser.DirSyncEnabled ; 
                    $hSummary.AADULastDirSyncTime = $AADUser.LastDirSyncTime ; 
                    $hSummary.AADSMTPProxyAddresses = $aaduuser.proxyaddresses | select $propsAxDUserSmtpProxyAddr  ;
                    $hSummary.AADUserPrincipalName = $AADUser.UserPrincipalName ; 

                    $hsummary.MbxServer = $mbx.server ;
                    $hsummary.MbxDatabase = $mbx.database ;
                    $hSummary.MbxRetentionPolicy = $mbx.MbxRetentionPolicy ;
                    $hSummary.WhenMailboxCreated = $mbx.WhenMailboxCreated ;
                    # for pipeline items, don't process unless there's a value... (err suppress)
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
                    if($mbx.MbxUseDatabaseQuotaDefaults){
                        $hSummary.MbxProhibitSendQuota = $mdbquotas[$mbx.database].ProhibitSendQuota ;
                        $hSummary.MbxProhibitSendReceiveQuota = $mdbquotas[$mbx.database].ProhibitSendReceiveQuota ;
                        $hSummary.MbxIssueWarningQuota = $mdbquotas[$mbx.database].IssueWarningQuota ;
                    } else {
                        write-verbose "(Custom Mbx Quotas configured...)" ; 
                        $hSummary.MbxProhibitSendQuota = $mbx.MbxProhibitSendQuota ;
                        $hSummary.MbxProhibitSendReceiveQuota = $mbx.MbxProhibitSendReceiveQuota ;
                        $hSummary.MbxIssueWarningQuota = $mbx.MbxIssueWarningQuota ;
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
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
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
                $Rpt | export-csv -NoTypeInformation -path $ofile ; 
            } 
        } else {
            $smsg = "(empty aggregator, nothing successfully processed)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level warn } #Error|Warn|Debug 
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        };
        $stopResults = Stop-transcript  ;
        $smsg = $stopResults ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    } ;  # END-E
}

#*------^ get-MailboxUseStatus.ps1 ^------