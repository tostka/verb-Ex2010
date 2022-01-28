# get-MailboxUseStatus.ps1

#*------v Function get-MailboxUseStatus v------
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
        [Parameter(Position=1,Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of Exchange OnPrem Mailbox Objects[-Mailboxes `$mailboxes]")]
        $Mailboxes,
        [Parameter(Mandatory=$true,HelpMessage="Ticket number[-Ticket 123456]")]
        $Ticket,
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

        if(!(get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
        foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
        if(!(get-variable rgxPSAllUsersScope -ea 0)){
            $rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps(((d|m))*)1|dll)$" ;
        } ;
        if(!(get-variable rgxPSCurrUserScope -ea 0)){
            $rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps((d|m)*)1|dll)$" ;
        } ;
        $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($whatif) ;} ;
        $pltSL.Tag = $Ticket ;
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

        
        $Rpt = @() ; 
        
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
            $adu = $mbxstat = $null ; 
            $Procd ++ ; 
            $sBnrS="`n#*------v PROCESSING : ($($Procd)/$($ttl)) $($mbx.UserPrincipalName) v------" ; 
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnrS)" ;

            rx10 ; 
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
                    MbxLastLogonTime = $null ;
                    MbxTotalItemSizeGB = $null ; 
                    WhenMailboxCreated = $mbx.WhenMailboxCreated ;
                    WhenChanged = $mbx.WhenChanged ;
                    WhenCreated  = $mbx.WhenCreated ;
                    ADEnabled = $null ; 
                    ADEmployeenumber = $null ; 
                    ADcreateTimeStamp = $null ; 
                    ADmodifyTimeStamp = $null ; 
                    ADCity = $null ; 
                    ADCompany = $null ; 
                    ADCountry = $null ; 
                    ADcountryCode = $null ; 
                    ADDepartment = $null ; 
                    ADDivision = $null ; 
                    ADemployeeType = $null ; 
                    ADGivenName = $null ; 
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
                } ; 
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

                if($mbxstat.LastLogonTime){
                    $hSummary.MbxLastLogonTime =  (get-date $mbxstat.LastLogonTime -format 'MM/dd/yyyy hh:mm tt'); 
                } else { 
                    $hSummary.MbxLastLogonTime = $null ; 
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
        } ;  # loop-E

    } ;  # PROC-E
    END {
        
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
        $stopResults = Stop-transcript  ;
        $smsg = $stopResults ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    } ;  # END-E
} ; 
#*------^ END Function get-MailboxUseStatus ^------