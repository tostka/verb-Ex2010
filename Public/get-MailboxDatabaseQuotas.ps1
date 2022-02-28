# get-MailboxDatabaseQuotas.ps1

#*------v Function get-MailboxDatabaseQuotas v------
function get-MailboxDatabaseQuotas {
<#
    .SYNOPSIS
    get-MailboxDatabaseQuotas - Queries all on-prem mailbox databases (get-mailboxdatabase) for default quota settings, and returns an indexed hashtable summarizing the values per database (indexed to each database 'name' value).
    .PARAMETER Mailboxes
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2022-02-25
    FileName    : get-MailboxDatabaseQuotas.ps1
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell
    REVISIONS
    * 4:27 PM 2/25/2022 init vers
    .DESCRIPTION
    get-MailboxDatabaseQuotas - Queries all on-prem mailbox databases (get-mailboxdatabase) for default quota settings, and returns an indexed hashtable summarizing the name and quotas per database (indexed to each database 'name' value).
    .PARAMETER TenOrg
TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']
    .EXAMPLE
    PS> $hQuotas = get-MailboxDatabaseQuotas -verbose ; 
    PS> $hQuotas['database2']
    Name           ProhibitSendReceiveQuotaGB ProhibitSendQuotaGB IssueWarningQuotaGB
    ----           -------------------------- ------------------- -------------------
    database2      12.000                     10.000              9.000
    Retrieve local org on-prem MailboxDatabase quotas and assign to a variable, with verbose outputs. Then output the retrieved quotas from the indexed hash returned, for the mailboxdatabase named 'database2'.
    .EXAMPLE
    PS> $pltGMDQ=[ordered]@{
            TenOrg= $TenOrg;
            verbose=$($VerbosePreference -eq "Continue") ;
            credential= $pltRXO.credential ;
            #(Get-Variable -name cred$($tenorg) ).value ;
        } ;
    PS> $smsg = "$($tenorg):get-MailboxDatabaseQuotas w`n$(($pltGMDQ|out-string).trim())" ;
    PS> if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS> $objRet = $null ;
    PS> $objRet = get-MailboxDatabaseQuotas @pltGMDQ ;
    PS> switch -regex ($objRet.GetType().FullName){
            "(System.Collections.Hashtable|System.Collections.Specialized.OrderedDictionary)" {
                if( ($objRet|Measure-Object).count ){
                    $smsg = "get-MailboxDatabaseQuotas:$($tenorg):returned populated MailboxDatabaseQuotas" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $mdbquotas = $objRet ;
                } else {
                    $smsg = "get-MailboxDatabaseQuotas:$($tenorg):FAILED TO RETURN populated MailboxDatabaseQuotas" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    THROW $SMSG ; 
                    break ; 
                } ;
            }
            default {
                $smsg = "get-MailboxDatabaseQuotas:$($tenorg):RETURNED UNDEFINED OBJECT TYPE!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Exit ;
            } ;
        } ;  
    PS> $smsg = "$(($mdbquotas|measure).count) quota summaries returned)" ;
    PS> if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    # given populuated $mbx 'mailbox object', lookup demo:
    PS> if($mbx.UseDatabaseQuotaDefaults){
            $MbxProhibitSendQuota = $mdbquotas[$mbx.database].ProhibitSendQuota ;
            $MbxProhibitSendReceiveQuota = $mdbquotas[$mbx.database].ProhibitSendReceiveQuota ;
            $MbxIssueWarningQuota = $mdbquotas[$mbx.database].IssueWarningQuota ;
        } else {
            write-verbose "(Custom Mbx Quotas configured...)" ;
            $MbxProhibitSendQuota = $mbx.ProhibitSendQuota ;
            $MbxProhibitSendReceiveQuota = $mbx.ProhibitSendReceiveQuota ;
            $MbxIssueWarningQuota = $mbx.IssueWarningQuota ;
        } ;    
    Expanded example with testing of returned object, and demoes use of the returned hash against a mailbox spec, steering via .UseDatabaseQuotaDefaults
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    #Requires -Modules verb-IO, verb-logging, verb-Text
    [OutputType('System.Collections.Hashtable')]
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        [string]$TenOrg = 'TOR',
        [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credTORSID
    ) ;
    
    ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
    $verbose = ($VerbosePreference -eq "Continue") ;
    
    # select db properties (converts dehydrated bytes string values to decimal gigabytes, via my verb-io module's convert-DehydratedBytesToGB())
    $propsMDB = 'Name',@{Name='ProhibitSendReceiveQuotaGB';Expression={$_.ProhibitSendReceiveQuota | convert-DehydratedBytesToGB }},
    @{Name='ProhibitSendQuotaGB';Expression={$_.ProhibitSendQuota | convert-DehydratedBytesToGB }},
    @{Name='IssueWarningQuotaGB';Expression={$_.IssueWarningQuota | convert-DehydratedBytesToGB }} ; 
    #'ProhibitSendReceiveQuota','ProhibitSendQuota','IssueWarningQuota' ; 
    
    #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======
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
        if($Credential){
            $pltGHOpCred.add('Credential',$Credential) ;
            if(get-Variable -Name "cred$($tenorg)OP" -scope Script -ea 0){
                set-Variable -Name "cred$($tenorg)OP" -scope Script -Value $Credential ;
            } else { New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $Credential } ;
        } else { 
            if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name "cred$($tenorg)OP" -scope Script -ea 0 ){ remove-Variable -Name "cred$($tenorg)OP" -scope Script } ;
                New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ;
                if($verbose){
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                Break ;
            } ;
        } ; 
        $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
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

    # check if using Pipeline input or explicit params:
    if ($PSCmdlet.MyInvocation.ExpectingInput) {
        write-verbose "Data received from pipeline input: '$($InputObject)'" ;
    } else {
        # doesn't actually return an obj in the echo
        #write-verbose "Data received from parameter input: '$($InputObject)'" ;
    } ;
    
    # building a CustObj (actually an indexed hash) with the default quota specs from all db's. The 'index' for each db, is the db's Name (which is also stored as Database on the $mbx)
    if($host.version.major -gt 2){$dbQuotas = [ordered]@{} } 
    else { $dbQuotas = @{} } ;
    
    $smsg = "(querying quotas from all local-org mailboxdatabases)" ; 
    if($verbose){
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    } ; 
    
    $error.clear() ;
    TRY {
        $dbQuotaDefaults=(get-mailboxdatabase -erroraction 'STOP' | sort server,name | select $propsMDB ) ;
    } CATCH {
        $ErrTrapd=$Error[0] ;
        $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #-=-record a STATUSWARN=-=-=-=-=-=-=
        $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
        if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
        if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
        #-=-=-=-=-=-=-=-=
        $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
    } ; 
    
    $ttl = ($dbQuotaDefaults|measure).count ; $Procd = 0 ; 
    foreach ($db in $dbQuotaDefaults){
        $Procd ++ ; 
        $sBnrS="`n#*------v PROCESSING : ($($Procd)/$($ttl)) $($db.name) v------" ; 
        $smsg = $sBnrS ; 
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        
        $name =$($db | select -expand Name) ; 
        $dbQuotas[$name] = $db ; 

        $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
    } ;  # loop-E

    if($dbQuotas){
        $smsg = "(Returning summary objects to pipeline)" ; 
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        $dbQuotas | Write-Output ; 
    } else {
        $smsg = "NO RETURNABLE `$dbQuotas OBJECT!" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        THROW $smsg ;
    } ; 
} ; 
#*------^ END Function get-MailboxDatabaseQuotas ^------