# start-ExMaintModeExpress.ps1


#region START_EXMAINTMODEEXPRESS ; #*------v start-ExMaintModeExpress v------
function start-ExMaintModeExpress{
            <#
        .SYNOPSIS
        start-ExMaintModeExpress.ps1 - Puts core subset of components into maintenance mode, doesn't really deliver full down, or target CAS processes (use start-Ex16MaintenanceMode)
        .NOTES
        Version     : 0.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2025-03-19
        FileName    : start-ExMaintModeExpress.ps1
        License     : MIT License
        Copyright   : (c) 2025 Todd Kadrie
        Github      : https://github.com/tostka/verb-XXX
        Tags        : Powershell
        AddedCredit : PietroCiaccio
        AddedWebsite: https://github.com/PietroCiaccio/
        AddedTwitter: URL
        REVISIONS
        * 10:45 AM 8/6/2025 added write-myOutput|Warning|Verbose support (for xopBuildLibrary/install-Exchange15.ps1 compat) 
        * 9:29 AM 7/24/2025 copied back from Install-Exchange15-TTC.ps1 (moving to here)
        .DESCRIPTION
        start-ExMaintModeExpress.ps1 - Puts core subset of components into maintenance mode, doesn't really deliver full down, or target CAS processes (use start-Ex16MaintenanceMode)
    
        .PARAMETER Identity ; 
        Specify the identity of the Exchange Server. This can be piped from Get-ExchangeServer or specified explicitly using a string.
        .INPUTS
        None. Does not accepted piped input.(.NET types, can add description)
        .OUTPUTS
        None. Returns no objects or output (.NET types)
        System.Boolean
        [| get-member the output to see what .NET obj TypeName is returned, to use here]
        .EXAMPLE
        PS> .\start-ExMaintModeExpress.ps1 -whatif -verbose
        EXSAMPLEOUTPUT
        Run with whatif & verbose
        .EXAMPLE
        PS> .\start-ExMaintModeExpress.ps1
        EXSAMPLEOUTPUT
        EXDESCRIPTION
        .LINK
        https://github.com/tostka/powershellBB/
        #>
        PARAM(
            # list of servercomponents related to user CAS visibility & mail handling, get them down pdq, immed after install.
            $CompState = @('HubTransport;Draining','FrontendTransport;Draining','ActiveSync;Inactive','Owa;Inactive','UMCallRouter;Inactive','EAS;Inactive','OAB;Inactive')
        ); 
        Import-ExchangeModule ; 
        $smsg = "FORCING INTO EXPRESS MAINTENANE MODE!.."
        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
        <#
        Set-ServerComponentState -Identity "ServerName" -Component HubTransport -State Draining -Requester Maintenance
        Set-ServerComponentState -Identity "ServerName" -Component FrontendTransport -State Draining -Requester Maintenance
        Set-ServerComponentState -Identity "ServerName" -Component ActiveSync -State Inactive -Requester Maintenance
        Set-ServerComponentState -Identity "ServerName" -Component Owa -State Inactive -Requester Maintenance
        Set-ServerComponentState -Identity "ServerName" -Component UMCallRouter -State Inactive -Requester Maintenance
        Set-ServerComponentState -Identity "ServerName" -Component EAS -State Inactive -Requester Maintenance
        Set-ServerComponentState -Identity "ServerName" -Component OAB -State Inactive -Requester Maintenance
        #>
        #$CompState = 'HubTransport;Draining','FrontendTransport;Draining','ActiveSync;Inactive','Owa;Inactive','UMCallRouter;Inactive','EAS;Inactive','OAB;Inactive' ; 
        $Components = $compstate |%{($_.split(';'))[0]} ; 
        [regex]$rgxComponents = ('(' + (($Components |%{[regex]::escape($_)}) -join '|') + ')') ;
        $pltSCS=[ordered]@{
            Identity = $env:COMPUTERNAME ;
            Component = $null ;
            State = $null ;
            Requester = 'Maintenance'  
            ErrorAction = 'STOP' 
        } ; 
        TRY{
            $CS = Get-ServerComponentState -Identity $env:computername -ea STOP ;
        } CATCH {
            $smsg = "Problem running:Get-ServerComponentState -Identity $($env:computername)"
            if(gcm Write-MyError -ea 0){Write-MyError $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error} else{ write-ERROR "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            $res= $false ;
        } ; 
        foreach($Act in $CompState){
            $Component,$State = $Act.split(';') ; 
            $pltSCS.Component = $Component ; 
            $pltSCS.State = $State  ; 
            if($cs.Component -contains $pltSCS.Component){
                $smsg = "Set-ServerComponentState w`n$(($pltSCS|out-string).trim())" ; 
                if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
                TRY{
                    Set-ServerComponentState @pltSCS 
                } CATCH {
                    $smsg = "Problem running:Set-ServerComponentState $($pltSCS.Component), $($pltSCS.State) " 
                    if(gcm Write-MyError -ea 0){Write-MyError $smsg } else {
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error} else{ write-ERROR "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;           
                } ;
            } else { 
                 $smsg = ('ServerComponent {0} is not configured/installed on {1}' -f $pltSCS.Component,$pltSCS.Identity)
                 if(gcm Write-MyVerbose -ea 0){Write-MyVerbose $smsg } else {
                    if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                } ;
            }    
        } ; 
        TRY{
            $CS = Get-ServerComponentState -Identity $env:computername -ea STOP ;    
            $smsg = "POST:get-ServerComponentState w`n$(($CS  | ?{$_.component -match $rgxComponents} | ft -a|out-string).trim())" ; 
            if(gcm Write-MyOutput -ea 0){Write-MyOutput $smsg } else {
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } ;
        } CATCH {
            $smsg = "Problem running:Get-ServerComponentState -Identity $($env:computername)"
            if(gcm Write-MyError -ea 0){Write-MyError $smsg } else {
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error} else{ write-ERROR "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;
            $res= $false ;
        } ; 
        
    }
#endregion START_EXMAINTMODEEXPRESS ; #*------^ END start-ExMaintModeExpress ^------

