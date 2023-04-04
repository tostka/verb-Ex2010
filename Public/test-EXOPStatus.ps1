#*----------v Function test-EXOPConnection() v----------
function test-EXOPConnection {
    <#
    .SYNOPSIS
    test-EXOPConnection.ps1 - Validate EXOP connection, and that the proper Tenant is connected (as per provided Credential)
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2022-06-24
    FileName    : test-EXOPConnection.ps1
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-EXO
    Tags        : Powershell
    REVISIONS
    *11:44 AM 9/12/2022 init ; port Test-EXO2Connection to EXOP support
    .DESCRIPTION
    test-EXOPConnection.ps1 - Validate EXOP connection, and that the proper Tenant is connected (as per provided Credential)
    .PARAMETER Credential
    Credential to be used for connection
    .PARAMETER MinNoWinRMVersion
    MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '2.0.6')[-MinimumVersion '2.0.6']
    .OUTPUT
    System.Boolean
    .EXAMPLE
    PS> $oRet = test-EXOPConnection -verbose ; 
    PS> if($oRet.Valid){
    PS>     $pssEXOP = $oRet.PsSession ; 
    PS>     write-host 'Validated EXOv2 Connected to Tenant aligned with specified Credential'
    PS> } else { 
    PS>     write-warning 'NO EXO USERMAILBOX TYPE LICENSE!'
    PS> } ; 
    Evaluate EXOP connection status with verbose output
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    #Requires -Version 3
    ##Requires -Modules AzureAD, verb-Text
    ##Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    [CmdletBinding()]
     Param(
        [Parameter(Mandatory=$False,HelpMessage="Credentials [-Credentials [credential object]]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credOpTORSID
        #,
        #[Parameter(HelpMessage = "MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '2.0.6')[-MinimumVersion '2.0.6']")]
        #[version] $MinNoWinRMVersion = '2.0.6'
    )
    BEGIN {
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        
        
        #*------v PSS & GMO VARIS v------
        # get-pssession session varis
        # select key differentiating properties:
        $pssprops = 'Id','ComputerName','ComputerType','State','ConfigurationName','Availability', 
            'Description','Guid','Name','Path','PrivateData','RootModuleModule', 
            @{name='runspace.ConnectionInfo.ConnectionUri';Expression={$_.runspace.ConnectionInfo.ConnectionUri} },  
            @{name='runspace.ConnectionInfo.ComputerName';Expression={$_.runspace.ConnectionInfo.ComputerName} },  
            @{name='runspace.ConnectionInfo.Port';Expression={$_.runspace.ConnectionInfo.Port} },  
            @{name='runspace.ConnectionInfo.AppName';Expression={$_.runspace.ConnectionInfo.AppName} },  
            @{name='runspace.ConnectionInfo.Credentialusername';Expression={$_.runspace.ConnectionInfo.Credential.username} },  
            @{name='runspace.ConnectionInfo.AuthenticationMechanism';Expression={$_.runspace.ConnectionInfo.AuthenticationMechanism } },  
            @{name='runspace.ExpiresOn';Expression={$_.runspace.ExpiresOn} } ; 
        
        if(-not $EXoPConfigurationName){$EXoPConfigurationName = "Microsoft.Exchange" };
        if(-not $rgxEXoPrunspaceConnectionInfoAppName){$rgxEXoPrunspaceConnectionInfoAppName = '^/(exadmin|powershell)$'}; 
        if(-not $EXoPrunspaceConnectionInfoPort){$EXoPrunspaceConnectionInfoPort = '80' } ; 
                
        # gmo varis
        # EXOP
        if(-not $rgxExoPsessionstatemoduleDescription){$rgxExoPsessionstatemoduleDescription = '/(exadmin|powershell)$' };
        if(-not $PSSStateOK){$PSSStateOK = 'Opened' };
        if(-not $PSSAvailabilityOK){$PSSAvailabilityOK = 'Available' };
        if(-not $EXOPGmoFilter){$EXOPGmoFilter = 'tmp_*' } ; 
        if(-not $EXOPGmoTestCmdlet){$EXOPGmoTestCmdlet = 'Add-ADPermission' } ; 
        #*------^ END PSS & GMO VARIS ^------

        # exop is dyn remotemod, not installed
        <#Try { Get-Module $modname -ErrorAction Stop | out-null } Catch {
            $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            Import-Module @pltIMod ;
        } ; # IsImported
        [boolean]$IsNoWinRM = [boolean]([version](get-module $modname).version -ge $MinNoWinRMVersion) ; 
        #>

    } ;  # if-E BEGIN    
    PROCESS {
        $oReturn = [ordered]@{
            PSSession = $null ; 
            Valid = $false ; 
        } ; 
        $isEXOPValid = $false ;
        
        if($pssEXOP = Get-PSSession |?{ (
            $_.runspace.connectioninfo.appname -match $rgxEXoPrunspaceConnectionInfoAppName) -AND (
            $_.runspace.connectioninfo.port -eq $EXoPrunspaceConnectionInfoPort) -AND (
            $_.ConfigurationName -eq $EXoPConfigurationName)}){
                    <# rem'd state/avail tests, run separately below: -AND (
                    $_.State -eq $PSSStateOK)  -AND (
                    $_.Availability -eq $PSSAvailabilityOK)
                    #>
            $smsg = "`n`nEXOP PSSessions:`n$(($pssEXOP | fl $pssprops|out-string).trim())" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

            if($pssEXOPGood = $pssEXOP | ?{ ($_.State -eq $PSSStateOK)  -AND (
                    $_.Availability -eq $PSSAvailabilityOK)}){

                # verify the exop cmdlets actually imported as a tmp_ module w specifid prefix & 1st cmdlet
                # tmp_prpfxxlb.ozy
                if ( (get-module -name $EXOPGmoFilter | ForEach-Object { 
                    Get-Command -module $_.name -name $EXOPGmoTestCmdlet -ea 0 
                })) {

                    $smsg = "(EXOPGmo Basic-Auth PSSession module detected)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $isEXOPValid = $true ; 
                } else { $isEXOPValid = $false ; }
            } else{
                # pss but disconnected state
                rxo2 ; 
            } ; 
            
        } else { 
            $smsg = "Unable to detect EXOP PSSession!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
            else{ write-host -ForegroundColor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            #throw $smsg ;
            #Break ; 
            $isEXOPValid = $false ; 
        } ; 

        if($isEXOPValid){
            $oReturn.PSSession = $pssEXOPGood ;
            $oReturn.Valid = $isEXOPValid ; 
        } else { 
            $smsg = "(invalid session `$isEXOPValid:$($isEXOPValid))" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            Disconnect-ex2010 ;
            $oReturn.PSSession = $pssEXOPGood ; 
            $oReturn.Valid = $isEXOPValid ; 
        } ; 

    }  # PROC-E
    END{
        $smsg = "Returning `$oReturn:`n$(($oReturn|out-string).trim())" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        New-Object PSObject -Property $oReturn | write-output ; 
    } ;
} ; 
#*------^ END Function test-EXOPConnection() ^------