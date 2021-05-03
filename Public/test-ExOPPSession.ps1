#*------v test-ExOPPSession.ps1 v------
Function test-ExOPPSession {
  <#
    .SYNOPSIS
    test-ExOPPSession - Does a *simple* - NO-ORG REVIEW - validation of functional PSSession with: ConfigurationName:'Microsoft.Exchange' -AND Name match  '^(Exchange2010|Session\sfor\simplicit\sremoting\smodule\sat\s.*)' -AND State:'Opened' -AND Availability:'Available' -AND can gcm -name 'Add-ADPermission'
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    Version     : 1.0.0
    CreatedDate : 2021-04-15
    FileName    : test-ExOPPSession()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell,Exchange,Exchange-2013,Exchange-2016
    REVISIONS   :
    * 12:30 PM 5/3/2021 init vers ; revised rgxRemsPSSName
    .DESCRIPTION
    test-ExOPPSession - Does a *simple* - NO-ORG REVIEW - validation of functional PSSession with: ConfigurationName:'Microsoft.Exchange' -AND Name match  '^(Exchange2010|Session\sfor\simplicit\sremoting\smodule\sat\s.*)' -AND State:'Opened' -AND Availability:'Available' -AND can gcm -name 'Add-ADPermission'.
    This does *NO* validation that any specific EXOnPrem org is attached! It just validates that an existing PSSession *exists* that *generically* matches a Remote Exchange Mgmt Shell connection in a usable state. Use case is scripts/functions that *assume* you've already pre-established a suitable connection, and just need to pre-test that *any* PSS is already open, before attempting commands. 
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    System.Management.Automation.Runspaces.PSSession. Returns the functional PSSession object(s)
    .EXAMPLE
    PS> if(test-ExOPPSession){'OK'} else { 'NOGO!'}  ;
    .LINK
    https://github.com/tostka/verb-Ex2010/
    #>
    [CmdletBinding()]
    #[Alias()]
    Param()  ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ;
        $rgxRemsPSSName = "^(Exchange\d{4})$" ; 
        $testCommand = 'Add-ADPermission' ; 
        $propsREMS = 'Id','Name','ComputerName','ComputerType','State','ConfigurationName','Availability' ; 
    } ;  # BEG-E
    PROCESS{
        $error.clear() ;
        TRY {
            if($RemsGood = Get-PSSession | where-object { ($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.Name -match $rgxRemsPSSName) -AND ($_.State -eq "Opened") -AND ($_.Availability -eq 'Available') }){
                $smsg = "valid EMS PSSession found:`n$(($RemsGood|ft -a $propsREMS |out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-VERBOSE "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                
                if($tmod = (get-command -name $testCommand).source){
                    $smsg = "(confirmed PSSession open/available, with $($testCommand) available)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $RemsGood | write-output ; ;
                } else { 
                    throw "NO FUNCTIONAL PSSESSION FOUND!" ; 
                } ; 
            } else {
                throw "No existing open/available Remote Exchange Management Shell found!"
            } ;
        } CATCH {
            $ErrTrapd = $_ ;
            write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            #-=-record a STATUSERROR=-=-=-=-=-=-=
            $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(gv passstatus -scope Script -ea 0 ){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script  -ea 0 ){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
            #-=-=-=-=-=-=-=-=
        } ;
        
    } ;  # PROC-E
    END {}
}
#*------^ test-ExOPPSession.ps1 ^------