#*------v remove-EMSLocalModule.ps1 v------
Function remove-EMSLocalModule {
  <#
    .SYNOPSIS
    remove-EMSLocalModule - remove/unload local server bin-module-based ExchOnPrem Mgmt Shell connection ; validated Exch2016)
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    Version     : 1.0.0
    CreatedDate : 2021-04-15
    FileName    : remove-EMSLocalModule()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell,Exchange,Exchange-2013,Exchange-2016
    REVISIONS   :
    * 10:03 AM 4/16/2021 init vers
    .DESCRIPTION
    remove-EMSLocalModule - remove/unload local server bin-module-based ExchOnPrem Mgmt Shell connection ; validated Exch2016)
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    remove-EMSLocalModule ;
    .LINK
    https://github.com/tostka/verb-Ex2010/
    #>
    [CmdletBinding()]
    #[Alias()]
    Param()  ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ;
    } ;  # BEG-E
    PROCESS{
        $error.clear() ;
        TRY {
            if($tMod = get-module ([System.Net.Dns]::gethostentry($(hostname))).hostname -ea 0){
                write-verbose "(Removing matched EMS module already loaded:)`n$(($tMod | ft -auto ModuleType,Version,Name,ExportedCommands|out-string).trim())" ;
                $tMod | Remove-Module ;
            } else {
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(No matching loaded local Exchange Mgmt Shell binary module found)" ;
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
        # 7:54 AM 11/1/2017 add titlebar tag
        if(gcm Remove-PSTitleBar-PSTitleBar -ea 0 ){Remove-PSTitleBar-PSTitleBar 'EMSL' ;} ;
        # tag E10IsDehydrated
        $Global:ExOPIsDehydrated = $null ;
    } ;  # PROC-E
    END {
        <#
        $tMod = $null ;
        if($tMod = GET-MODULE ([System.Net.Dns]::gethostentry($(hostname))).HOSTNAME){
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(local EMS module loaded:)`n$(($tMod | ft -auto ModuleType,Version,Name,ExportedCommands|out-string).trim())" ;
        } else {
            throw "Unable to resolve target local EMS module:GET-MODULE $([System.Net.Dns]::gethostentry($(hostname))).HOSTNAME)" ;
        } ;
        #>
    }
}

#*------^ remove-EMSLocalModule.ps1 ^------