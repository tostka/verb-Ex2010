#*------v import-EMSLocalModule v------
Function import-EMSLocalModule {
  <#
    .SYNOPSIS
    import-EMSLocalModule - Setup local server bin-module-based ExchOnPrem Mgmt Shell connection (contrasts with Ex2007/10 snapin use ; validated Exch2016)
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    Version     : 1.0.0
    CreatedDate : 2021-04-15
    FileName    : import-EMSLocalModule()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell,Exchange,Exchange-2013,Exchange-2016
    REVISIONS   :
    * 9:21 AM 4/16/2021 renamed load-emsmodule -> import-EMSLocalModule, added pretest and post verify
    * 10:14 AM 4/12/2021 init vers
    .DESCRIPTION
    import-EMSLocalModule - Setup local server bin-module-based ExchOnPrem Mgmt Shell connection (contrasts with Ex2007/10 snapin use ; validated Exch2016)
    Wraps the native ex server desktop EMS, .lnk commands:
    . 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'; 
    Connect-ExchangeServer -auto -ClientApplication:ManagementShell
    Handy for loading local non-dehydrated support in ISE, regular PS etc, where existing code relies on non-dehydrated objects.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    import-EMSLocalModule ; 
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
                write-verbose "(local EMS module already loaded:)`n$(($tMod | ft -auto ModuleType,Version,Name,ExportedCommands|out-string).trim())" ; 
            } else { 
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(importing local Exchange Mgmt Shell binary module)" ; 
                if($env:ExchangeInstallPath){
                    $rExps1 = "$($env:ExchangeInstallPath)bin\RemoteExchange.ps1"
                    if(test-path $rExps1){
                        . $rExps1 ; 
                        if(gcm Connect-ExchangeServer){
                            Connect-ExchangeServer -auto -ClientApplication:ManagementShell ; 
                        } else { 
                            throw "Unable to gcm Connect-ExchangeServer!" ; 
                        } ; 
                    } else { 
                        throw "Unable to locate: `$(`$env:ExchangeInstallPath)bin\RemoteExchange.ps1" ; 
                    } ; 
                } else { 
                    throw "Unable to locate: `$env:ExchangeInstallPath Environment Variable (Exchange does not appear to be locally installed)" ; 
                } ; 
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
        if(gcm Add-PSTitleBar -ea 0 ){Add-PSTitleBar 'EMSL' ;} ; 
        # tag E10IsDehydrated 
        $Global:ExOPIsDehydrated = $false ;        
    } ;  # PROC-E
    END {
        $tMod = $null ; 
        if($tMod = GET-MODULE ([System.Net.Dns]::gethostentry($(hostname))).HOSTNAME){
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(local EMS module loaded:)`n$(($tMod | ft -auto ModuleType,Version,Name,ExportedCommands|out-string).trim())" ; 
        } else { 
            throw "Unable to resolve target local EMS module:GET-MODULE $([System.Net.Dns]::gethostentry($(hostname))).HOSTNAME)" ; 
        } ; 
    }
} ; 
#*------^ import-EMSLocalModule ^------