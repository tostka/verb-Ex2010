#*------v load-EMSModule.ps1 v------
Function load-EMSModule {
  <#
    .SYNOPSIS
    load-EMSModule - Setup local server bin-module-based ExchOnPrem Mgmt Shell connection (contrasts with Ex2007/10 snapin use ; validated Exch2016)
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    Version     : 1.0.0
    CreatedDate : 2021-04-15
    FileName    : load-EMSModule()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell,Exchange,Exchange-2013,Exchange-2016
    REVISIONS   :
    * 10:14 AM 4/12/2021 init vers
    .DESCRIPTION
    load-EMSModule - Setup local server bin-module-based ExchOnPrem Mgmt Shell connection (contrasts with Ex2007/10 snapin use ; validated Exch2016)
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    load-EMSModule ; 
    .LINK
    https://github.com/tostka/verb-Ex2010/
    #>
    [CmdletBinding()]
    #[Alias()]
    Param()  ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ; 
        <# wraps the native ex server desktop EMS, .lnk config
        . 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell
        #>
    } ;  # BEG-E
    PROCESS{
        $error.clear() ;
        TRY {
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
        } CATCH {
            $ErrTrapd = $_ ; 
            write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            #-=-record a STATUSERROR=-=-=-=-=-=-=
            $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(gv passstatus -scope Script){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
            #-=-=-=-=-=-=-=-=
        } ;
        # 7:54 AM 11/1/2017 add titlebar tag
        Add-PSTitleBar 'EMSL' ;
        # tag E10IsDehydrated 
        $Global:E10IsDehydrated = $false ;        
    } ;  # PROC-E
    END {}
} ; 
#*------^ load-EMSModule.ps1 ^------