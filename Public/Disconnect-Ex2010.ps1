#*------v Disconnect-Ex2010.ps1 v------
Function Disconnect-Ex2010 {
  <#
    .SYNOPSIS
    Disconnect-Ex2010 - Clear Remote Exch2010 Mgmt Shell connection
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    AddedCredit : Inspired by concept code by ExactMike Perficient, Global Knowl... (Partner)
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Version     : 1.1.0
    CreatedDate : 2020-02-24
    FileName    : Connect-Ex2010()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell,Exchange,ExchangeOnline
    REVISIONS   :
    * 10:47 AM 7/11/2024 cleared debugging NoSuch etc meta tests
    * 1:34 PM 6/21/2024 ren $Global:E10Sess -> $Global:EXOPSess ; add: prereq checks, and $isBased support, to devert into most connect-exchangeServerTDO, get-ADExchangeServerTDO 100% generic fall back support; sketched in Ex2013 disconnect support
    # 11:12 AM 10/25/2021 added trailing null $Global:E10Sess  (to avoid false conn detects on that test)
    # 9:44 AM 7/27/2021 add -PsTitleBar EMS[ctl] support by dyn gathering range of all 1st & last $Meta.Name[0,2] values
    # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
    * 1:14 PM 3/1/2021 added color reset
    * 4:13 PM 10/22/2020 added pretest of $Global:*'s before running at remove-module (suppresses errors)
    * 12:23 PM 5/27/2020 updated cbh, moved aliases:Disconnect-EMSR','dx10' win func
    * 10:51 AM 2/24/2020 updated attrib
    * 6:59 PM 1/15/2020 cleanup
    * 8:01 AM 11/1/2017 added Remove-PSTitlebar 'EMS', and Disconnect-PssBroken to the bottom - to halt growth of unrepaired broken connections. Updated example to pretest for reqMods
    * 12:54 PM 12/9/2016 cleaned up, add pshelp, implented and debugged as part of verb-Ex2010 set
    * 2:37 PM 12/6/2016 ported to local EMSRemote
    * 2/10/14 posted version
    .DESCRIPTION
    Disconnect-Ex2010 - Clear Remote Exch2010 Mgmt Shell connection
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    $reqMods="Remove-PSTitlebar".split(";") ;
    $reqMods | % {if( !(test-path function:$_ ) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing $($_) function. EXITING." } } ;
    Disconnect-Ex2010 ;
    .LINK
    #>
    [CmdletBinding()]
    [Alias('Disconnect-EMSR','dx10')]
    Param()
    BEGIN{
        
        #region CHKPREREQ ; #*------v CHKPREREQ v------
        # critical dependancy Meta variables
        $MetaNames = ,'TOR','CMW','TOL' #,'NOSUCH' ; 
        # critical dependancy Meta variable properties
        $MetaProps = '' ; #'OP_rgxEMSComputerName','DOESNTEXIST' ; 
        # critical dependancy parameters
        $gvNames = '' ;
        $isBased = $true ; $gvMiss = @() ; $ppMiss = @() ; 
        foreach($met in $metanames){
            write-verbose "chk:`$$($met)Meta" ; 
            if(-not (gv -name "$($met)Meta" -ea 0)){$isBased = $false; $gvMiss += "$($met)Meta" } ; 
            if($MetaProps){
                foreach($mp in $MetaProps){
                    write-verbose "chk:`$$($met)Meta.$($mp)" ; 
                    if(-not (gv -name "$($met)Meta" -ea 0).value[$mp]){$isBased = $false; $ppMiss += "$($met)Meta.$($mp)" } ; 
                } ; 
            } ; 
        } ; 
        if($gvNames){
            foreach($gvN in $gvNames){
                write-verbose "chk:`$$($gvN)" ; 
                if(-not (gv -name "$($gvN)" -ea 0)){$isBased = $false; $gvMiss += "$($gvN)" } ; 
            } ; 
        } ; 
        if($gvmiss){write-warning "Missing Dependant Meta variables:`n$(($gvMiss |%{"`$$($_)" }) -join ',')" } ; 
        if($ppMiss){write-warning "Missing Dependant Meta vari properties:`n$(($ppMiss |%{"`$$($_)" }) -join ',')" } ; 
        if(-not $isBased){ write-warning  "missing critical dependancy profile config!" } ; 
        #endregion CHKPREREQ ; #*------^ END CHKPREREQ ^------

        if($Global:E10Mod){$Global:E10Mod | Remove-Module -Force -verbose:$($false) } ;
        if($Global:EXOPSess){$Global:EXOPSess | Remove-PSSession -verbose:$($false)} ;
        if($isBased){
            $Metas=(get-variable *meta|?{$_.name -match '^\w{3}Meta$'}).name ; 
            # 7:56 AM 11/1/2017 remove titlebar tag
            #Remove-PSTitlebar 'EMS' -verbose:$($VerbosePreference -eq "Continue")  ;
            # 9:21 AM 7/27/2021 expand to cover EMS[tlc]
            #Remove-PSTitlebar 'EMS[ctl]' -verbose:$($VerbosePreference -eq "Continue")  ;
            # make it fully dyn: build range of all meta 1sts & last chars
            [array]$chrs = $metas.substring(0,3).substring(0,1) ; 
            $chrs+= $metas.substring(0,3).substring(2,1) ; 
            $chrs=$chrs.tolower()|select -unique ;
            $sTitleBarTag = "EMS$('[' + (($chrs |%{[regex]::escape($_)}) -join '') + ']')" ; 
            write-verbose "remove PSTitleBarstring:$($sTitleBarTag)" ; 
            Remove-PSTitlebar $sTitleBarTag -verbose:$($VerbosePreference -eq "Continue")  ;
        }  ; 
        # should pull TenOrg if no other mounted 
        <#$sXopDesig = 'xp' ;
        $sXoDesig = 'xo' ;
        #>
        #$xxxMeta.rgxOrgSvcs : $ExchangeServer = (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server|get-random ;
        # normally would be org specific, but we don't have a cred or a TenOrg ref to resolve, so just check xx's version
        # -replace 'EMS','' -replace '\(\|','(' -replace '\|\)',')'
        #if($host.ui.RawUI.WindowTitle -notmatch ((Get-Variable  -name "TorMeta").value.rgxOrgSvcs-replace 'EMS','' -replace '\(\|','(' -replace '\|\)',')' )){
        # drop the current tag being removed from the rgx...
        <# # at this point, if we're no longer using explict Org tag (EMS[tlc] instead), don't need to remove, they'll come out with the EMS removel
        [regex]$rgxsvcs = ('(' + (((Get-Variable  -name "TorMeta").value.OrgSvcs |?{$_ -ne 'EMS'} |%{[regex]::escape($_)}) -join '|') + ')') ;
        if($host.ui.RawUI.WindowTitle -notmatch $rgxsvcs){
            write-verbose "(removing TenOrg reference from PSTitlebar)" ; 
            #Remove-PSTitlebar $TenOrg ;
            # split the rgx into an array of tags
            #sTitleBarTag = (((Get-Variable  -name "TorMeta").value.rgxOrgSvcs) -replace '(\\s\(|\)\\s)','').split('|') ; 
            # no remove all meta tenorg tags - shouldn't be cross-org connecting
            #$Metas=(get-variable *meta|?{$_.name -match '^\w{3}Meta$'}).name ; 
            $sTitleBarTag = $metas.substring(0,3) ; 
            Remove-PSTitlebar $sTitleBarTag -verbose:$($VerbosePreference -eq "Continue") ;
        } else {
            write-verbose "(detected matching OrgSvcs in PSTitlebar: *not* removing TenOrg reference)" ; 
        } ; 
        #>
    }  # BEG-E
    PROCESS{
        # kill any other sessions using distinctive name; add verbose, to ensure they're echo'd that they were missed
        #Get-PSSession | Where-Object { $_.name -eq 'Exchange2010' } | Remove-PSSession -verbose:$($false);
        <# Ex2013 [PS] C:\scripts>get-pssession | fl name,configuration
        Name              : Session1
        ConfigurationName : Microsoft.Exchange
        ComputerName      : server.domain.tld
        #>
        Get-PSSession | Where-Object { $_.ConfigurationName='Microsoft.Exchange'} | Remove-PSSession -verbose:$($false); #version agnostic
        Get-PSSession | Where-Object { $_.name -match 'Exchange2010' } | Remove-PSSession -verbose:$($false); # my older customized connection filtering
        

        # should splice in Ex2013/16 support as well
        # kill any broken PSS, self regen's even for L13 leave the original borked and create a new 'Session for implicit remoting module at C:\Users\', toast them, they don't reopen. Same for Ex2010 REMS, identical new PSS, indistinguishable from the L13 regen, except the random tmp_xxxx.psm1 module name. Toast them, it's just a growing stack of broken's
        Disconnect-PssBroken ;
        #[console]::ResetColor()  # reset console colorscheme
        # null $Global:EXOPSess 
        if($Global:EXOPSess){$Global:EXOPSess = $null } ; 
    } ;  # PROC-E
}

#*------^ Disconnect-Ex2010.ps1 ^------