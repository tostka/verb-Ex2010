#*------v resolve-RecipientEAP.ps1 v------
function resolve-RecipientEAP {
    <#
    .SYNOPSIS
    resolve-RecipientEAP.ps1 - Resolve a recipient against the onprem local EmailAddressPolicies, and return the matching/applicable EAP Policy object
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-08-18
    FileName    : resolve-RecipientEAP.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 3:27 PM 8/23/2021 revised patched in new preview-EAPUpdate() support; added 
    default EAP cheatsheet output dump to console; suppress get-EAP warning ; 
    revised recipientfilter support to simple ($(existingRcpFltr) -AND (alias -eq $rcp.alias)).
    Much less complicated, should work on any eap with a recip fltr. 
    * 3:00 PM 8/19/2021 tested, fixed perrotpl issue (overly complicated rcpfltr), 
    pulls a single recipient back on a match on any domain. Considered running a 
    blanket 'get all matches' on each, and then post-filtering for target user(s) 
    but: filtered to a single recip in the rcptfilter, takes 8s for @toro.com; for 
    all targeted is's 1m+. And, just running for broad matches, wo considering 
    priority isn't valid: higher priority matches shut down laters, so you *need* 
    to run them in order, one at a time, and quit on first match. You can't try to 
    datacolect & postfilter wo considering priority, given user may match mult EAPs.
    * 11:11 AM 8/18/2021 init
    .DESCRIPTION
    resolve-RecipientEAP.ps1 - Resolve an array of recipients against the onprem local EmailAddressPolicies, and return the matching/applicable EAP Policy object
    Runs a single recipient (rather than an array) because you really can't pre-collect full populations and stop. Need to run the EAPs in priority order, filter population returned, and quit on first match.
    .PARAMETER  Recipient
    Array of recipient descriptors: displayname, emailaddress, UPN, samaccountname[-recip some.user@domain.com]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER useAltFilter
    Switch to attempt broad append '(existing) -AND (Alias -eq '`$(alias)' to eap.recipientfilter, rather than fancy search/replc on clauses (defaulted TRUE) [-useAltFilter]
    .INPUTS
    None. Does not accepted piped input
    .OUTPUTS
    System.Management.Automation.PSCustomObject of matching EAP
    .EXAMPLE
    PS> $matchedEAP = resolve-RecipientEAP -rec todd.kadrie@toro.com -verbose ;
    PS> if($matchedEAP){"User matches $($matchedEAP.name"} else { "user matches *NO* existing EAP! (re-run with -verbose for further details)" } ; 
    .EXAMPLE
    "user1@domain.com","user2@domain.com"|%{resolve-RecipientEAP -rec $_ -verbose} ; 
    Foreach-object loop an array of descriptors 
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    ###Requires -Version 5
    ###Requires -Modules verb-Ex2010 - disabled, moving into the module
    #Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.toro\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    PARAM(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of recipient descriptors: displayname, emailaddress, UPN, samaccountname[-recip some.user@domain.com]")]
        #[ValidateNotNullOrEmpty()]
        #[Alias('ALIAS1', 'ALIAS2')]
        $Recipient,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2,
        [Parameter(HelpMessage="Display EmailAddressPolicy format strings 'cheatsheet' (defaults true) [-showCheatsheet]")]
        [switch] $showCheatsheet=$true
    ) ;
    BEGIN{
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        $rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ; 
        $rgxDName = "^([a-zA-Z]{2,}\s[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)" ; 
        $rgxSamAcctNameTOR = "^\w{2,20}$" ; # up to 20k, the limit prior to win2k
        #$rgxSamAcctName = "^[^\/\\\[\]:;|=,+?<>@?]+$" # no char limit ;
        $propsEAPFiltering = 'EmailAddressPolicyEnabled','CustomAttribute5','primarysmtpaddress','Office','distinguishedname','Recipienttype','RecipientTypeDetails' ; 
        $rgxEmailDirLegalChars = "[0-9a-zA-Z-._+&']" ; 
        $hCheatSheet = @"

Email Address Policy AddressTemplate format variables:
|Vari |Value
|-----|-------------------------------------|
|%d   |Display name                       
|%g   |Given name                         
|%i   |Middle initial                      
|%m   |Exchange alias                      
|%rxy |Replace all occurrences of x with y 
|%rxx |Remove all occurrences of x         
|%s   |Surname                 
|%ng  |The first n letters of the givenname.
|%ns  |The first n letters of the surname. 

All smtpaddr-illegal chars are dropped from source string. 
Commonly-permitted SmtpAddrChars:
$($rgxEmailDirLegalChars)
(RFC 5322 technically permits broader set, but frequently blocked as risks)

"@ ; 
        
        rx10 -Verbose:$false ; 
        #rxo  -Verbose:$false ; cmsol  -Verbose:$false ;

        # move the properties out to a separate vari
        [array]$eapprops = 'name','RecipientFilter','RecipientContainer','EnabledPrimarySMTPAddressTemplate','EnabledEmailAddressTemplates',
            'DisabledEmailAddressTemplates','Enabled' ; 
        # append an expression that if/then's Priority text value: coercing IsNumeric()'s to [int], else - only non-Numeric is 'Default' - replacing that Priority with [int](EAPs.count)+1
        $eapprops += @{Name="Priority";Expression={ 
            if($_.priority.trim() -match "^[-+]?([0-9]*\.[0-9]+|[0-9]+\.?)$"){
                [int]$_.priority 
            } else { 
                [int]($eaps.count+1) 
            }
            } } ; 
       
        # pull EAP's and sub sortable integer values for Priority (Default becomes EAPs.count+1)
        $smsg = "(polling:Get-EmailAddressPolicy...)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        $sw = [Diagnostics.Stopwatch]::StartNew();
        # use -warningaction silentlycontinue to suppress the 'WARNING: Recipient policy objects that don't contain e-mail address won't be shown unless you include the IncludeMailboxSettingOnlyPolicy'
        $eaps = Get-EmailAddressPolicy -WarningAction 0 ;
        $sw.Stop() ;
        $eaps = $eaps | select $eapprops | sort Priority  ; 
        
        $smsg = ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                
        
    } 
    PROCESS{
       
        $hSum = [ordered]@{
            OPRcp = $OPRcp;
            xoRcp = $xoRcp;
        } ;
                    
        $sBnr="===vInput: '$($Recipient)' v===" ;
        $smsg = $sBnr ;        
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        $xMProps="samaccountname","windowsemailaddress","DistinguishedName","Office","RecipientTypeDetails" ;
        
        $pltgM=[ordered]@{} ; 
        $smsg = "processing:'identity':$($Recipient)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

        $pltgM.add('identity',$Recipient) ;
            
        $smsg = "get-recipient w`n$(($pltgM|out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

        rx10 -Verbose:$false -silent ;
        $error.clear() ;

        $sw = [Diagnostics.Stopwatch]::StartNew();
        $hSum.OPRcp=get-recipient @pltgM -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'}
        $sw.Stop() ;
        $smsg = ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 


        if($hSum.OPRcp){
            $smsg = "`Matched On-Premesis Recipient:`n$(($hSum.OPRcp|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

            $hMsg=@"
Recipient $($hSum.OpRcp.primarysmtpaddress) has the following EmailAddressPolicy-related settings:

$(($hSum.OPRcp | fl $propsEAPFiltering|out-string).trim())

The above settings need to exactly match one or more of the EAP's to generate the desired match...

"@ ;
            
            $smsg = $hMsg ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            if($hSum.OPRcp.EmailAddressPolicyEnabled -eq $false){
                $smsg = "Recipient $($hSum.OpRcp.primarysmtpaddress) is DISABLED for EAP use:`n" ; 
                $smsg += "$(($hSum.OPRcp | fl EmailAddressPolicy|out-string).trim())`n`n" ; 
                $smsg += "This user will *NOT* be governed by any EAP until this value is reset to `$true!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $smsg = "Recipient $($hSum.OpRcp.primarysmtpaddress) properly has:`n$(($hSum.OPRcp | fl EmailAddressPolicyEnabled|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            }  ;

            $bBadRecipientType =$false ;
            switch -regex ($hSum.OPRcp.recipienttype){
                "UserMailbox" {
                    $smsg = "'UserMailbox'"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } 
                "MailUser" {
                    $smsg = "'MailUser'" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ;
                default {
                    $smsg = "Unsupported RecipientType:($hSum.OPRcp.recipienttype). EXITING!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $bBadRecipientType = $true ;
                    Break ; 
                }
            }

            if(!$bBadRecipientType ){
                $error.clear() ;
                TRY {
                   
                    $matchedEAP = $null ; 
                    $propsEAP = 'name','RecipientFilter','RecipientContainer','Priority','EnabledPrimarySMTPAddressTemplate',
                        'EnabledEmailAddressTemplates','DisabledEmailAddressTemplates','Enabled' ; 
                    $aliasmatch = $hSum.OPRcp.alias ;

                    write-host "`n(Comparing to $(($Eaps|measure).count) EmailAddressPolicies for filter-match...)" ;
                    foreach($eap in $eaps){
                        if(!$verbose){
                            write-host "." -NoNewLine ;
                        } ; 
                        $smsg = "`n`n==$($eap.name):$($eap.RecipientFilter)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        
                        # try a simple (existing) -AND "Alias -eq '$($aliasmatch)'" filter mod
                        $tmpfilter = "($($eap.recipientfilter)) -and (Alias -eq '$($aliasmatch)')" ; 
                        
                        $smsg = "using `$tmpfilter recipientFilter:`n$($tmpfilter)" ;  
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        $pltGRcpV=[ordered]@{
                            RecipientPreviewFilter=$tmpfilter ;
                            OrganizationalUnit=$eap.RecipientContainer ;
                            resultsize='unlimited';
                            ErrorAction='STOP';
                        } ;
                        $smsg = "get-recipient w`n$(($pltGRcpV|out-string).trim())`n$(($pltGRcpV.RecipientPreviewFilter|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        
                        $sw = [Diagnostics.Stopwatch]::StartNew();
                        if($rcp =get-recipient @pltGRcpV| ?{$_.alias -eq $aliasmatch} ){
                            $sw.Stop() ;
                            write-host "MATCHED!:$($Eap.name)`n" ;
                            $matchedEAP = $eap ;
                            $smsg = ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
                            $smsg = "Matched OnPremRecipient $($Sum.OPRcp.alias) to EAP Preview grp:$($rcp.primarysmtpaddress)`n" ; 
                            $smsg += "filtered under EmailAddressPolicy:`n$(($eap | fl ($propsEAP |?{$_ -ne 'EnabledEmailAddressTemplates'}) |out-string).trim())`n" ; 
                            $smsg += "EnabledEmailAddressTemplates:`n$(($eap | select -expand EnabledEmailAddressTemplates |out-string).trim())`n" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            $genEml = preview-EAPUpdate  -eap $matchedEAP -Recipient $hSum.OPRcp -Verbose:($VerbosePreference -eq 'Continue')
                            if($geneml -ne $hSum.OPRcp.PrimarySmtpAddress){
                                $smsg = "`n===Specified recip's PrimarySmtpAddress ($hSum.OPRcp.PrimarySmtpAddress))`n"
                                $smsg += "does *not* appear to match specified template!`n" ; 
                                $smsg += "*manualy review* the template specs *validate*`n"
                                $smsg += "that the desired scheme is being applied!`n==="
                                #write-warning $smsg ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            }else {
                                $smsg = "`n===PrimarysmtpAddr $($hSum.OPRcp.PrimarySmtpAddress))`n"
                                $smsg += "roughly conforms to specified template primary addr`n" ;
                                $smsg += "$($matchedEAP.EnabledPrimarySMTPAddressTemplate)...===`n" ;
                                #write-host -foregroundcolor Green $smsg ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            } ;

                            break ;
                        } else {
                            $sw.Stop() ;
                            $smsg = ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        } ;
                    }; # E-loop
                    
                    if($showCheatsheet){
                        write-host $hCheatSheet
                    } ; 

                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ; 
                } ; 

            } else { 
                 $smsg = "-Recipient:$($Recipient) is of an UNSUPPORTED type by this script! (only Mailbox|MailUser are supported)"   ; 
                 if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
            
        } else { 
            $smsg = "(no matching EXOP recipient object:$($Recipient))"   
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
    } #  PROC-E
    END{
        if( $matchedEAP){
            $matchedEAP | write-output ; 
        } else { 
            $smsg = "Failed to resolve specified recipient $($user) to a matching EmailAddressPolicy" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $false | write-output ;
        } ; 
        $smsg = "$($sBnr.replace('=v','=^').replace('v=','^='))" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
     }

     #*======^ END SUB MAIN ^======
 } ;
#*------^ resolve-RecipientEAP.ps1 ^------