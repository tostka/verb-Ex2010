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
    * 11:11 AM 8/18/2021 init
    .DESCRIPTION
    resolve-RecipientEAP.ps1 - Resolve an array of recipients against the onprem local EmailAddressPolicies, and return the matching/applicable EAP Policy object
    .PARAMETER  Recipient
    Array of recipient descriptors: displayname, emailaddress, UPN, samaccountname[-recip some.user@domain.com]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    system.systemobject of matching EAP
    .EXAMPLE
    PS> resolve-RecipientEAP -rec todd.kadrie@toro.com -verbose ;
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    ###Requires -Version 5
    #Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Ex2010
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
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $outObject,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $expandAll
    ) ;
    BEGIN{
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        $rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ; 
        $rgxDName = "^([a-zA-Z]{2,}\s[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)" ; 
        $rgxSamAcctNameTOR = "^\w{2,20}$" ; # up to 20k, the limit prior to win2k
        #$rgxSamAcctName = "^[^\/\\\[\]:;|=,+?<>@?]+$" # no char limit ;
        $propsEAPFiltering = 'EmailAddressPolicyEnabled','CustomAttribute5','primarysmtpaddress','Office','distinguishedname' ; 
        
        rx10 -Verbose:$false ; 
        #rxo  -Verbose:$false ; cmsol  -Verbose:$false ;

        # pull EAP's and sub sortable integer values for Priority (Default becomes EAPs.count+1)
        $eaps = Get-EmailAddressPolicy ; 
        $eaps = $eaps | select name,RecipientFilter,RecipientContainer, @{Name="Priority";Expression={ 
            if($_.priority.trim() -match "^[-+]?([0-9]*\.[0-9]+|[0-9]+\.?)$"){
                [int]$_.priority 
            } else { 
                [int]($eaps.count+1) 
            }
            } } | sort priority ; 
    } 
    PROCESS{
       
        $hSum = [ordered]@{
            #dname = $dname;
            #fname = $fname;
            #lname = $lname;
            OPRcp = $OPRcp;
            xoRcp = $xoRcp;
            #OPMailbox = $OPMailbox;
            #OPRemoteMailbox = $OPRemoteMailbox;
            #ADUser = $ADUser;
            #Federator = $null  ;
            #xoMailbox = $xoMailbox;
            #xoUser = $xoUser ;
            #xoMemberOf = $xoMemberOf ;
            #txGuest = $txGuest;
            #MsolUser = $MsolUser ;
            #LicenseGroup = $LicenseGroup ;
        } ;
        #$procd++ ; 
        #write-verbose "processing:$($Recipient)" ; 
            
        #$sBnr="===v ($($Procd)/$($ttl)):Input: '$($Recipient)' v===" ;
        $sBnr="===vInput: '$($Recipient)' v===" ;
        
        $xMProps="samaccountname","windowsemailaddress","DistinguishedName","Office","RecipientTypeDetails" ;
        #$lProps = @{Name='HasLic'; Expression={$_.IsLicensed }},@{Name='LicIssue'; Expression={$_.LicenseReconciliationNeeded }} ;
        #$adprops = "samaccountname","UserPrincipalName","distinguishedname" ; 
            
        #$rgxOPLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ; 
        #$rgxXLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ; 
        #write-host -foreground yellow "get-Rmbx/xMbx: " -nonewline;

        # $isEml=$isDname=$isSamAcct=$false ; 
        $pltgM=[ordered]@{} ; 
        write-verbose "processing:'identity':$($Recipient)" ; 
        $pltgM.add('identity',$Recipient) ;
            
        write-verbose "get-recipient w`n$(($pltgM|out-string).trim())" ; 
        rx10 -Verbose:$false -silent ;
        $error.clear() ;

        $hSum.OPRcp=get-recipient @pltgM -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'}
        #write-verbose "get-exorecipient w`n$(($pltgM|out-string).trim())" ; 
        #rxo  -Verbose:$false -silent ; 
        #$hSum.xoRcp=get-exorecipient @pltgM -ea 0 ;


        if($hSum.OPRcp){

             write-verbose "`$hSum.OPRcp:`n$(($hSum.OPRcp|out-string).trim())" ; 

            $hMsg=@"
Recipient $($hSum.OpRcp.primarysmtpaddress) has the following EmailAddressPolicy-related settings:

$(($hSum.OPRcp | fl $propsEAPFiltering|out-string).trim())

The above settings need to exactly match one or more of the EAP's to generate the desired match...

"@ ;
            
            write-host -foregroundcolor green $hMsg ; 
            
            if($hSum.OPRcp.EmailAddressPolicyEnabled -eq $false){

                write-warning "Recipient $($hSum.OpRcp.primarysmtpaddress) is DISABLED for EAP use:`n$(($hSum.OPRcp | fl EmailAddressPolicy|out-string).trim())`n`nThis user will *NOT* be governed by any EAP until this value is reset to `$true!" ; 
            } else {
                write-verbose  "Recipient $($hSum.OpRcp.primarysmtpaddress) properly has:`n$(($hSum.OPRcp | fl EmailAddressPolicyEnabled|out-string).trim())" ; 
            }  ;

            $bBadRecipientType =$false ;
            switch -regex ($hSum.OPRcp.recipienttype){
                "UserMailbox" {
                    write-verbose "'UserMailbox':get-mailbox $($hSum.OPRcp.identity)"
                    #$hSum.OPMailbox=get-mailbox $hSum.OPRcp.identity ;
                    #write-verbose "`$hSum.OPMailbox:`n$(($hSum.OPMailbox|out-string).trim())" ; 
                } 
                "MailUser" {
                    write-verbose "'MailUser':get-remotemailbox $($hSum.OPRcp.identity)"
                    #$hSum.OPRemoteMailbox=get-remotemailbox $hSum.OPRcp.identity  ;
                    #write-verbose "`$hSum.OPRemoteMailbox:`n$(($hSum.OPRemoteMailbox|out-string).trim())" ; 
                } ;
                default {
                    write-warning "$((get-date).ToString('HH:mm:ss')):Unsupported RecipientType:($hSum.OPRcp.recipienttype). EXITING!" ; 
                    $bBadRecipientType = $true ;
                    Break ; 
                }
            }

            if(!$bBadRecipientType ){
                $error.clear() ;
                TRY {
                   
                    $matchedEAP = $null ; 
                    $propsEAP = 'name','RecipientFilter','RecipientContainer','Priority' ; 
                    $aliasmatch = $hSum.OPRcp.alias ;
                    $aliasClauseMatch = "Alias -ne `$null" ; # alias type original clause substring to find
                    $aliasClauseReplace = "Alias -eq '$($aliasmatch)'" ; # updated recipient-targeting updated clause to use for recipientpreview tests
                    $RecipientTypeMatch = "((RecipientType -eq 'UserMailbox') -or (RecipientType -eq 'MailUser'))" ; # recipienttype original clause substring to find
                    # updated recipient-targeting updated recipienttype clause to use for recipientpreview tests
                    $RecipientTypeReplace = "( Alias -eq '$($aliasmatch)') -AND ((RecipientType -eq 'UserMailbox') -or (RecipientType -eq 'MailUser'))"; 

                    foreach($eap in $eaps){
                        write-verbose "`n==$($eap.name):$($eap.RecipientFilter)" ;
                        # match to clauses, and sub in Alias -eq $Sum.OPRcp.alias to return solely the single user
                        #if($eap.RecipientFilter.indexof("Alias -ne `$null") -ge 0 ){
                        if($eap.RecipientFilter.indexof($aliasClauseMatch) -ge 0 ){

                            #$tmpfilter = $eap.recipientfilter.replace("Alias -ne `$null","Alias -eq 'kadrits'") ; 
                            $tmpfilter = $eap.recipientfilter.replace($aliasClauseMatch,$aliasClauseReplace) ; 

                        #} elseif( $eap.RecipientFilter.indexof("((RecipientType -eq 'UserMailbox') -or (RecipientType -eq 'MailUser'))") -gt 0 ) {
                        } elseif( $eap.RecipientFilter.indexof($RecipientTypeMatch) -gt 0 ) {
                            $tmpfilter = $eap.recipientfilter.replace($RecipientTypeMatch,$RecipientTypeReplace) ;
                        } else {
                            write-warning "Unable to match $($EAP.name) a RecipentFilter clause...`n$($eap.RecipientFilter)`n... to either:`nan $($aliasClauseMatch) clause`nor a $($RecipientTypeMatch ) clause!" ; 
                            $tmpfilter = 'NOMATCH' ; 
                        } ; ; 
                        write-verbose "using `$tmpfilter recipientFilter:`n$($tmpfilter)" ;  
                        $pltGRcpV=[ordered]@{
                            RecipientPreviewFilter=$tmpfilter ;
                            OrganizationalUnit=$eap.RecipientContainer ;
                            resultsize='unlimited';
                            ErrorAction='STOP';
                        } ;
                        write-verbose "get-recipient w`n$(($pltGRcpV|out-string).trim())`n$(($pltGRcpV.RecipientPreviewFilter|out-string).trim())" ; 
                        #if($rcp =get-recipient -RecipientPreviewFilter $tmpfilter -OrganizationalUnit $eap.RecipientContainer -resultsize unlimited |?{$_.alias -eq 'kadrits'} ){
                        if($rcp =get-recipient @pltGRcpV| ?{$_.alias -eq $aliasmatch} ){
                            $matchedEAP = $eap ; 
                            write-verbose "Matched OnPremRecipient $($Sum.OPRcp.alias) to EAP Preview grp:$($rcp.primarysmtpaddress)`nfiltered under EmailAddressPolicy:`n$(($eap | fl $propsEAP |out-string).trim())" ; 
                            break ;
                        } ;
                    };
                
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ; 
                } ; 

            } else { 
                 write-warning "-Recipient:$($Recipient) is of an UNSUPPORTED type by this script! (only Mailbox|MailUser are supported)"   ; 
            } ; 
            
        } else { write-warning "(no matching EXOP recipient object:$($Recipient))"   } ; 
    } #  PROC-E
    END{
         if( $matchedEAP){
             $matchedEAP | write-output ; 
         } else { 
            write-warning "Failed to resolve specified recipient $($user) to a matching EmailAddressPolicy" ; 
            $false | write-output ;
         } ; 
     }
 }

#*------^ resolve-RecipientEAP.ps1 ^------