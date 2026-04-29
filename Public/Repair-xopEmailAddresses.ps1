# Repair-xopEmailAddresses.ps1

#region REPAIR_XOPEMAILADDRESSES ; #*------v Repair-xopEmailAddresses v------
function Repair-xopEmailAddresses{
    <#
    .SYNOPSIS
    Repair-xopEmailAddresses.ps1 - Checks specified mailbox/remotemailbox identifier for malformed email addreses - those with banned characters, and repairs the damaged addresses. 
    .NOTES
    Version     : 0.0.1
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2026-04-24
    FileName    : Repair-xopEmailAddresses.ps1
    License     : MIT License
    Copyright   : (c) 2026 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 1:18 PM 4/27/2026 feature complete finally
    .DESCRIPTION
    Repair-xopEmailAddresses.ps1 - Checks specified mailbox/remotemailbox identifier for malformed email addreses - those with banned characters, and repairs the damaged addresses. 

    Found Exchange Server - where specified Displayname contianed illegal email address characters (ampersand & for example) 
    creating email addresses with the ampersand intact.  
    
    Reportedly Exchange santizes the email address on creation - via EAP or GUI. BS!
    
    Could be a byproduct of manually setting the UPN - possibly a custom UPN forces a matching email address? 
        (no, EAP DOES IT REGARDLESS OF UPN; AND AT THIS POINT UPN IS PRE-CLEANED BY MY CODE)
    Anyway, saw evidence of issues tracking replication to cloud where these banned (never came back as replicated)
    chars are in place, so wrote this to sanitize manually, and ensure we have a 
    valid full set of EmailAddresses on the object after creation.  
    At end this restores EmailAddressPolicyEnabled:$true - not sure if the ampersand will generate another invalid email address as a byproduct. 

    Did full scope test: Restoring EAPEnable to True, COMPLETLEY REVERSES THE FIXES!
    THE EAP ADDED THE BAD ADDRESSES BEING REMOVED. IT'S GOING TO KEEP DOING THAT AS LONG AS ITS IN EFFECT!

    .PARAMETER Identity
    Mailbox identifier[-identity UPN@DOMAIN.COM]
    .PARAMETER RestoreEAPPolicy
    Switch to reset EmailAddressPolicyEnabled to match initial setting (vs disabled used during updates - THIS WILL IMMED RE-ADD THE REMOVED DIRS)[-RestoreEAPPolicy]
    .PARAMETER domaincontroller
    domaincontroller[-domaincontroller 'Dc1']
    .PARAMETER whatIf
    Whatif Flag  [-whatIf]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    System.Management.Automation.PSObject returns updated mailbox object
    .EXAMPLE
    PS> .\Repair-xopEmailAddresses.ps1 -identifier UPN@domain.com -whatif ;
    Run with whatif & verbose
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$true,HelpMessage="Mailbox identifier[-identity UPN@DOMAIN.COM]")]
              [string]$Identity,
        [Parameter(HelpMessage="Switch to reset EmailAddressPolicyEnabled to match initial setting (vs disabled used during updates - THIS WILL IMMED RE-ADD THE REMOVED DIRS)[-RestoreEAPPolicy]")]
            [switch]$RestoreEAPPolicy,
        [Parameter(Mandatory=$true,HelpMessage = "domaincontroller[-domaincontroller 'Dc1']")]
            [string]$domaincontroller,
        [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
            [switch] $whatIf=$true
    ) ; 
    BEGIN{
        $rgxBanEmlDirName = "[\s@!\#\$%\^&\*\(\)\+=\{}\[\]\|\\:;','<>,\?/]" ;
        $xopRcp = get-recipient -identity $identity -domaincontroller $domaincontroller -ea STOP ; 
        switch ($xopRcp.recipienttype){
            'Mailuser'{$MailObject = $xopRcp | get-remotemailbox -domaincontroller $domaincontroller -EA STOP ; $smsgCmd = 'set-remotemailbox '}
            'UserMailbox'{$MailObject = $xopRcp | get-mailbox -domaincontroller $domaincontroller -EA STOP ; $smsgCmd = 'set-mailbox '}
            default{
                $smsg = "Unsupported Recipienttype!:$($xopRcp.recipienttype)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                throw $smsg ; 
                return $false ; 
            }
        } ; 
        $initialEOP = [boolean]$MailObject.EmailAddressPolicyEnabled ; 
        $isUpdated = $false ;
    } ;  # BEG-E
    PROCESS{
        #if($badAddrs = ($MailObject.emailaddresses | ?{$_ -match 'smtp:'} ) -REPLACE 'smtp:','' |?{$_.split('@')[0].replace('smtp:','') -match $rgxbanemldirname}){
        if($badAddrs = ($MailObject.emailaddresses | ?{$_ -match 'smtp:'} ) |?{$_.split('@')[0].replace('smtp:','') -match $rgxbanemldirname}){
            foreach($beml in $badAddrs){
                $sPrefix = $beml.split(':')[0] ; 
                $localpart,$domainname = $beml.split('@') -replace '^smtp:','' ;    
                <#Gemini say:
                Set-Mailbox -Identity "user@domain.com" `
                -EmailAddresses @{Add="newprimary@domain.com"; Remove="oldprimary@domain.com"} `
                -PrimarySmtpAddress newprimary@domain.com
                => use PrimarySmtpAddress as well, to specify in same transaction
                #>
                # 9:31 AM 4/27/2026 better to always prefix; it will auto support ensuring there's a SMTP, if initial is bad and being torn out
                # doesn't work, won't recognize add:SMTP:dirname@domain as an smtp address => type isn't supported by add/remove verb
                #$fixsmtp = "$($sPrefix):$(-join (($localpart -replace $rgxBanEmlDirName,'').ToCharArray()[0..62]))@$($domainname)" ;        
                # back to conditional, instead of spec'ing type, add-PrimarySmtpAddress when SMTP type
                # Nope, stacking=>  EmailAddressees mods with PrimarySmtpAddress: banned: Err:"You can't use the PrimarySmtpAddress and EmailAddresses parameters at the same time."
                # only way, wo completely rewriting the entire EmailAddresses with all types intact, is:
                # 1. add the fix as add & psmtp: Set-Mailbox –Identity candy@contoso.com -WindowsEmailAddress can.dy@contoso.com
                # 2. remove the old psmtp which is now an alias
                $fixsmtp = "$(-join (($localpart -replace $rgxBanEmlDirName,'').ToCharArray()[0..62]))@$($domainname)" ;                
                $pltSMbx = [ordered]@{
                    identity=$MailObject.exchangeguid.guid ; 
                    EmailAddressPolicyEnabled = $false ;         
                    domaincontroller = $domaincontroller ;
                    whatif = $($whatif)
                }; 
                $hsEmailAddresses = @{
                    remove= $beml ; 
                } ;
                # using -WindowsEmailAddress (or -PrimarySMTPAddress, which onlly works OnPrem, wineml works xop & xo)
                # -> if existting, flips it to SMTP; if non-exist, addsd and sets SMTP: no need to pretest                 
                switch -regex -CaseSensitive ($sPrefix){
                    'smtp'{
                        $hsEmailAddresses.add('add',$fixsmtp) ; 
                    }
                    'SMTP'{
                        # we added above, wo type designator, so we have to force it here (or err: no primary specified)
                        $pltSMbx.add('PrimarySmtpAddress',$fixsmtp) ;                               
                    } ; 
                } ; 
                <# prior
                if((($MailObject.emailaddresses | ?{$_ -match 'smtp:'} ) -replace 'smtp:') -contains $fixsmtp){
                }else{
                    #$hsEmailAddresses.add('add',$fixsmtp);
                } ; 
                #>
                TRY{
                    # 1. if configured for a prim assign/set, do it first
                    if($pltSMbx.keys -contains 'PrimarySmtpAddress'){
                        $smsg = "$($smsgCmd) w`n$(($pltSMbx|out-string).trim())" ; 
                        if($pltSMbx.keys -contains 'EmailAddresses'){
                            $smsg += "`n`n-- EmailAddresses:`n$(($pltSMbx.EmailAddresses|out-string).trim())" ; 
                        }
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT} 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        switch ($xopRcp.recipienttype){
                            'Mailuser'{set-remotemailbox @pltSMbx ; $isUpdate = $true }
                            'UserMailbox'{set-mailbox @pltSMbx ; $isUpdate = $true  }
                        } ;
                    } ; 
                    # then 2. run add/remove of alias addresses (can't use remove: on a primary addr, throws 'no primary' error)
                    if($hsEmailAddresses.add -OR $hsEmailAddresses.remove){
                        if($pltSMbx.keys -contains 'PrimarySmtpAddress'){
                            # can't run both, remove prior use conflict
                            $pltSMbx.remove('PrimarySmtpAddress') ; 
                        }
                        $pltSMbx.add('EmailAddresses',$hsEmailAddresses) ; 
                    }
                    $smsg = "$($smsgCmd) w`n$(($pltSMbx|out-string).trim())" ; 
                    if($pltSMbx.keys -contains 'EmailAddresses'){
                        $smsg += "`n`n-- EmailAddresses:`n$(($pltSMbx.EmailAddresses|out-string).trim())" ; 
                    }
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT} 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    switch ($xopRcp.recipienttype){
                        'Mailuser'{set-remotemailbox @pltSMbx ; $isUpdate = $true }
                        'UserMailbox'{set-mailbox @pltSMbx ; $isUpdate = $true  }
                    } ;
                } CATCH {$ErrTrapd=$Error[0] ;
                    write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                } ;        
                
            } ;  # loop-E
            $pltSmbxEnd = @{} ; $pltGmbx= @{} ; 
            $pltsmbx.GetEnumerator() |?{$_.key -match 'identity|domaincontroller|erroraction|whatif'} | foreach-object { $pltSmbxEnd.add($_.name,$_.value)  } ;
            $pltsmbx.GetEnumerator() |?{$_.key -match 'identity|domaincontroller|erroraction'} | foreach-object { $pltGmbx.add($_.name,$_.value)  } ;
            TRY{
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsgCmd) w`n$(($pltSmbxEnd|out-string).trim())" ; 
                switch ($xopRcp.recipienttype){
                    'Mailuser'{$postMailObject = get-remotemailbox @pltGmbx }
                    'UserMailbox'{$postMailObject = get-mailbox @pltGmbx }
                } ;
                $smsg = "POST: $($smsgCmd.replace('set-','get-')) -identity $($pltSmbxEnd.identity)" ; 
                $smsg += ".EmailAddresses smtp:`n$(($postMailObject | select -expand emailaddresses |?{$_ -match 'smtp:'} |out-string).trim())" ;                         
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } CATCH {$ErrTrapd=$Error[0] ;
                write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
            } ;   
            if($RestoreEAPPolicy){
                $smsg = "-RestoreEAPPolicy: restoring initial EmailAddressPolicyEnabled:$($initialEOP)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                $pltSmbxEnd.add('EmailAddressPolicyEnabled',$initialEOP); 
                TRY{
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsgCmd) w`n$(($pltSmbxEnd|out-string).trim())" ; 
                    switch ($xopRcp.recipienttype){
                        'Mailuser'{
                            set-remotemailbox @pltSmbxEnd                         
                            $postMailObject = get-remotemailbox @pltGmbx ;                 
                        }
                        'UserMailbox'{
                            set-mailbox @pltSmbxEnd  ; 
                            $postMailObject = get-mailbox @pltGmbx ; 
                        }
                    } ;
                    $smsg = "POST-RestoreEAPPolicy: $($smsgCmd.replace('set-','get-')) -identity $($pltSmbxEnd.identity)" ; 
                    $smsg += ".EmailAddresses smtp:`n$(($postMailObject | select -expand emailaddresses |?{$_ -match 'smtp:'} |out-string).trim())" ;                         
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } CATCH {$ErrTrapd=$Error[0] ;
                    write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                } ;                    
            }else{
                $smsg = "EmailAddressPolicyEnabled:$($postMailObject) is being left intact, to protect repairs made to assigned EmailAddresses" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            }
        }else{
            $smsg = "($identifier) has no illegal email addresses: no changes)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
        } ; 
    } ;  # PROC-E
    END{        
        switch ($xopRcp.recipienttype){
            'Mailuser'{$MailObject = $xopRcp | get-remotemailbox -domaincontroller $domaincontroller -EA STOP }
            'UserMailbox'{$MailObject = $xopRcp | get-mailbox -domaincontroller $domaincontroller -EA STOP }
        }
        if($MailObject.EmailAddressPolicyEnabled -eq $true -AND $initialEOP){
            $smsg = "-RestoreEAPPolicyEmailAddressPolicyEnabled has been restored to original value:$($initialEOP)" ; 
            $smsg += "`nTHIS WILL HAVE REVERSEd ANY REPAIRS MADE BY THIS FUNCTION" ;
            $smsg += "`nWITHOUT ANY CHANGES TO THE INITIAL PROCESS: IF THE EOP CREATED THE BAD ADDRESS, THE EOP WILL RECREATE THE SAME ADDRESS!" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        }
        $MailObject | write-output ; 
    }
} ; 
#endregion REPAIR_XOPEMAILADDRESSES ; #*------^ END Repair-xopEmailAddresses ^------