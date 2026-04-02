# Resolve-ADUserToLocalMailObjectExchangeGuid.ps1

#region RESOLVE_ADUSERTOREMOTEMAILBOXEXCHANGEGUID ; #*------v Resolve-ADUserToLocalMailObjectExchangeGuid v------
Function Resolve-ADUserToLocalMailObjectExchangeGuid{
    <#
    .SYNOPSIS
    Resolve-ADUserToLocalMailObjectExchangeGuid - Resolves an ADUser object (or it's msExchMailboxGuid string) to the linked OnPrem Hybrid RemoteMailbox by specifying the converted ADUser.msExchangeGuid
    .NOTES
    Version     : 0.0.1
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2026-03-31
    FileName    : Resolve-ADUserToLocalMailObjectExchangeGuid.ps1
    License     : MIT License
    Copyright   : (c) 2026 Todd Kadrie
    Github      : https://github.com/tostka/verb-mg
    Tags        : Powershell,MicrosoftGraph,User,HardMatch,ImmutableID
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 2:15 PM 4/1/2026 ren Resolve-ADUserToRemoteMailboxExchangeGuid -> Resolve-ADUserToLocalMailObjectExchangeGuid
    * 2:29 PM 3/31/2026 init
    .DESCRIPTION
    Resolve-ADUserToLocalMailObjectExchangeGuid - Resolves an ADUser object (or it's msExchMailboxGuid string) to the linked OnPrem Hybrid RemoteMailbox by specifying the converted ADUser.msExchangeGuid
    
    Although ADUser.DistinguishedName should represent the same unique object, 
    using the msExchangeGuid -> ExchangeGuid matches the actual low-level linked objects 
    rather than those with the same UPN or other descriptor (where Conflicts may result in multiple recipient objects).

    .PARAMETER InputObject
    ADUser object or ADUser.ExchangeGuid string to be converted to matching Exchange Mailbox[-InputObject `$myADUser]
    .PARAMETER ResolveChain
    Switch to resolve any chain of objects ADUser -> RemoteMailbox -> EXOMailbox. Otherwise it solely returns the local resolved Mail recipient object mapped to the ADUser.msExchangeGuid
    .INPUTS
    Accepts piped input
    System.Management.Automation.PSObject (Mailbox Object)
    System.Guid
    System.String
    .OUTPUTS
    Microsoft.ActiveDirectory.Management.ADUser
    .EXAMPLE
    PS> $adu = (get-aduser SAMACCOUNTNAME -property msExchMailboxGuid)
    PS> $mbx = $adu | Resolve-ADUserToLocalMailObjectExchangeGuid ;     
    Pipeline demo
    .EXAMPLE
    PS> $mgu = Get-MgUser -user TARGETUPN -prop onPremisesImmutableId,userprincipalname,id ; 
    PS> Resolve-ADUserToOPMailboxDistinguishedName -input (get-aduser SAMACCOUNTNAME -property msExchMailboxGuid)
    Commandline demo MGUser input, note, the -property msExchMailboxGuid must be specified, to have get-aduser return the key property
    .LINK
    https://github.com/tostka/verb-MG
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'HIGH')]
    [Alias('ADUserToOPMailObject')]
    PARAM(        
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline = $True,HelpMessage="ADUser object or ADUser.DistinguishedName string to be converted to matching Exchange Mailbox[-InputObject `$myADUser")]
            $InputObject,
        [Parameter(HelpMessage="Switch to resolve any chain of objects ADUser -> RemoteMailbox -> EXOMailbox. Otherwise it solely returns the local resolved Mail recipient object mapped to the ADUser.msExchangeGuid")]
            [switch]$ResolveChain
    ) ;
    #region MODULES_FORCE_LOAD ; #*------v MODULES_FORCE_LOAD v------
    # core modes in dep order
    $tmods = @('verb-IO','verb-Text','verb-logging','verb-Desktop','verb-dev','verb-Mods','verb-Network','verb-Auth','verb-ADMS','VERB-ex2010','verb-EXO','VERB-mg') ; 
    # task mods in dep order
    $tmods += @('ExchangeOnlineManagement','ActiveDirectory','Microsoft.Graph.Users')
    $oWPref = $WarningPreference ; $WarningPreference = 'SilentlyContinue' ; 
    $tmods | %{ $thismod = $_ ; TRY{$thismod | ipmo -fo  -ea STOP }CATCH{write-host -foregroundcolor yellow "Missing module:$($thismod)`ntrying find-module lookup..." ; find-module $thismod}} ; $WarningPreference = $oWPref ; 
    #endregion MODULES_FORCE_LOAD ; #*------^ END MODULES_FORCE_LOAD ^------
    
    TRY{
        switch -regex ($inputobject.gettype().fullname){
            'Microsoft\.ActiveDirectory\.Management\.ADUser|System\.Management\.Automation\.PSObject|System\.Collections\.Hashtable|System\.Management\.Automation\.PSCustomObject'{                
                if($inputobject.msExchMailboxGuid.GetType().fullname -eq 'System.Byte[]'){
                    $inputobject = [guid]$inputobject.msExchMailboxGuid
                }else{
                    $smsg = "unrecognized/unpopulated msExchMailboxGuid: $($inputobject.msExchMailboxGuid)" ; 
                    $smsg += "`n(non-mail enabled ADUser?)" ;
                    write-warning $smsg ;
                } ; 
            }
            'System\.Guid'{
                if($inputobject.guid){}
            }
            'System\.Byte\[]'{
                $inputobject = [system.guid]::new($inputobject)
            }
            'System\.String'{
                if($inputobject = [guid]$inputobject){}
            }
            default{
                $smsg = "UNRECOGNIZED -inputobject type:$($inputobject.gettype().fullname)" ; 
                $smsg += "`nPlease specify an ADUser object, or a Guid value" ;
                write-warning $smsg ;
                throw $smsg ; 
            }
        }
        #if($inputobject.gettype().fullname -eq 'System.String'){
        if($inputobject.gettype().fullname -eq 'System.Guid'){
            $MailChain = @() ; 
            if($OPMailObj = get-remotemailbox -identity $inputobject.guid -erroraction silentlycontinue){
            }elseif($OPMailObj = get-mailbox -identity $inputobject.guid  -erroraction silentlycontinue){
            }
            if($OPMailObj){
                $smsg = "ADUser.msExchMailboxGuid matched:`n$($OPMailObj.recipienttype)\$($OPMailObj.recipienttypedetails)"
                write-host $smsg ; 
                if(-not $ResolveChain){
                    $OPMailObj| write-output  ; 
                }else{
                    $MailChain += @($OPMailObj) ; 
                    if($OPMailObj.recipienttype -eq 'MailUser'){
                        write-host -foregroundcolor gray "Checking for matching XOMailbox..." ; 
                        if($xoMailbox = get-xomailbox -identity $inputobject.guid  -erroraction silentlycontinue){
                            $smsg = "... which is ExchangeGuid Linked to`n$(($xoMailbox|out-string).trim())" ; 
                            write-host $smsg ; 
                            $mailChain += @($xoMailbox) ;
                        } 
                    }else{
                        write-host -foregroundcolor yellow "$($OPMailObj.recipienttype) is a terminal object: No remote connected object is discoverable" ; 
                    }
                    $MailChain | write-output ; 
                } ; 
            }else{
                $false | write-output 
            } ; 
        } else { $false | write-output }        
    }CATCH {        
        $ErrTrapd=$Error[0] ;
        write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
        write-warning "$($smsg)" ;
    }
        
} ;
#endregion RESOLVE_ADUSERTOREMOTEMAILBOXEXCHANGEGUID ; #*------^ END Resolve-ADUserToOPMailboxDistinguishedName ^------