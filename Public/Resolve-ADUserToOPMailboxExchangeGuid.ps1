# Resolve-ADUserToOPMailboxExchangeGuid.ps1

#region RESOLVE_MGUSERTOADUSERHARDMATCH ; #*------v Resolve-ADUserToOPMailboxExchangeGuid v------
Function Resolve-ADUserToOPMailboxExchangeGuid{
    <#
    .SYNOPSIS
    Resolve-ADUserToOPMailboxExchangeGuid - Resolves an ADUser object (or it's msExchMailboxGuid string) to the linked OnPrem Exchange Server Mailbox by specifying the converted ADUser.msExchangeGuid
    .NOTES
    Version     : 0.0.1
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2026-03-31
    FileName    : Resolve-ADUserToOPMailboxExchangeGuid.ps1
    License     : MIT License
    Copyright   : (c) 2026 Todd Kadrie
    Github      : https://github.com/tostka/verb-mg
    Tags        : Powershell,MicrosoftGraph,User,HardMatch,ImmutableID
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 1:15 PM 3/31/2026 init
    .DESCRIPTION
    Resolve-ADUserToOPMailboxExchangeGuid - Resolves an ADUser object (or it's msExchMailboxGuid string) to the linked OnPrem Exchange Server Mailbox by specifying the converted ADUser.msExchangeGuid
    
    Although ADUser.DistinguishedName should represent the same unique object, 
    using the msExchangeGuid -> ExchangeGuid matches the actual low-level linked objects 
    rather than those with the same UPN or other descriptor (where Conflicts may result in multiple recipient objects).

    .PARAMETER InputObject
    ADUser object or ADUser.ExchangeGuid string to be converted to matching Exchange Mailbox[-InputObject `$myADUser]
    .INPUTS
    Accepts piped input
    System.Management.Automation.PSObject (Mailbox Object)
    System.Guid
    System.String
    .OUTPUTS
    Microsoft.ActiveDirectory.Management.ADUser
    .EXAMPLE
    PS> $adu = (get-aduser SAMACCOUNTNAME -property msExchMailboxGuid)
    PS> $mbx = $adu | Resolve-ADUserToOPMailboxExchangeGuid ;     
    Pipeline demo
    .EXAMPLE
    PS> $mgu = Get-MgUser -user TARGETUPN -prop onPremisesImmutableId,userprincipalname,id ; 
    PS> Resolve-ADUserToOPMailboxDistinguishedName -input (get-aduser SAMACCOUNTNAME -property msExchMailboxGuid)
    Commandline demo MGUser input, note, the -property msExchMailboxGuid must be specified, to have get-aduser return the key property
    .LINK
    https://github.com/tostka/verb-MG
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'HIGH')]
    [Alias('ADUserToOPMailbox')]
    PARAM(        
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline = $True,HelpMessage="ADUser object or ADUser.DistinguishedName string to be converted to matching Exchange Mailbox[-InputObject `$myADUser")]
            $InputObject
    ) ;
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
            if($OPMailbox = get-mailbox -identity $inputobject.guid ){
                $OPMailbox| write-output  ; 
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
#endregion RESOLVE_MGUSERTOADUSERHARDMATCH ; #*------^ END Resolve-ADUserToOPMailboxDistinguishedName ^------