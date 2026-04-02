# Resolve-OPMailObjectToADUsermsExchMailboxGuid.ps1

#region RESOLVE_MGUSERTOADUSERHARDMATCH ; #*------v Resolve-OPMailObjectToADUsermsExchMailboxGuid v------
Function Resolve-OPMailObjectToADUsermsExchMailboxGuid{
    <#
    .SYNOPSIS
    Resolve-OPMailObjectToADUsermsExchMailboxGuid - Resolves an OnPrem Exchange Server Mailbox object (or it's ExchangeGuid guid) to the linked ADUser by filtering on msExchMailboxGuid match.
    .NOTES
    Version     : 0.0.1
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2026-03-31
    FileName    : Resolve-OPMailObjectToADUsermsExchMailboxGuid.ps1
    License     : MIT License
    Copyright   : (c) 2026 Todd Kadrie
    Github      : https://github.com/tostka/verb-mg
    Tags        : Powershell,MicrosoftGraph,User,HardMatch,ImmutableID
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 2:28 PM 4/1/2026 ren Resolve-OPMailboxToADUsermsExchMailboxGuid -> Resolve-OPMailObjectToADUsermsExchMailboxGuid
    * 1:15 PM 3/31/2026 init
    .DESCRIPTION
    Resolve-OPMailObjectToADUsermsExchMailboxGuid - Resolves an OnPrem Exchange Server Mailbox object (or it's ExchangeGuid guid) to the linked ADUser by filtering on msExchMailboxGuid match.

    Represents the actual low-level linked objects, rather than those with the same UPN or other descriptor (where Hybrid Conflicts may result in multiple splitbrain mailboxes respectively on ADUser and MGUser).

    .PARAMETER InputObject
    Exchange Mailbox object (or it's ExchangeGuid guid string) to be resolved to matching ADUser object[-InputObject `$myExoMailbox]
    .INPUTS
    Accepts piped input
    System.Management.Automation.PSObject (Mailbox Object)
    System.Guid
    System.String
    .OUTPUTS
    Microsoft.ActiveDirectory.Management.ADUser
    .EXAMPLE
    PS> $xoMbx = get-xomailbox TARGETUPN ; 
    PS> $mgu = $XOmBX | Resolve-OPMailObjectToADUsermsExchMailboxGuid ;     
    Pipeline demo
    .EXAMPLE
    PS> $mgu = Get-MgUser -user TARGETUPN -prop onPremisesImmutableId,userprincipalname,id ; 
    PS> Resolve-OPMailObjectToADUsermsExchMailboxGuid -inputobject $mgu; 
    Commandline demo MGUser input    
    .EXAMPLE
    PS> get-remotemailbox -id aaaaaaa | Resolve-OPMailObjectToADUsermsExchMailboxGuid

        DistinguishedName : AA=Aaaa Aaaaaa,AA=AA,AA=Aaaaa,AA=AAA,AA=aaaaaa,AA=aa,AA=aaaa,AA=aaa
        Enabled           : True
        GivenName         : Aaaa
        msExchMailboxGuid : {192, 129, 180, 240...}
        Name              : Aaaa Aaaaaa
        ObjectClass       : user
        ObjectGUID        : 9aaa9a99-9999-9999-aaaa-a9999aa999aa
        SamAccountName    : aaaaaaa
        SID               : A-9-9-99-9999999999-999999999-9999999999-99999
        Surname           : AaaaaaAaaaaa
        UserPrincipalName : Aaaa.Aaaaaa@aaaa.aaa

    Pipeline demo
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    [CmdletBinding()]
    [Alias('Resolve-OPMailObjectToADUser')]
    PARAM(        
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeline = $True,HelpMessage="Exchange Mailbox object (or it's ExchangeGuid guid string) to be resolved to matching ADUser object[-InputObject `$myExoMailbox]")]
            $InputObject
    ) ;
    BEGIN{

    }
    PROCESS{
        foreach($item in $inputObject){
            TRY{
                switch -regex ($item.gettype().fullname){
                    'System\.Management\.Automation\.PSObject|System\.Collections\.Hashtable|System\.Management\.Automation\.PSCustomObject'{                
                        if($item.ExchangeGuid.GetType().fullname -eq 'System.Guid'){
                            $item = $item.ExchangeGuid
                        }else{
                            throw "unrecognized ExchangeGuid: $($item.ExchangeGuid)" ; 
                        } ; 
                    }
                    'System\.Guid'{
                        if($item.guid){}
                    }
                    'System\.String'{
                        if($item = [guid]$item){}
                    }
                    default{
                        $smsg = "UNRECOGNIZED -inputobject type:$($item.gettype().fullname)" ; 
                        $smsg += "`nPlease specify an ADUser object, or a Guid value" ;
                        write-warning $smsg ;
                        throw $smsg ; 
                    }
                }
                #if($item.gettype().fullname -eq 'System.String'){
                if($item.gettype().fullname -eq 'System.Guid'){
                    if($ADUser = Get-ADUser -Filter { msExchMailboxGuid -eq $item } -Properties msExchMailboxGuid){
                        $ADUser| write-output  ; 
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
        }
    }  # PROC-E; 
} ;
#endregion RESOLVE_MGUSERTOADUSERHARDMATCH ; #*------^ END Resolve-OPMailObjectToADUsermsExchMailboxGuid ^------