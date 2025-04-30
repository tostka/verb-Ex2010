# Initialize-xopEventIDTable.ps1

#region INITIALIZE_XOPEVENTIDTABLE ; #*------v Initialize-xopEventIDTable v------
if(-not(gci function:Initialize-xopEventIDTable -ea 0)){
    function Initialize-xopEventIDTable {
        <#
        .SYNOPSIS
        Initialize-xopEventIDTable - Builds an indexed hash tabl of Exchange Server Get-MessageTrackingLog EventIDs
        .NOTES
        Version     : 1.0.0
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2025-04-22
        FileName    : Initialize-xopEventIDTable
        License     : (none asserted)
        Copyright   : (none asserted)
        Github      : https://github.com/tostka/verb-Ex2010
        Tags        : Powershell,EmailAddress,Version
        AddedCredit : Bruno Lopes (brunokktro )
        AddedWebsite: https://www.linkedin.com/in/blopesinfo
        AddedTwitter: @brunokktro / https://twitter.com/brunokktro
        REVISIONS
        * 2;58 pm 4/28/2025 Updated table again, and found Ex2016/19 eventid specifications online, added. Did find that 
        the online doc doesn't document the edge SendExternal event id (added below, manually).             
        * 1:47 PM 7/9/2024 CBA github field correction
        * 1:22 PM 5/22/2024init
        .DESCRIPTION
        Initialize-xopEventIDTable - Builds an indexed hash tabl of Exchange Server Get-MessageTrackingLog EventIDs

        ## Exchange 2019 EventID reference:

        [Event types in the message tracking log | Microsoft Learn](https://learn.microsoft.com/en-us/exchange/mail-flow/transport-logs/message-tracking?view=exchserver-2019#event-types-in-the-message-tracking-log)

        Doesn't include Edge eventid: 
        "SENDEXTERNAL          | A message was sent by SMTP to sent to the SMTP server responsible to receive the email for the external email address."
    
        (needs to be manually spliced in below 'SEND' during updates from source MS documentation)
    
        .OUTPUT
        System.Collections.Hashtable returns an Indexed Hash of EventIDs EventName to Description
        .EXAMPLE
        PS> $eventIDLookupTbl = Initialize-EventIDTable ; 
        PS> $smsg = "`n`n## EventID Definitions:" ; 
        PS> $TrackMsgs | group eventid | select -expand Name | foreach-object{                   
        PS>     $smsg += "`n$(($eventIDLookupTbl[$_] | ft -hidetableheaders |out-string).trim())" ; 
        PS> } ; 
        PS> $smsg += "`n`n" ; 
        PS> if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
        PS> else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        Demo resolving histogram eventid uniques, to MS documented meansings of each event id in the msgtrack.        
        .LINK
        https://github.com/tostka/verb-Ex2010
        #>
        [CmdletBinding()]
        #[Alias('ALIASNAME')]
        PARAM() ;
        BEGIN {
            ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
            $verbose = $($VerbosePreference -eq "Continue")
            $rgxSMTPAddress = "([0-9a-zA-Z]+[-._+&='])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}" ; 
            $sBnr="#*======v $($CmdletName): v======" ;
            write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
            $eventIDsMD = @"
EventName             | Description                                                                                                                                                                                                                                                                                                                                                                                                                                                                            
--------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
AGENTINFO             | This event is used by transport agents to log custom data.                                                                                                                                                                                                                                                                                                                                                                                                                             
BADMAIL               | A message submitted by the Pickup directory or the Replay directory that can't be delivered or returned.                                                                                                                                                                                                                                                                                                                                                                               
CLIENTSUBMISSION      | A message was submitted from the Outbox of a mailbox.                                                                                                                                                                                                                                                                                                                                                                                                                                  
DEFER                 | Message delivery was delayed.                                                                                                                                                                                                                                                                                                                                                                                                                                                          
DELIVER               | A message was delivered to a local mailbox.                                                                                                                                                                                                                                                                                                                                                                                                                                            
DELIVERFAIL           | An agent tried to deliver the message to a folder that doesn't exist in the mailbox.                                                                                                                                                                                                                                                                                                                                                                                                   
DROP                  | A message was dropped without a delivery status notification (also known as a DSN, bounce message, non-delivery report, or NDR). For example:<br/>- Completed moderation approval request messages.<br/>- Spam messages that were silently dropped without an NDR.                                                                                                                                                                                                                     
DSN                   | A delivery status notification (DSN) was generated.                                                                                                                                                                                                                                                                                                                                                                                                                                    
DUPLICATEDELIVER      | A duplicate message was delivered to the recipient. Duplication may occur if a recipient is a member of multiple nested distribution groups. Duplicate messages are detected and removed by the information store.                                                                                                                                                                                                                                                                     
DUPLICATEEXPAND       | During the expansion of the distribution group, a duplicate recipient was detected.                                                                                                                                                                                                                                                                                                                                                                                                    
DUPLICATEREDIRECT     | An alternate recipient for the message was already a recipient.                                                                                                                                                                                                                                                                                                                                                                                                                        
EXPAND                | A distribution group was expanded.                                                                                                                                                                                                                                                                                                                                                                                                                                                     
FAIL                  | Message delivery failed. Sources include SMTP, DNS, QUEUE, and ROUTING.                                                                                                                                                                                                                                                                                                                                                                                                                
HADISCARD             | A shadow message was discarded after the primary copy was delivered to the next hop. For more information, see Shadow redundancy in Exchange Server.                                                                                                                                                                                                                                                                                                                                   
HARECEIVE             | A shadow message was received by the server in the local database availability group (DAG) or Active Directory site.                                                                                                                                                                                                                                                                                                                                                                   
HAREDIRECT            | A shadow message was created.                                                                                                                                                                                                                                                                                                                                                                                                                                                          
HAREDIRECTFAIL        | A shadow message failed to be created. The details are stored in the source-context field.                                                                                                                                                                                                                                                                                                                                                                                             
INITMESSAGECREATED    | A message was sent to a moderated recipient, so the message was sent to the arbitration mailbox for approval. For more information, see Manage message approval.                                                                                                                                                                                                                                                                                                                       
LOAD                  | A message was successfully loaded at boot.                                                                                                                                                                                                                                                                                                                                                                                                                                             
MODERATIONEXPIRE      | A moderator for a moderated recipient never approved or rejected the message, so the message expired. For more information about moderated recipients, see Manage message approval.                                                                                                                                                                                                                                                                                                    
MODERATORAPPROVE      | A moderator for a moderated recipient approved the message, so the message was delivered to the moderated recipient.                                                                                                                                                                                                                                                                                                                                                                   
MODERATORREJECT       | A moderator for a moderated recipient rejected the message, so the message wasn't delivered to the moderated recipient.                                                                                                                                                                                                                                                                                                                                                                
MODERATORSALLNDR      | All approval requests sent to all moderators of a moderated recipient were undeliverable, and resulted in non-delivery reports (also known as NDRs or bounce messages).                                                                                                                                                                                                                                                                                                                
NOTIFYMAPI            | A message was detected in the Outbox of a mailbox on the local server.                                                                                                                                                                                                                                                                                                                                                                                                                 
NOTIFYSHADOW          | A message was detected in the Outbox of a mailbox on the local server, and a shadow copy of the message needs to be created.                                                                                                                                                                                                                                                                                                                                                           
POISONMESSAGE         | A message was put in the poison message queue or removed from the poison message queue.                                                                                                                                                                                                                                                                                                                                                                                                
PROCESS               | The message was successfully processed.                                                                                                                                                                                                                                                                                                                                                                                                                                                
PROCESSMEETINGMESSAGE | A meeting message was processed by the Mailbox Transport Delivery service.                                                                                                                                                                                                                                                                                                                                                                                                             
RECEIVE               | A message was received by the SMTP receive component of the transport service or from the Pickup or Replay directories (source: SMTP), or a message was submitted from a mailbox to the Mailbox Transport Submission service (source: STOREDRIVER).                                                                                                                                                                                                                                    
REDIRECT              | A message was redirected to an alternative recipient after an Active Directory lookup.                                                                                                                                                                                                                                                                                                                                                                                                 
RESOLVE               | A message's recipients were resolved to a different email address after an Active Directory lookup.                                                                                                                                                                                                                                                                                                                                                                                    
RESUBMIT              | A message was automatically resubmitted from Safety Net. For more information, see Safety Net in Exchange Server.                                                                                                                                                                                                                                                                                                                                                                      
RESUBMITDEFER         | A message resubmitted from Safety Net was deferred.                                                                                                                                                                                                                                                                                                                                                                                                                                    
RESUBMITFAIL          | A message resubmitted from Safety Net failed.                                                                                                                                                                                                                                                                                                                                                                                                                                          
SEND                  | A message was sent by SMTP between transport services.                                                                                                                                                                                                                                                                                                                                                                                                                                 
SENDEXTERNAL          | A message was sent by SMTP to sent to the SMTP server responsible to receive the email for the external email address.                                                                                                                                                                                                                                                                                                                                                                                                                           
SUBMIT                | The Mailbox Transport Submission service successfully transmitted the message to the Transport service. For SUBMIT events, the source-context property contains the following details:<br/>- MDB: The mailbox database GUID.<br/>- Mailbox: The mailbox GUID.<br/>- Event: The event sequence number.<br/>- MessageClass: The type of message. For example, IPM.Note.<br/>- CreationTime: Date-time of the message submission.<br/>- ClientType: For example, User, OWA, or ActiveSync.
SUBMITDEFER           | The message transmission from the Mailbox Transport Submission service to the Transport service was deferred.                                                                                                                                                                                                                                                                                                                                                                          
SUBMITFAIL            | The message transmission from the Mailbox Transport Submission service to the Transport service failed.                                                                                                                                                                                                                                                                                                                                                                                
SUPPRESSED            | The message transmission was suppressed.                                                                                                                                                                                                                                                                                                                                                                                                                                               
THROTTLE              | The message was throttled.                                                                                                                                                                                                                                                                                                                                                                                                                                                             
TRANSFER              | Recipients were moved to a forked message because of content conversion, message recipient limits, or agents. Sources include ROUTING or QUEUE.
"@ ; 
            # UPDATE NOTE: MANUAL UNDOCUMENTED ADDITION: "SENDEXTERNAL          | A message was sent by SMTP to sent to the SMTP server responsible to receive the email for the external email address."
            # (needs to be manually spliced in below 'SEND' during updates from source MS documentation)
            $Object = $eventIDsMD | convertfrom-MarkdownTable ; 
            $Key = 'EventName' ; 
            $Hashtable = @{}
        }
        PROCESS {
            Foreach ($Item in $Object){
                $Procd++ ; 
                $Hashtable[$Item.$Key.ToString()] = $Item ; 
                if($ShowProgress -AND ($Procd -eq $Every)){
                    write-host -NoNewline '.' ; $Procd = 0 
                } ; 
            } ;                 
        } # PROC-E
        END{
            $Hashtable | write-output ; 
            write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
        }
    };
} ;      
#endregion INITIALIZE_XOPEVENTIDTABLE ; #*------^ Initialize-xopEventIDTable ^------
