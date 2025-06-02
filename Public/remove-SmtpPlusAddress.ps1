# remove-SmtpPlusAddress.ps1

#region REMOVE_SMTPPLUSADDRESS ; #*------v remove-SmtpPlusAddress v------
#if(-not (gcm remove-SmtpPlusAddress -ea 0)){
    function remove-SmtpPlusAddress {
        <#
        .SYNOPSIS
        remove-SmtpPlusAddress - Strips any Plus address Tag present in an smtp address, and returns the base address
        .NOTES
        Version     : 1.0.0
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2024-05-22
        FileName    : remove-SmtpPlusAddress.ps1
        License     : (none asserted)
        Copyright   : (none asserted)
        Github      : https://github.com/tostka/verb-Ex2010
        Tags        : Powershell,EmailAddress,Version
        AddedCredit : Bruno Lopes (brunokktro )
        AddedWebsite: https://www.linkedin.com/in/blopesinfo
        AddedTwitter: @brunokktro / https://twitter.com/brunokktro
        REVISIONS
        * 2:37 PM 6/2/2025 add to vx10 ; 
        * 1:47 PM 7/9/2024 CBA github field correction
        * 1:22 PM 5/22/2024init
        .DESCRIPTION
        remove-SmtpPlusAddress - Strips any Plus address Tag present in an smtp address, and returns the base address

        Plus Addressing is supported in Exchange Online, Gmail, and other select hosts. 
        It is *not* supported for Exchange Server onprem. Any + addressed email will read as an unresolvable email address. 
        Supporting systems will truncate the local part (in front of the @), after the +, to resolve the email address for normal routing:

        monitoring+whatever@domain.tld, is cleaned down to: monitor@domain.tld. 

        .PARAMETER EmailAddress
        SMTP Email Address
        .OUTPUT
        String
        .EXAMPLE
        PS> 
        PS> $returned = remove-SmtpPlusAddress -EmailAddress 'monitoring+SolarWinds@toro.com';  
        PS> $returned ; 
        Demo retrieving get-EmailAddress, assigning to output, processing it for version info, and expanding the populated returned values to local variables. 
        .EXAMPLE
        ps> remove-SmtpPlusAddress -EmailAddress 'monitoring+SolarWinds@toro.com;notanemailaddresstoro.com,todd+spam@kadrie.net' -verbose ;
        Demo with comma and semicolon delimiting, and an invalid address (to force a regex match fail error).
        .LINK
        https://github.com/brunokktro/EmailAddress/blob/master/Get-ExchangeEnvironmentReport.ps1
        .LINK
        https://github.com/tostka/verb-Ex2010
        #>
        [CmdletBinding()]
        #[Alias('rvExVers')]
        PARAM(
            [Parameter(Mandatory = $true,Position=0,HelpMessage="Object returned by a get-EmailAddress command[-EmailAddress `$ExObject]")]
                [string[]]$EmailAddress
        ) ;
        BEGIN {
            ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
            $verbose = $($VerbosePreference -eq "Continue")
            $rgxSMTPAddress = "([0-9a-zA-Z]+[-._+&='])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}" ; 
            $sBnr="#*======v $($CmdletName): v======" ;
            write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
            if($EmailAddress -match ','){
                $smsg = "(comma detected, attempting split on commas)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $EmailAddress = $EmailAddress.split(',') ; 
            } ; 
            if($EmailAddress -match ';'){
                $smsg = "(semi-colon detected, attempting split on semicolons)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $EmailAddress = $EmailAddress.split(';') ; 
            } ; 
        }
        PROCESS {
            foreach ($item in $EmailAddress){
                if($item -match $rgxSMTPAddress){
                    if($item.split('@')[0].contains('+')){
                        write-verbose  "Remove Plus Addresses from: $($item)" ; 
                        $lpart,$domain = $item.split('@') ; 
                        $item = "$($lpart.split('+')[0])@$($domain)" ; 
                        $smsg = "Cleaned Address: $($item)" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    }
                    $item | write-output ; 
                } else { 
                    write-warning  "$($item)`ndoes not match a standard SMTP Email Address (skipping):`n$($rgxSmtpAddress)" ; 
                    continue ;
                } ; 
            } ;     
        
        } # PROC-E
        END{
            write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
        }
    }; 
#} ; 
#endregion REMOVE_SMTPPLUSADDRESS ; #*------^ END remove-SmtpPlusAddress ^------