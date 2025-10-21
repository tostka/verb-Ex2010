# get-IISLocalBoundCertifcateTDO.ps1


#region GET_IISLOCALBOUNDCERTIFCATETDO ; #*------v get-IISLocalBoundCertifcateTDO v------
Function get-IISLocalBoundCertifcateTDO {
        <#
        .SYNOPSIS
        get-IISLocalBoundCertifcateTDO - Retrieves the certificate bound to the IIS Default Website (also handy for independantly verifyinig Exchange Service bindings, when get-exchangecertificate is mangled by Exchange Auth cert expiration). 
        .NOTES
        Version     : 0.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 20250929-1026AM
        FileName    : get-IISLocalBoundCertifcateTDO
        License     : MIT License
        Copyright   : (c) 2025 Todd Kadrie
        Github      : https://github.com/tostka/verb-io
        Tags        : Powershell,IIS,Web,Website,Exchange,Certificate
        AddedCredit : 
        AddedWebsite: 
        AddedTwitter: 
        REVISIONS
        * 2:40 PM 10/1/2025 updated CBH example; added to xopBuildLibrary.psm1
        * 4:01 PM 9/30/2025 init
        .DESCRIPTION
        get-IISLocalBoundCertifcateTDO - Retrieves the certificate bound to the IIS Default Website (also handy for independantly verifyinig Exchange Service bindings, when get-exchangecertificate is mangled by Exchange Auth cert expiration). 
        
        Driven by get-exchangecertificate fundemental on-install break, when Nov2023 SU patching packet signing mandates break any time the Exchange Auth certificate is non-functional
        Unfortunately, fixing the issue requires rerunning Hybrid Configuration Wizard as part of the process (Change approval requiremment PITA).
        So this is one way to commandline confirm Exchange Cert Service binding, without a functional get-exchangecertificate return (or using EAC web site).
        .INPUTS
        None, no piped input.
        .OUTPUTS
        System.Security.Cryptography.X509Certificates.X509Certificate2 bound certificate object
        .EXAMPLE
        PS> if($iisCert = get-IISLocalBoundCertifcateTDO){
        PS>     write-host "IIS has a bound certificate: implies the IIS ExchangeCertifidate Binding is intact"
        PS> } ; 
        Demo default output        
        .LINK
        https://github.org/tostka/verb-network/
        #>
        [CmdletBinding()]
        #[alias('get-LocalDiskFreeSpace')]
        PARAM() ; 
        if(-not (get-module Webadministration)){import-module -name Webadministration -fo -verb } ; 
        if($site = Get-ChildItem -Path "IIS:\Sites" | where {( $_.Name -eq "Default Web Site" )}){
            if($binding = $site.bindings.collection | ?{$_.protocol -eq 'https' -and $_.bindingInformation -eq ':443:'}){                
                $smsg = "`n$(($binding | fl * |out-string).trim())" ; 
                if($VerbosePreference -eq "Continue"){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                if($iisCert = gci -Path "cert:\localmachine\$($binding.certificateStoreName)\$($binding.certificateHash)"){
                    $smsg = "IIS bound cert:`nSubjectName: $($iiscert.subjectname.name)`n$(($iisCert | fl thumbprint,notbefore,notafter,hasprivatekey|out-string).trim())" ; 
                    $smsg += "`nSANS:`n`n$(($iisCert.DnsNameList.unicode|out-string).trim())" ; 
                    write-host -foregroundcolor green $smsg ;   
                    [pscustomobject]$iisCert | write-output ;           
                }else{
                    $smsg = "No bound certificate found for local IIS Default Web Site" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ; 
            } ; 
        }else{
            $smsg = "No local IIS Sites found" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ;  
    }
#endregion GET_IISLOCALBOUNDCERTIFCATETDO ; #*------^ END get-IISLocalBoundCertifcateTDO ^------

