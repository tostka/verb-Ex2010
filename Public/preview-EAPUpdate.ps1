#*----------v Function preview-EAPUpdate() v----------
    function preview-EAPUpdate {
        <#
        .SYNOPSIS
        preview-EAPUpdate.ps1 - Code to approximate EmailAddressTemplate-generated email addresses
        .NOTES
        Version     : 0.0.
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2021-08-23
        FileName    : 
        License     : MIT License
        Copyright   : (c) 2021 Todd Kadrie
        Github      : https://github.com/tostka/verb-XXX
        Tags        : Powershell
        AddedCredit : REFERENCE
        AddedWebsite: URL
        AddedTwitter: URL
        REVISIONS
        * 2:00 PM 8/23/2021 added drop of illegal chars (shows up distinctively as spaces in dname :P); fixed bug in regex ps replace; haven't ested balance of fancy substring replace options; %g, %s working.
        .DESCRIPTION
        preview-EAPUpdate.ps1 - Code to approximate EmailAddressTemplate-generated email addresses
        Note: This is a quick & dirty *approximation* of the generated email address. Doesn't support multiple %rxy replaces on mult format codes. Just does the first one; plus any non-replace %d|s|i|m|g's. 
        Don't rely on this, it's just intended to quickly confirm assigned primarysmtpaddress roughly matches the intended EAP template.
        If it doesn't, put eyes on it and confirm, don't use this to drive any revision of the email address!

        Latest specs:[Email address policies in Exchange Server | Microsoft Docs - docs.microsoft.com/](https://docs.microsoft.com/en-us/exchange/email-addresses-and-address-books/email-address-policies/email-address-policies?view=exchserver-2019)

        Address types: SMTP| GWISE| NOTES| X400
        Address format variables:
        |Variable |Value|
        |---|---|
        |%d |Display name|
        |%g |Given name (first name)|
        |%i |Middle initial|
        |%m |Exchange alias|
        |%rxy |Replace all occurrences of x with y|
        |%rxx |Remove all occurrences of x|
        |%s |Surname (last name)|
        |%ng |The first n letters of the first name. For example, %2g uses the first two letters of the first name.|
        |%ns |The first n letters of the last name. For example, %2s uses the first two letters of the last name.|
        Use of email-legal ASCII chars:
        |Example |Exchange Management Shell equivalent|
        |---|---|
        |<alias>@contoso.com |%m@contoso.com|
        |elizabeth.brunner@contoso.com |%g.%s@contoso.com|
        |ebrunner@contoso.com |%1g%s@contoso.com|
        |elizabethb@contoso.com |%g%1s@contoso.com|
        |brunner.elizabeth@contoso.com |%s.%g@contoso.com|
        |belizabeth@contoso.com |%1s%g@contoso.com|
        |brunnere@contoso.com |%s%1g@contoso.com|

        %RXY REPLACEMENT EXAMPLES. 
        |Source properties||
        |---|---|
        |user logon name|"jwilson"|
        |Display name|James C. Wilson|
        |Surname|Wilson|
        |Given name|James|
        note: In %rXY, if X = Y - same char TWICE - the character will be DELETED rather than REPLACED.
        |Replacement String|SMTP Address Generated|Comment|
        |---|---|---|
        |%d@domain.com|JamesCWilson@domain.com|"Displayname@domain"|
        |%g.%s@microsoft.com|James.Wilson@microsoft.com|"givenname.surname@domain"|
        |@microsoft.com|JamesW@microsoft.com|"userLogon@domain" (default)|
        |%1g%s@microsoft.com|JWilson@microsoft.com|"[1stcharGivenName][surname]@domain"|
        |%1g%3s@microsoft.com|JWil@microsoft.com|"[1stcharGivenName][3charsSurname]@domain"|
        |@domain.com|<email-alias>@domain.com (this is the one item always a part of the Default policy)|
        |%r._%d@microsoft.com|JamesC_Wilson@microsoft.com|"[replace periods in displayname with underscore]@domain"|
        |%r..%d@microsoft.com|JamesC.Wilson@microsoft.com|"[DELETE periods in displayname]@domain",(avoids double period if name trails with a period)|
        |%g.%r''%s@domain.com|James.Wilson@domain|"[givenname].[surname,delete all APOSTROPHES]@domain"|
        |%r''%g.%r''%s@domain.com|James.Wilson@domain|"[givenname,delete all APOSTROPHES].[surname,delete all APOSTROPHES]@domain"|

        .PARAMETER  EmailAddressPolicy
        EmailAddressPolicy object to be modeled for primarysmtpaddress update
        .PARAMETER  Recipient
        Recipient object to be modeled
        .PARAMETER useEXOv2
        Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
        .INPUTS
        None. Does not accepted piped input.(.NET types, can add description)
        .OUTPUTS
        None. Returns no objects or output (.NET types)
        System.Boolean
        [| get-member the output to see what .NET obj TypeName is returned, to use here]
        .EXAMPLE
        PS> preview-EAPUpdate  -eap $eaps[16] -Recipient $trcp -verbose ;
        Preview specified recipient using the specified EAP (17th in the set in the $eaps variable).
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate '%5s%3g@contoso.com' -Recipient $trcp -verbose ;
        Preview target recipient using a template that takes surname[0-4]givename[0-3]@contoso.com
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate '%5s%1g%1i@contoso.com' -Recipient $trcp -verbose ;
        Preview target recipient using a template that takes surname[0-4]givename[0]mi[0]@contoso.com
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate '%1g%s@contoso.com' -Recipient $trcp -verbose ;
        Preview target recipient using %1g%s@contoso.com|JWilson@contoso.com|"[1stcharGivenName][surname]@domain"|
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate '%r._%d@contoso.com' -Recipient $trcp -verbose ;
        Preview target recipient using %r._%d@contoso.com|JamesC_Wilson@contoso.com|"[replace periods in displayname with underscore]@domain"|
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate "%g.%r''%s@domain.co" -Recipient $trcp -verbose ;
        Preview target recipient using %g.%r''%s@domain.com|James.Wilson@domain|"[givenname].[surname,delete all APOSTROPHES]@domain"|
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate " %r''%g.%r''%s@domain.com" -Recipient $trcp -verbose ;
        Preview target recipient using %r''%g.%r''%s@domain.com|James.Wilson@domain|"[givenname,delete all APOSTROPHES].[surname,delete all APOSTROPHES]@domain"
        .EXAMPLE
        PS> preview-EAPUpdate  -addresstemplate ' %d@contoso.com' -Recipient $trcp -verbose ;
        Preview target recipient using %d@domain.com|JamesCWilson@domain.com|"Displayname@domain"
        .EXAMPLE
        PS> $genEml = preview-EAPUpdate  -addresstemplate ' %d@contoso.com' -Recipient $trcp -verbose ;
            if(($geneml -ne $trcp.primarysmtpaddress)){
                write-warning "Specified recip's PrimarySmtpAddress does *not* appear to match specified template!`nmanualy *review* the template specs`nand validate that the desired scheme is being applied!"
            }else {
                "PrimarysmtpAddr $($trcp.primarysmtpaddress) roughly conforms to specified template primary addr..."
            } ;
        Example testing output against $trcp primarySmtpAddress.
        .LINK
        https://github.com/tostka/verb-ex2010
        .LINK
        https://docs.microsoft.com/en-us/exchange/email-addresses-and-address-books/email-address-policies/email-address-policies?view=exchserver-2019
        #>
        ###Requires -Version 5
        ###Requires -Modules verb-Ex2010 - disabled, moving into the module
        #Requires -RunasAdministrator
        # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.toro\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
        ## [OutputType('bool')] # optional specified output type
        [CmdletBinding(DefaultParameterSetName='EAP')]
        PARAM(
            [Parameter(ParameterSetName='EAP',Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Exchange EmailAddressPolicy object[-eap `$eaps[16]]")]
            [ValidateNotNullOrEmpty()]
            #[Alias('ALIAS1', 'ALIAS2')]
            $EAP,
            [Parameter(ParameterSetName='Template',Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Specify raw AddressTemplate string, for output modeling[-AddressTemplate '%3g%5s@microsoft.com']")]
            [ValidateNotNullOrEmpty()]
            #[Alias('ALIAS1', 'ALIAS2')]
            $AddressTemplate,
            [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of recipient descriptors: displayname, emailaddress, UPN, samaccountname[-recip some.user@domain.com]")]
            #[ValidateNotNullOrEmpty()]
            #[Alias('ALIAS1', 'ALIAS2')]
            $Recipient,
            [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
            [switch] $useEXOv2
        ) ;
        BEGIN { 
            # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
            ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
            $rgxEmailDirLegalChars = "[0-9a-zA-Z-._+&']" ; 
            reconnect-ex2010 -verbose:$false ; 
    
        } ;  # BEGIN-E
        PROCESS {
            $Error.Clear() ; 
            if($EAP){
                $ptmpl= $eap.EnabledPrimarySMTPAddressTemplate ;
            } elseif($AddressTemplate){
                $ptmpl= $AddressTemplate ;
            } ; 
            $error.clear() ;
            TRY {
                if($Recipient.alias){
                    $Recipient = get-recipient $Recipient.alias ;
                } else { throw "-recipient invalid, has no Alias property!`nMust be a valid Exchange Recipient object" } ; 
                $usr = get-user -id $Recipient.alias -ErrorAction STOP ; 
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #-=-record a STATUSWARN=-=-=-=-=-=-=
                $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                #-=-=-=-=-=-=-=-=
                $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
            } ; 
            $genAddr = $null ; 
            # [0-9a-zA-Z-._+&]{1,64} # alias rgx legal
            # [0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+
            # dirname
            # ^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+ 
            # @domain: @([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,6}$

            #if($ptmpl -match '^@.*$'){$genAddr = $ptmpl = $ptmpl.replace('@',"$($Recipient.alias)@")} ;
            if($ptmpl -match '^@.*$'){
                write-verbose "(matched Alias@domain)" ;
                $string = ($Recipient.alias.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl.replace('@',"$($string)@")
            } ;
            # do replace first, as the %d etc are simpler matches and don't handle the leading %r properly.
            if($ptmpl -match "%r([a-zA-Z_0-9-._+&])([a-zA-Z_0-9-._+&])%(d|g|i|m|s)@.*$"){
                $x = $matches[1] ; 
                $y = $matches[2] ; 
                $vari = $matches[3] ; 
                write-verbose "Parsed:`nx:$($x)`ny:$($y)`nvari:$($vari)" ;
                switch ($vari){
                    'd' {
                        write-verbose "(matched replace $($x) w $($y) on displayname)" ;
                        #$ptmpl = $ptmpl.replace("%d",$usr.displayname.replace($x,$y)) ;
                        $string = ($usr.displayname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                        #$genAddr = $ptmpl = $ptmpl = $ptmpl.replace("%d",$string.replace($x,$y))
                        # subout %rxy first, then the trailing %d w name
                        $genAddr = $ptmpl = $ptmpl = $ptmpl -replace "%r([a-zA-Z_0-9-._+&])([a-zA-Z_0-9-._+&])",'' -replace "%d",$string.replace($x,$y)
                    }
                    'g'  {
                        write-verbose "(matched replace $($x) w $($y) on givenname)" ;
                        #$ptmpl = $ptmpl.replace("%g",$usr.firstname.replace($x,$y)) ;
                        $string = ($usr.firstname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                        #$genAddr = $ptmpl = $ptmpl = $ptmpl.replace("%g",$string.replace($x,$y))
                        $genAddr = $ptmpl = $ptmpl = $ptmpl -replace "%r([a-zA-Z_0-9-._+&])([a-zA-Z_0-9-._+&])",'' -replace "%g",$string.replace($x,$y)
                    }
                    'i' {
                        write-verbose "(matched replace $($x) w $($y) on initials)" ;
                        #$ptmpl = $ptmpl.replace("%i",$usr.Initials.replace($x,$y)) ;
                        $string = ($usr.Initials.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                        #$genAddr = $ptmpl = $ptmpl = $ptmpl.replace("%i",$string.replace($x,$y))
                        $genAddr = $ptmpl = $ptmpl = $ptmpl -replace "%r([a-zA-Z_0-9-._+&])([a-zA-Z_0-9-._+&])",'' -replace "%i",$string.replace($x,$y)
                    }
                    'm' {
                        write-verbose "(matched replace $($x) w $($y) on alias)" ;
                        #$ptmpl = $ptmpl.replace("%m",$Recipient.alias.replace($x,$y)) ;
                        $string = ($Recipient.alias.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                        #$genAddr = $ptmpl = $ptmpl = $ptmpl.replace("%m",$string.replace($x,$y))
                        $genAddr = $ptmpl = $ptmpl = $ptmpl -replace "%r([a-zA-Z_0-9-._+&])([a-zA-Z_0-9-._+&])",'' -replace "%m",$string.replace($x,$y)
                    }
                    's' {
                        write-verbose "(matched replace $($x) w $($y) on surname)" ;
                        #$ptmpl = $ptmpl.replace("%s",$usr.lastname.replace($x,$y)) ;
                        $string = ($usr.lastname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                        #$genAddr = $ptmpl = $ptmpl = $ptmpl.replace("%s",$string.replace($x,$y))
                        $genAddr = $ptmpl = $ptmpl = $ptmpl -replace "%r([a-zA-Z_0-9-._+&])([a-zA-Z_0-9-._+&])",'' -replace "%s",$string.replace($x,$y)
                    }
                    default {
                        throw "unrecognized template: replace (%r) character with no targeted variable (%(d|g|i|m|s))" ;
                    }
                } ;
            } ; 
            if($ptmpl.contains('%g')){
                write-verbose "(matched %g:displayname)" ;
                #$genAddr = $ptmpl = $ptmpl.replace('%g',$usr.firstname)
                $string = ($usr.firstname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl.replace('%g',$string)
            } ;
            if($ptmpl.contains('%s')){
                write-verbose "(matched %s:surname)" ;
                #$genAddr = $ptmpl = $ptmpl.replace('%s',$usr.lastname)} ;
                $string = ($usr.lastname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl.replace('%s',$string)
            };
            if($ptmpl.contains('%d')){
                write-verbose "(matched %d:displayname)" ;
                #$genAddr = $ptmpl = $ptmpl.replace('%d',$usr.displayname)
                $string = ($usr.displayname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl.replace('%d',$string)
            } ;
            if($ptmpl.contains('%i')){
                write-verbose "(matched %i:initials)" ;
                #$genAddr = $ptmpl = $ptmpl.replace('%i',$usr.Initials)} ;
                $string = ($usr.Initials.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl.replace('%i',$string)
            } ;
            if($ptmpl.contains('%m')){
                write-verbose "(matched %m:alias)" ;
                #$genAddr = $ptmpl = $ptmpl.replace('%m',$Recipient.alias)} ;
                $string = ($Recipient.alias.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl.replace('%m',$string)
            } ; 
            if($ptmpl -match '(%(\d)g)'){
                $ltrs = $matches[2] ; 
                write-verbose "(matched %g:displayname, first $($ltrs) chars)" ;
                #$genAddr = $ptmpl = $ptmpl -replace "(%\dg)","$1,$($usr.firstname.substring(0,$ltrs))" ;
                #$genAddr = $ptmpl = $ptmpl -replace "(%\dg)","$($usr.firstname.substring(0,$ltrs))"
                $string = ($usr.firstname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl -replace "(%\dg)","$($string.substring(0,$ltrs))"
            } ; 
            if($ptmpl -match "(%(\d)s)"){
                $ltrs = $matches[2] ; 
                write-verbose "(matched %s:surname, first $($ltrs) chars)" ;
                #$genAddr = $ptmpl = $ptmpl -replace "(%\ds)","$1,$($usr.lastname.substring(0,$ltrs))"
                #$genAddr = $ptmpl =$ptmpl -replace "(%\ds)","$($usr.lastname.substring(0,$ltrs))" ;
                $string = ($usr.lastname.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl -replace "(%\ds)","$($string.substring(0,$ltrs))"
            } ; 
            if($ptmpl -match "(%(\d)i)"){
                $ltrs = $matches[2] ; 
                write-verbose "(matched %i:initials, first $($ltrs) chars)" ;
                #$genAddr = $ptmpl = $ptmpl -replace "(%\ds)","$1,$($usr.lastname.substring(0,$ltrs))"
                #$genAddr = $ptmpl =$ptmpl -replace "(%\di)","$($usr.initials.substring(0,$ltrs))" ;
                $string = ($usr.initials.ToCharArray() |?{$_ -match $rgxEmailDirLegalChars})  -join '' ;
                $genAddr = $ptmpl = $ptmpl -replace "(%\di)","$($string.substring(0,$ltrs))"
            } ; 

        } ;  # PROC-E
        END {
            if($genAddr){
                write-verbose "returning generated address:$($genAddr)" ; 
                $genAddr| write-output 
            } else {
                write-warning "Unable to generate a PrimarySmtpAddress model for user" ; 
                $false | write-output l
            };
        } ;  
    } ; #*------^ END Function preview-EAPUpdate ^------