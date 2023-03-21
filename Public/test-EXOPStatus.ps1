# test-EXOPStatus.ps1
#*----------v Function test-EXOPStatus() v----------
function test-EXOPStatus {
    <#
    .SYNOPSIS
    test-EXOPStatus.ps1 - Run a quick status confirmation on a block of (or single) OnPrem Exchange servers. Supports name as a wildcard spec.
    .NOTES
    Version     : 0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2022-
    FileName    : test-EXOPStatus.ps1
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell, Exchange, OnPremises,Monitoring,Status
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 4:37 PM 3/21/2023 pulled recursive pound-requires v-x10
    * 10:14 AM 5/2/2022 ren'd get-ExoPStatus.ps1 -> test-EXoPStatus, and moved to verb-ex2010
    * 9:21 AM 5/2/2022 init
    .DESCRIPTION
    test-EXOPStatus.ps1 - Run a quick status confirmation on a block of (or single) OnPrem Exchange servers. Supports name as a wildcard spec.
    .PARAMETER  Server
    Specifies the Mailbox server(s) that you want to test (supports wildcard, as it is a name postfilter applied to the get-exchangeserver cmdlet)[-server Server1*]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    None. Returns no objects or output (.NET types)
    System.Boolean
    [| get-member the output to see what .NET obj TypeName is returned, to use here]
    .EXAMPLE
    PS> test-EXOPStatus.ps1 -server S1B12* -verbose
    Run with server "-like" wildcard filter, verbose
    .EXAMPLE
    PS> test-EXOPStatus.ps1 -server '(S2|S1)B12.*' -verbose ;
    Run with server "-match" regex filter, verbose
    .EXAMPLE
    PS> (Get-DatabaseAvailabilityGroup SomeDAG | select -expand servers ) | test-EXOPStatus.ps1 -verbose ;
    Run Get-DatabaseAvailabiulityGroup servers output array through (auto-resolves -server as an array, with a -like post-filter).
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    #Requires -Version 3
    #requires -PSEdition Desktop
    #Requires -Modules verb-Auth, verb-IO, verb-logging, verb-Network, verb-Text
    #Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("US","GB","AU")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)]#positiveInt:[ValidateRange(0,[int]::MaxValue)]#negativeInt:[ValidateRange([int]::MinValue,0)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    ###[Alias('Alias','Alias2')]
    PARAM(
        #[Parameter(ValueFromPipeline=$true)] # this will cause params to match on matching type, [array] input -> [array]$param
        #[Parameter(ValueFromPipelineByPropertyName=$true)] # this will cause params to
            # match on type, but *also* must have _same param name_ (must be an inbound property
            # named 'arrayVariable' to match the -arrayVariable param, and it must be an
            # [array] type, for the initial match
        # -- if you use both Pipeline & ByPropertyName, you'll get mixed results.
        # -> if it breaks, strip back to ValueFromPipeline and ensure you have type matching on inbound object and typed parameter.
        # see Trace-Command use below, for t-shooting
        #On type matches: ```[array]$arrayVariable``` param will be matched with
        #  inbound pipeline [array] type data, (and other type-to-type matching).  including
        #  typed array variants like: ```[string[]]$stringArrayVariable```
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,HelpMessage="Specifies the Mailbox server(s) that you want to test (supports wildcard, as it is a name postfilter applied to the get-exchangeserver cmdlet)[-server Server1*]")]
        [ValidateNotNullOrEmpty()]
        #[Alias('ALIAS1', 'ALIAS2')]
        [string[]]$Server
    ) ;
    <# #-=-=-=MUTUALLY EXCLUSIVE PARAMS OPTIONS:-=-=-=-=-=
    # designate a default paramset, up in cmdletbinding line
    [CmdletBinding(DefaultParameterSetName='SETNAME')]
    # * set blank, if none of the sets are to be forced (eg optional mut-excl params)
    # * force exclusion by setting ParameterSetName to a diff value per exclusive param

    # example:single $Computername param with *multiple* ParameterSetName's, and varying Mandatory status per set
        [Parameter(ParameterSetName='LocalOnly', Mandatory=$false)]
        $LocalAction,
        [Parameter(ParameterSetName='Credential', Mandatory=$true)]
        [Parameter(ParameterSetName='NonCredential', Mandatory=$false)]
        $ComputerName,
        # $Credential as tied exclusive parameter
        [Parameter(ParameterSetName='Credential', Mandatory=$false)]
        $Credential ;
        # effect:
        -computername is mandetory when credential is in use
        -when $localAction param (w localOnly set) is in use, neither $Computername or $Credential is permitted
        write-verbose -verbose:$verbose "ParameterSetName:$($PSCmdlet.ParameterSetName)"
        Can also steer processing around which ParameterSetName is in force:
        if ($PSCmdlet.ParameterSetName -eq 'LocalOnly') {
            return "some localonly stuff" ;
        } ;
    #
    #-=-reports on which parameters can be used in each parameter set.=-=-=-=-=-=-=
    (gcm SCRIPT.ps1).ParameterSets | Select-Object -Property @{n='ParameterSetName';e={$_.name}}, @{n='Parameters';e={$_.ToString()}} ;
    #-=-=-=-=-=-=-=-=
    #>

    <# #-=-=-=v PARAMETERBINDING TSHOOTING:-=-=-=-=-=
        1. TO T-SHOOT PARAMETER BINDING ERRORS: pull the param block from problem
            func/script, and put just the PARAM() block, and some vari echoes
            in the test-object{} func below: (if you used the unfinished script, the
            trace-command would *excute the full script!*, use a dummy echo version!)

            #region Test Function #*------v Function Test Function v------
            Function Test-Object  {
                # below is the intact cmdletbinding through param block, complete.
                [CmdletBinding()]
                #[ paste in your complete PARAM() BLOCK HERE]
                # The stock PROCESS echos the pipeline ($_) and is tweaked to the echo the explicit variable mappings as well
                PROCESS  {
                    write-host "pipeline `$_ :" ;
                    $_ ;
                    # ECHOING PARAMS
                    write-host "`n`$users:`n$(($users|out-string).trim())`n" ;
                    write-host "`n`$useEXOv2:`n$(($useEXOv2|out-string).trim())`n" ;
                    write-host "`n`$outObject:`n$(($outObject|out-string).trim())`n" ;
                } ;
            } ;
            #endregion Test Function #*------^ END Function Test Function ^------

        2. # TRACE PARAM BINDING ON THE ABOVE: - put your commandline call into the -expression {} block.
            -- Pre-Filter to only the first 2 in arrays - ensure it's [array], but iterate/output all of the objects (it repeats the binding per pipeline element)
            # busier output above using a $Names vari for the -Name spec: (cleaner):
            $Names = @( 'ParameterBinderBase', 'ParameterBinderController', 'ParameterBinding', 'TypeConversion' ) ;
            Trace-Command -Name $Names -Expression {$badaddresses | select -first 2 | Test-Object -outObject }  -PSHost ;
            # or simplified to just parambinding
            Trace-Command -Name ParameterBinding -Expression  {$badaddresses | select -first 2 | Test-Object -outObject }  -PSHost ;

        3. When you run pipeline input on Adv Functions, you'll see that the bound -
            $users - variable is blank in the BEGIN{} block, but suddenly within PROCESS{}
            it's *populated* with a single element - the 1st in array, interatively The
            entire PROCESS BLOCK get's iterated with the bound - $users - variable.

            So to handle both named params & pipeline you need both:
            - $users as a -users param, with  a foreach ($user in $users) in the PROCESS{}
            block, so that  ```some-function -users $arrayofusers``` works.
            - And as a PROCESS {} block AKA "functional pipeline loop", that iterates the
            pipeline through it's bound variable, setting  the vari - $users - to each
            element of the pipeline.
            - so don't bother testing pipeline variable values in BEGIN{}
            (though ,Mandetory=$true is fine, the bindings verified even if the data isn't present yet).
        4. when you Trace-Command a pipeline using the above, you'll see the $_ echo
            empty (it dumps out after full executi8on, like a write-output). And youj'll
            see the bound variable - $users - run by once per each pipeline element, as the
            Process{} block is iterated for each. (along with full parambinding resolution
            info for each iterated process pass).
    #-=-=-=v END PARAMETERBINDING TSHOOTING:-=-=-=-=-=
    #>
    BEGIN {
        #region CONSTANTS-AND-ENVIRO #*======v CONSTANTS-AND-ENVIRO v======

        $pprps = 'PSComputerName', 'Address', 'ProtocolAddress', 'ResponseTime' ;
        $tmprps = 'Server', 'Database', 'Result', 'Error' ;
        $rpprps = 'Server','Check','Result','Error' ;

        # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        # Get parameters this function was invoked with
        #$PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        if ($PSScriptRoot -eq "") {
            if ($psISE) { $ScriptName = $psISE.CurrentFile.FullPath }
            elseif ($context = $psEditor.GetEditorContext()) {$ScriptName = $context.CurrentFile.Path }
            elseif ($host.version.major -lt 3) {
                $ScriptName = $MyInvocation.MyCommand.Path ;
                $PSScriptRoot = Split-Path $ScriptName -Parent ;
                $PSCommandPath = $ScriptName ;
            } else {
                if ($MyInvocation.MyCommand.Path) {
                    $ScriptName = $MyInvocation.MyCommand.Path ;
                    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
                } else {throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$MyInvocation IS BLANK!" } ;
            };
            $ScriptDir = Split-Path -Parent $ScriptName ;
            $ScriptBaseName = split-path -leaf $ScriptName ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($ScriptName) ;
        } else {
            $ScriptDir = $PSScriptRoot ;
            if ($PSCommandPath) {$ScriptName = $PSCommandPath }
            else {
                $ScriptName = $myInvocation.ScriptName
                $PSCommandPath = $ScriptName ;
            } ;
            $ScriptBaseName = (Split-Path -Leaf ((& { $myInvocation }).ScriptName))  ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
        } ;
        if ($showDebug) { write-debug -verbose:$true "`$ScriptDir:$($ScriptDir)`n`$ScriptBaseName:$($ScriptBaseName)`n`$ScriptNameNoExt:$($ScriptNameNoExt)`n`$PSScriptRoot:$($PSScriptRoot)`n`$PSCommandPath:$($PSCommandPath)" ; } ;
        # silently stop any running transcripts
        $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
        #endregion CONSTANTS-AND-ENVIRO #*======^ END CONSTANTS-AND-ENVIRO ^======

        #region START-LOG #*======v START-LOG OPTIONS v======
        #region START-LOG-HOLISTIC #*------v START-LOG-HOLISTIC v------
        # Single log for script/function example that accomodates detect/redirect from AllUsers scope'd installed code, and hunts a series of drive letters to find an alternate logging dir (defers to profile variables)
        #${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        if(!(get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
        foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
        if(!(get-variable rgxPSAllUsersScope -ea 0)){
            $rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps(((d|m))*)1|dll)$" ;
        } ;
        if(!(get-variable rgxPSCurrUserScope -ea 0)){
            $rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps((d|m)*)1|dll)$" ;
        } ;
        $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($whatif) ;} ;
        if($server -is [system.array]){
            $pltSL.Tag = $server -join ',' ;
        } else {
            $pltSL.Tag = $server.replace('*','STAR') ;
        } ;
        if($script:PSCommandPath){
            if(($script:PSCommandPath -match $rgxPSAllUsersScope) -OR ($script:PSCommandPath -match $rgxPSCurrUserScope)){
                $bDivertLog = $true ;
                switch -regex ($script:PSCommandPath){
                    $rgxPSAllUsersScope{$smsg = "AllUsers"}
                    $rgxPSCurrUserScope{$smsg = "CurrentUser"}
                } ;
                $smsg += " context script/module, divert logging into [$budrv]:\scripts"
                write-verbose $smsg  ;
                if($bDivertLog){
                    if((split-path $script:PSCommandPath -leaf) -ne $cmdletname){
                        # function in a module/script installed to allusers|cu - defer name to Cmdlet/Function name
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
                    } else {
                        # installed allusers|CU script, use the hosting script name
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $script:PSCommandPath -leaf)) ;
                    }
                } ;
            } else {
                $pltSL.Path = $script:PSCommandPath ;
            } ;
        } else {
            if(($MyInvocation.MyCommand.Definition -match $rgxPSAllUsersScope) -OR ($MyInvocation.MyCommand.Definition -match $rgxPSCurrUserScope) ){
                $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $script:PSCommandPath -leaf)) ;
            } elseif(test-path $MyInvocation.MyCommand.Definition) {
                $pltSL.Path = $MyInvocation.MyCommand.Definition ;
            } elseif($cmdletname){
                $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
            } else {
                $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$CMDLETNAME, FROM WHICH TO BUILD A START-LOG.PATH!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                BREAK ;
            } ;
        } ;
        write-verbose "start-Log w`n$(($pltSL|out-string).trim())" ;
        $logspec = start-Log @pltSL ;
        $error.clear() ;
        TRY {
            if($logspec){
                $logging=$logspec.logging ;
                $logfile=$logspec.logfile ;
                $transcript=$logspec.transcript ;
                $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                $startResults = start-Transcript -path $transcript ;
            } else {throw "Unable to configure logging!" } ;
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
        #endregion START-LOG-HOLISTIC #*------^ END START-LOG-HOLISTIC ^------

        #endregion START-LOG #*======^ START-LOG OPTIONS ^======

$cwv = get-colorcombo -C 1 ;
    $ccN = get-colorcombo -C 34 ;
    $ccTC = get-colorcombo -C 30 ;
    $ccTSH = get-colorcombo -C 39 ;
    $ccUT = get-colorcombo -C 62 ;
    $ccTM = get-colorcombo -C 37 ;
    $ccRS = get-colorcombo -C 24 ;

        if ($PSCmdlet.MyInvocation.ExpectingInput) {
            write-verbose "Data received from pipeline input: '$($InputObject)'" ;
        } else {
            #write-verbose "Data received from parameter input: '$($InputObject)'" ;
            write-verbose "(non-pipeline - param - input)" ;
        } ;

    } ;  # BEGIN-E
    PROCESS {
        $Error.Clear() ;
        # call func with $PSBoundParameters and an extra (includes Verbose)
        #call-somefunc @PSBoundParameters -anotherParam

        # - Pipeline support will iterate the entire PROCESS{} BLOCK, with the bound - $array -
        #   param, iterated as $array=[pipe element n] through the entire inbound stack.
        # $_ within PROCESS{}  is also the pipeline element (though it's safer to declare and foreach a bound $array param).

        # - foreach() below alternatively handles _named parameter_ calls: -array $objectArray
        # which, when a pipeline input is in use, means the foreach only iterates *once* per
        #   Process{} iteration (as process only brings in a single element of the pipe per pass)

        foreach($item in $server) {

            # dosomething w $item
            # put your real processing in here, and assume everything that needs to happen per loop pass is within this section.
            # that way every pipeline or named variable param item passed will be processed through.
            Write-Output "CHECK EXOP SERVERS MATCHING -SERVER FILTER" | out-null ;


            $error.clear() ;
            if(!(Get-Command Reconnect-Ex2010)){   if(!(Get-Module verb-ex2010 -list)){
                write-warning "MISSING verb-Ex2010 MOD!";
                break ;
            } } else {reconnect-ex2010} ;

            TRY {$srvrs = Get-ExchangeServer } CATCH {
                $ErrTrapd=$Error[0] ;
                write-warning "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
            } ;
            # post filter: determine if -server spec is a regex or like wildcard.

            if($item -is [system.array]){
                write-verbose "(-servers specified is an array - treating as explicit server name via -like postfilter)" ;
                $srvrs = $srvrs |Where-Object{$_.name -like $item} ;
            } elseif(test-IsRegexPattern -pattern $item){
                write-verbose "(-servers specified - $($item): passes test-IsRegexPattern())" ;
                if(([regex]::matches($item,'\*').count) -AND ([regex]::matches($item,'\.').count -eq 0)){
                    write-verbose "(-servers specified - $($item): has wildcard *, but no period => 'like filter')" ;
                    $srvrs = $srvrs |Where-Object{$_.name -like $item}
                } elseIf($srvrs = $srvrs |Where-Object{$_.name -match $item}){
                    write-verbose "(-servers specified - $($item) - worked as a regex, using -match postfilter)" ;
                    # treat it as a regex replace
                    #$haystack -replace $pattern,$newString;
                    #$likeResults | write-output ;
                } elseif ($srvrs = $srvrs |Where-Object{$_.name -like $item}){
                    write-verbose "(-servers specified - $($item) - *failed* as a regex, but worked, using -like postfilter)" ;
                    # use non-regex replace syntax
                    #$target.replace($pattern,$newString);
                    #$likeResults | write-output ;
                } ;
            } elseif ($srvrs = $srvrs |Where-Object{$_.name -like $item}){
                write-verbose "(-servers specified - $($item) - would not pass test-IsRegexPattern: used a -like postfilter)" ;
                # use non-regex replace syntax
                #$target.replace($pattern,$newString);
                #$likeResults | write-output ;
            } ;
            $srvrs = $srvrs | sort-object name ;
            $ttl = ($srvrs|measure-object).count ;
            $procd = 0 ;
            foreach( $srvr in $srvrs){
                $procd++ ;
                $idstr = "($($procd)/$($ttl))" ;
                write-host @cwv "$($idstr):(Test-Connection -ComputerName $($srvr.name) -count 1...)" ;
                $error.clear() ;
                TRY {$ping = Test-Connection -ComputerName $srvr.name -count 1 } CATCH {     $ErrTrapd=$Error[0] ;
                write-warning "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                } ;

                write-host @cwv "$($idstr):(Test-ServiceHealth -Server $($srvr.name))..." ;
                $error.clear() ;
                TRY {$sHlth = Test-ServiceHealth -Server $srvr.name } CATCH {     $ErrTrapd=$Error[0] ;
                write-warning "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                } ;

                write-host @cwv "$($idstr):(get-uptime -ComputerName $($srvr.name)...)" ;
                $uptime = (get-uptime -ComputerName $srvr.name).uptimestr ;

                if($srvr.serverrole.split(',').trim() -contains 'Mailbox'){

                    write-host @cwv "$($idstr):(Test-ReplicationHealth -id $($srvr.name)...)" ;
                    TRY {$tRepl = Test-ReplicationHealth -id $srvr.name} CATCH {       $ErrTrapd=$Error[0] ;
                        write-warning "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                    } ;

                    write-host @cwv "$($idstr):(Test-MAPIConnectivity -S $($srvr.name)...)" ;
                    TRY {$tMapi = Test-MAPIConnectivity -S $srvr.name} CATCH {       $ErrTrapd=$Error[0] ;
                        write-warning "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                    } ;

                    write-host @ccN "`n==$($idstr):$($srvr.name):" ;
                    write-host @ccTC "`n`n--$($idstr):Test-Connection:`n$(($ping | Format-Table -a $pprps |out-string).trim())" ;
                    write-host @ccTSH "`n`n--$($idstr):Test-ServiceHealth:`n$(($shlth | Format-List Role,RequiredServicesRunning,ServicesNotRunning|out-string).trim())" ;
                    if($tRepl){ write-host @ccTM "`n`n--$($idstr):Test-ReplicationHealth:`n$(($tRepl | format-table -a $rpprps |out-string).trim())" ;};
                    if($tMapi){ write-host @ccRS "`n`n--$($idstr):Test-MAPIConnectivity:`n$(($tMapi | format-table -a $tmprps |out-string).trim())" ;}
                    else {
                        write-host @ccTM "`n`n--$($idstr):Test-MAPIConnectivity:`n(The operation could not be performed because no mailbox database is currently hosted on server $($srvr.name))" ;
                    }

                } ; # if-E 'Mailbox'
                write-host @ccUT "`n`n--$($idstr):get-uptime:$(($uptime|out-string).trim())`n" ;
            } ; # loop-E $srvr in $srvrs

        } # loop-E # $item in $server

    } ;  # PROC-E
    END {
        $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
        write-host $stopResults ;
    } ;  # END-E
} ;
#*------^ END Function test-EXOPStatus ^------
