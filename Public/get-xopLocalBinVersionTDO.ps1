# get-xopLocalBinVersionTDO.ps1


#region GET_XOPLOCALEXSETUPVERSIONTDO ; #*------v get-xopLocalExSetupVersionTDO v------
    Function get-xopLocalExSetupVersionTDO {
        <#
        .SYNOPSIS
        get-xopLocalExSetupVersionTDO - Discover local Exchange Server CAB ExSetup.exe (or use specified -ExSetupPath), from common paths and return summary (FullName,Name,ProductVersion,Length,Lastwritetime)
        .NOTES
        Version     : 0.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 20250929-1026AM
        FileName    : get-xopLocalExSetupVersionTDO.ps1
        License     : MIT License
        Copyright   : (c) 2025 Todd Kadrie
        Github      : https://github.com/tostka/verb-ex2010
        Tags        : ExchangeServer,Version,Install,Maintenance
        AddedCredit : 
        AddedWebsite: 
        AddedTwitter: 
        REVISIONS
        * 1:29 PM 10/2/2025 init 
        .DESCRIPTION
        get-xopLocalExSetupVersionTDO - Discover local Exchange Server CAB ExSetup.exe (or use specified -ExSetupPath), from common paths and return summary (FullName,Name,ProductVersion,Length,Lastwritetime)
        .PARAMETER ExSetupPath
        Optional full path to ExSetup.exe file to be examined [-ExSetupPath c:\pathto\ExSetup.exe]
        .INPUTS
        None, no piped input.
        .OUTPUTS
        PSCustomObject summary of ExSetup.exe (FullName,Name,ProductVersion,Length,Lastwritetime)
        .EXAMPLE
        PS> $cabinfo = get-xopLocalExSetupVersionTDO -verbose 
        PS> $cabinfo 

        15:22:22:No -ExSetupPath: Attempting to discover latest local cab version, hunting across drives:r|d|c
        15:22:23:Taking first resolved $CabExSetup:

            FullName       : D:\cab\ExchangeServer2016-x64-CU23-ISO\unpacked\Setup\ServerRoles\Common\ExSetup.EXE
            Name           : ExSetup.EXE
            ProductVersion : 15.1.2507.6
            Length         : 36256
            LastWriteTime  : 3/26/2022 3:02:53 PM

        Demo autodiscovery hunting through configured drives on standard paths
        .EXAMPLE
        PS> $cabinfo = get-xopLocalExSetupVersionTDO -ExSetupPath "D:\cab\ExchangeServer2016-x64-CU23-ISO\unpacked\Setup\ServerRoles\Common\ExSetup.EXE" -verbose ; 
        PS> $cabinfo 

            FullName       : D:\cab\ExchangeServer2016-x64-CU23-ISO\unpacked\Setup\ServerRoles\Common\ExSetup.EXE
            Name           : ExSetup.EXE
            ProductVersion : 15.1.2507.6
            Length         : 36256
            LastWriteTime  : 3/26/2022 3:02:53 PM

        Demo resolving against a specified full path to the ExSetup.exe to be examined.
        .LINK
        https://github.org/tostka/verb-ex2010/
        #>
        [CmdletBinding()]
        [alias('get-xopLocalExSetupVersion')]
        PARAM(
            [Parameter(Mandatory = $False,Position = 0,ValueFromPipeline = $True, HelpMessage = 'Optional full path to ExSetup.exe file to be examined [-ExSetupPath c:\pathto\ExSetup.exe]')]
                [Alias('PsPath')]
                #[AllowNull()]
                #[ValidateScript({Test-Path $_ -PathType 'Container'})]
                #[System.IO.DirectoryInfo[]]$Path,
                [ValidateScript({Test-Path $_})]
                [system.io.fileinfo[]]$ExSetupPath
                #[string[]]$ExSetupPath,
        ) ;  
        BEGIN{}
        PROCESS{       
            if(-not $ExSetupPath){
                if(-not (get-variable CabDrives -ea 0)){$CabDrives = 'r','d','c' };        
                $smsg = "No -ExSetupPath: Attempting to discover latest local cab version, hunting across drives:$($CabDrives -join '|')" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                # Resolve LATEST LOCAL CAB VERSION, HUNTING ACROSS DRIVES
                #$SourcePath = 'D:\cab\ExchangeServer2016-x64-CU23-ISO\unpacked\Setup\ServerRoles\Common\ExSetup.EXE'  ; 
                # wildcard to span versions & cu/su combos.
                $SourcePath = 'D:\cab\ExchangeServer*-x64-*-ISO\unpacked\Setup\ServerRoles\Common\ExSetup.EXE'  ; 
                $SourceLeaf = ($SourcePath.split('\') | select -skip 1 ) -join '\' ;     
                foreach($cabdrv in $CabDrives){
                    if(-not (test-path -path  "$($cabdrv):" -ea 0)){Continue} ;
                    $testpath = (join-path -path "$($cabdrv):" -child $SourceLeaf) ;
                    $CabExSetup = resolve-path $testpath | select -expand path |foreach-object{
                        $thisfile = gci $_ ;
                        $finfo = [ordered]@{
                            FullName = $thisfile.fullname;
                            Name = $thisfile.Name ; 
                            ProductVersion = [version]$thisfile.versioninfo.productversion ; 
                            Length = $thisfile.length ; 
                            LastWriteTime = $thisfile.LastWriteTime ; 
                        } ;
                        [pscustomobject]$finfo | write-output ;            
                    } | sort productversion | select -last 1 ;
                    if($CabExSetup){
                        $smsg = "Taking first resolved `$CabExSetup:`n`n$(($CabExSetup|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        $CabExSetup | write-output ; 
                        Break ; 
                    } ; 
                } ;
            } else{
                FOREACH($exfile IN $ExSetupPath){
                    $thisfile = gci $EXFILE -ea STOP ;
                    $finfo = [ordered]@{
                        FullName = $thisfile.fullname;
                        Name = $thisfile.Name ; 
                        ProductVersion = [version]$thisfile.versioninfo.productversion ; 
                        Length = $thisfile.length ; 
                        LastWriteTime = $thisfile.LastWriteTime ; 
                    } ;
                    [pscustomobject]$finfo | write-output 
                } ; 
            }; 
        } ;  # PROC-E
    } ; 
    #endregion GET_XOPLOCALEXSETUPVERSIONTDO ; #*------^ END get-xopLocalExSetupVersionTDO ^------