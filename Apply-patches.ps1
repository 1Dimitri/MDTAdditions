
function Import-Csv {
    <#

    .ForwardHelpTargetName Import-Csv
    .ForwardHelpCategory Cmdlet

    #>

    [CmdletBinding(DefaultParameterSetName='Delimiter')]
    param(
        [Parameter(ParameterSetName='Delimiter', Position=1)]
        [ValidateNotNull()]
        [System.Char]
        ${Delimiter},

        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Alias('PSPath')]
        [System.String[]]
        ${Path},

        [Parameter(ParameterSetName='UseCulture', Mandatory=$true)]
        [ValidateNotNull()]
        [Switch]
        ${UseCulture},

        [ValidateNotNullOrEmpty()]
        [System.String[]]
        ${Header},

        [ValidateNotNullOrEmpty()]
        [System.String[]]
        ${Type},

        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable]
        ${TypeMap},

        [ValidateNotNullOrEmpty()]
        [System.String]
        ${As},

        [Alias('UseETS')]
        [ValidateNotNull()]
        [Switch]
        ${UseExtendedTypeSystem},

        [ValidateNotNull()]
        [Switch]
        ${OverwriteTypeHierarchy}
    )

    begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Import-Csv', [System.Management.Automation.CommandTypes]::Cmdlet)

            #region Initialize helper variables used in the processing of the additional parameters.
            $scriptCmdPipeline = ''
            #endregion

            #region Process and remove the Type parameter if it is present, modifying the pipelined command appropriately.
            if ($Type) {
                $PSBoundParameters.Remove('Type') | Out-Null
                $scriptCmdPipeline += @'
 | ForEach-Object {
    for ($index = 0; ($index -lt @($_.PSObject.Properties).Count) -and ($index -lt @($Type).Count); $index++) {
        $typeObject = [System.Type](@($Type)[$index])
        $propertyName = @($_.PSObject.Properties)[$index].Name
        $_.$propertyName = & $ExecutionContext.InvokeCommand.NewScriptBlock("[$($typeObject.FullName)]`$_.`$propertyName")
    }
    $_
}
'@
            }
            #endregion

            #region Process and remove the TypeMap parameter if it is present, modifying the pipelined command appropriately.
            if ($TypeMap) {
                $PSBoundParameters.Remove('TypeMap') | Out-Null
                $scriptCmdPipeline += @'
 | ForEach-Object {
     foreach ($key in $TypeMap.keys) {
        if ($TypeMap[$key] -is [System.Type]) {
            $typeObject = $TypeMap[$key]
        } else {
            $typeObject = [System.Type]($TypeMap[$key])
        }
        $_.$key = & $ExecutionContext.InvokeCommand.NewScriptBlock("[$($typeObject.FullName)]`$_.`$key")
    }
    $_
}
'@
            }
            #endregion

            #region Process and remove the As, UseExtendedTypeSystem and OverwriteTypeHierarchy parameters if they are present, modifying the pipelined command appropriately.
            if ($As) {
                $PSBoundParameters.Remove('As') | Out-Null
                $customTypeName = $As
                if ($UseExtendedTypeSystem) {
                    $PSBoundParameters.Remove('UseExtendedTypeSystem') | Out-Null
                    $customTypeName = '$($_.PSObject.TypeNames[0] -replace ''#.*$'','''')#$As'
                }
                if ($OverwriteTypeHierarchy) {
                    $PSBoundParameters.Remove('OverwriteTypeHierarchy') | Out-Null
                    $scriptCmdPipeline += @"
 | ForEach-Object {
     `$typeName = "$customTypeName"
     `$_.PSObject.TypeNames.Clear()
    `$_.PSObject.TypeNames.Insert(0,`$typeName)
    `$_
}
"@
                } else {
                    $scriptCmdPipeline += @"
 | ForEach-Object {
     `$typeName = "$customTypeName"
    `$_.PSObject.TypeNames.Insert(0,`$typeName)
    `$_
}
"@
                }
            } else {
                if ($UseExtendedTypeSystem) {
                    $PSBoundParameters.Remove('UseExtendedTypeSystem') | Out-Null
                }
                if ($OverwriteTypeHierarchy) {
                    $PSBoundParameters.Remove('OverwriteTypeHierarchy') | Out-Null
                }
            }
            #endregion

            $scriptCmd = {& $wrappedCmd @PSBoundParameters}

            #region Append our pipeline command to the end of the wrapped command script block.
            $scriptCmd = $ExecutionContext.InvokeCommand.NewScriptBlock(([string]$scriptCmd + $scriptCmdPipeline))
            #endregion

            $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch {
            throw
        }
    }

    process {
        try {
            $steppablePipeline.Process($_)
        }
        catch {
            throw
        }
    }

    end {
        try {
            $steppablePipeline.End()
        }
        catch {
            throw
        }
    }
}


Function Get-Patches {
param(
$patches,
[string] $TargetPath,
[string] $TempDir
)


$patches | ForEach-Object {
   Write-Verbose "Downloading $($_.URL)"
   $url = $_.URL
   $patchname = $url.Split('/')[-1]
   $filename =  ("{0:000}-" -f $_.Order)+$patchname
   $filepath = Join-Path $TargetPath $filename
   Start-BitsTransfer -Priority Foreground -Source $url -Destination $filepath -DisplayName $filename

   $_ | Add-Member -MemberType NoteProperty LocalPath $filepath
   $_ | Add-Member -MemberType NoteProperty LocalName $filename
   $_ | Add-Member -MemberType NoteProperty PatchName $patchname

   $cmdnoun =  $_.PackageFlags
   If ($cmdnoun -ne "") {
        $order = $_.Order
        Write-Verbose "Performing $($cmdnoun) transformation on files "
        $fnname = 'New-'+$cmdnoun
        Write-Verbose "Function to call: $fnname"
        if (Get-Command $fnname -ea SilentlyContinue) {
             Write-Verbose "calling $fnname"
            New-Directory $TempDir
            $newfiles = Invoke-Expression "& $fnname -Filename $($filepath) -TempDir $TempDir"

            Write-Verbose $newfiles
             $_ | Add-Member -MemberType NoteProperty ProcessedNames @()
             $o = $_
            $newfiles | Foreach-Object {
               Write-Verbose "Handling: $_ [$($_.FullName)]"
               $newfn = ("{0:000}-" -f $Order)+($_.Name)
               $newpath = Join-Path $TargetPath $newfn
                Write-Verbose "Moving $_ to $newpath"
                Move-Item $($_.FullName) $newpath -Force
                $o.ProcessedNames+=$newpath
                Write-Verbose "New file names: $newpath"
              }
              

       } # Get-Command
       else {
        Write-Verbose "transformation wasn't found/run correctly, using original file names"
         $_ | Add-Member -MemberType NoteProperty ProcessedNames @($filepath)
       } # else Get-Command
    } else {
      Write-Verbose "No transformation on files needed"
         $_ | Add-Member -MemberType NoteProperty ProcessedNames @($filepath)
        }
  }

  $patches
}

Function New-Directory {
param(
[string[]] $Path
)
  Write-Verbose "Removing Paths"
  $mp = (Get-WindowsImage -Mounted).MountPath
  $Path | ForEach-Object {
    Write-Verbose "Removing $_"
    if ($_ -in ((Get-WindowsImage -Mounted).MountPath)) {
        # Mounted Dir is owned by Trusted Installer
          Dismount-WindowsImage -Path $mp -Discard
        }
        else {
        Remove-Item -Path $_ -Force -Recurse -EA SilentlyContinue
        }
    }
  Write-Verbose "Creating Paths"
  
  $Path | ForEach-Object { 
  Write-Verbose "Creating $_"
  New-Item  -ItemType Directory -Path $_ -Ea SilentlyContinue | Out-Null
  }
}

Function New-IECAB {
param(
  [string]$filename,
  [string]$TempDir=$env:TEMP
)

    New-Directory $TempDir
	Write-Verbose "Extracting $($FileName) in $($TempDir)"
	# Wait for the end Out-Null trick
	& $filename /X:"$TempDir" | Out-Null
	
	Get-ChildItem -LiteralPath $TempDir -recurse | % {
		if ( $_.Name -ne 'IE-Win7.CAB') {  
			try 
			{
				Remove-Item $_.FullName -Force 
			} catch [UnauthorizedAccessException]
			{
				Write-Host "Permission access error on $($_.FullName)"
			}
		}
	}
 
   Get-ChildItem -LiteralPath $TempDir
  
}



Function Apply-PatchesToIndexImage {
param(
$Patches,
[string] $ImagePath,
[int] $ImageIndex,
[string] $LogDir,
[string]$MountedPath
)
  Set-ItemProperty -Path $ImagePath -name IsReadOnly -value $false
  #$LogDir
  # $MountedPath = 'E:\MountDir'
  $TempDir = 'Q:\Temp'
 
  New-Directory -Path $MountedPath,$LogDir,$TempDir

  $LogPath = Join-Path $LogDir "000-Mount.log"
  $BuildLogPath = Join-Path (Join-Path $MountedPath 'Windows') "WindowsBuild.log"

  Mount-WindowsImage -ImagePath $ImagePath -Index $ImageIndex -Path $MountedPath -LogPath $LogPath -ScratchDirectory $TempDir


  # Passing all packages in the folder at once doesn't work e.g. KB 3125574 after servicing stack
  # Add-WindowsPackage -PackagePath $PatchesDir -Path $MountedPath -LogPath $LogPath -NoRestart
  $Patches | ForEach-Object {
   
   $patch = $_
   $_.ProcessedNames | Foreach-Object {
   $fname = split-path $_ -Leaf
   $LogPath = Join-Path $LogDir ($fname+'.log')
  
    Write-Verbose "Add-WindowsPackage -PackagePath [$_] -Path [$MountedPath] -LogPath [$LogPath] -NoRestart -ScratchDirectory [$TempDir]"
       "$($_.PatchName) [$($_.Comments)]" | Add-Content $BuildLogPath
    Add-WindowsPackage -PackagePath $_ -Path $MountedPath -LogPath $LogPath -NoRestart -ScratchDirectory $TempDir    }
 
  
    if ($_.DismFlags -contains "remount") {
      Write-Verbose "remounting..."
      $orderAsString = "{0:000}" -f $patch.Order
      $LogPath = Join-Path $LogDir "$($orderAsString)-Dismount.log"
      Write-Verbose "Dismount-WindowsImage -Path [$MountedPath] -Save -LogPath [$LogPath] -ScratchDirectory [$TempDir]"
      Dismount-WindowsImage -Path $MountedPath -Save -LogPath $LogPath -ScratchDirectory $TempDir
      $LogPath = Join-Path $LogDir "$($orderAsString)-Mount.log"
      Write-Verbose "Mount-WindowsImage -ImagePath [$ImagePath] -Index [$ImageIndex] -Path [$MountedPath] -LogPath [$LogPath] -ScratchDirectory [$TempDir]"
      Mount-WindowsImage -ImagePath $ImagePath -Index $ImageIndex -Path $MountedPath -LogPath $LogPath -ScratchDirectory $TempDir
      Write-Verbose "remounting done"
    }

  }

  
  $LogPath = Join-Path $LogDir "999-Dismount.log"
  
    Dismount-WindowsImage -Path $MountedPath -Save -LogPath $LogPath -ScratchDirectory $TempDir

}

function Apply-PatchListToOS {
param(
[string] $SourceRoot,
[string] $DestinationPath,
[string] $TempPath,
[string] $PatchesList,
[string] $LogDir,
[string]$MountedPath,
[int[]] $Indexes
)
$InstallWimFolder = Join-Path $SourceRoot 'sources'
$InstallWinPath = Join-Path $InstallWimFolder 'install.wim'
New-Directory $TempPath
$TempInstallWimPath = Join-Path $TempPath 'install.wim'
$DestinationInstallWimPath = Join-Path (Join-Path $DestinationPath 'sources') 'install.wim'
#$csv = Import-CSV 'E:\patchesround\win2008r2-patches-201706.csv' -TypeMap @{Order='Int';URL='String';DismFlags='String';PackageFlags='String';Comments='String'}
$csv = Import-CSV $PatchesList -TypeMap @{Order='Int';URL='String';DismFlags='String';PackageFlags='String';Comments='String'}
$Patches = Get-Patches -Patches $csv -TargetPath 'E:\downloads2' -TempDir 'Q:\TempCAB'
Copy-Item $InstallWinPath $TempInstallWimPath -Force
Write-Verbose "Getting WIM Details from $TempInstallWimPath"
if ($PSBoundParameters.ContainsKey('Indexes')) {
   $ImageIndexes = $Indexes
}
else {
    $ImageInfo = Get-WindowsImage -ImagePath $TempInstallWimPath 
    $ImageIndexes = $ImageInfo.ImageIndex 
}

$ImageIndexes | ForEach-Object { Apply-PatchesToIndexImage -Patches $patches -ImagePath $TempInstallWimPath -ImageIndex $_ -LogDir $LogDir -MountedPath $MountedPath}
New-Directory $DestinationPath
robocopy $SourceRoot $DestinationPath /E /XF install.wim
Copy-Item  $TempInstallWimPath $DestinationInstallWimPath -Force
}

#Apply-PatchListToOS -SourceRoot 'H:\' -DestinationPath E:\result -TempPath Q:\TempWIM -PatchesList E:\patchesround\win2012r2-patches-201708.csv -MOuntedPath 'E:\MountDir' -LOgDir 'Q:\Log'
#Apply-PatchListToOS -SourceRoot 'H:\' -DestinationPath E:\result -TempPath Q:\TempWIM -PatchesList E:\patchesround\win2012r2-patches-201708.csv -MOuntedPath 'Q:\DirTEst' -LogDIr 'Q:\Log'
#Apply-PatchListToOS -SourceRoot 'I:\' -DestinationPath E:\result2008r2 -TempPath Q:\TempWIM -PatchesList E:\patchesround\win2008r2-patches-201709.csv -MountedPath 'E:\MountDir' -LogDir 'Q:\Log'
Apply-PatchListToOS -SourceRoot 'I:\' -DestinationPath E:\result2008r2 -TempPath Q:\TempWIM -PatchesList E:\patchesround\win2008r2-patches-201709.csv -MountedPath 'E:\MountDir' -LogDir 'Q:\Log' -Indexes 1