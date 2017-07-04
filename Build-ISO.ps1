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
             $newnames = $_.ProcessedNames
            $newfiles | Foreach-Object {
               Write-Verbose "Handling: $_ [$($_.FullName)]"
               $newfn = ("{0:000}-" -f $Order)+($_.Name)
               $newpath = Join-Path $TargetPath $newfn
                Write-Verbose "Moving $_ to $newpath"
                Move-Item $($_.FullName) $newpath
                $newnames+=$newpath
              }

       } # Get-Command
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
  $Path | Remove-Item -Force -Recurse -EA SilentlyContinue
  Write-Verbose "Creating Paths"
  
  $Path | ForEach-Object { 
  Write-Verbose "Creating $_"
  New-Item  -ItemType Directory -Path $_ | Out-Null
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
[int] $ImageIndex
)
  Set-ItemProperty -Path $ImagePath -name IsReadOnly -value $false
  $MountedPath = 'E:\MountDir'
  $LogDir ='Q:\Logs'
  $TempDir = 'Q:\Temp'
 
  New-Directory -Path $MountedPath,$LogDir,$TempDir

  $LogPath = Join-Path $LogDir "000-Mount.log"
  Mount-WindowsImage -ImagePath $ImagePath -Index $ImageIndex -Path $MountedPath -LogPath $LogPath -ScratchDirectory $TempDir

  # Passing all packages in the folder at once doesn't work e.g. KB 3125574 after servicing stack
  # Add-WindowsPackage -PackagePath $PatchesDir -Path $MountedPath -LogPath $LogPath -NoRestart
  $Patches | ForEach-Object {
   
   $LogPath = Join-Path $LogDir "$($_.LocalName).log"
  
    Add-WindowsPackage -PackagePath $_.LocalPath -Path $MountedPath -LogPath $LogPath -NoRestart -ScratchDirectory $TempDir
  
    if ($_.DismFlags -contains "remount") {
      Write-Verbose "remounting..."
      Dismount-WindowsImage -Path $MountedPath -Save -LogPath $LogPath -ScratchDirectory $TempDir
      Mount-WindowsImage -ImagePath $ImagePath -Index $ImageIndex -Path $MountedPath -LogPath $LogPath -ScratchDirectory $TempDir
      Write-Verbose "remounting done"
    }

  }

  
  $LogPath = Join-Path $LogDir "999-Dismount.log"
  
    Dismount-WindowsImage -Path $MountedPath -Save -LogPath $LogPath -ScratchDirectory $TempDir

}

$csv = Import-CSV 'E:\patchesround\win2008r2-patches-201706.csv' -TypeMap @{Order='Int';URL='String';DismFlags='String';PackageFlags='String'}
Get-Patches -Patches $csv -TargetPath 'E:\downloads2' -TempDir 'Q:\TempCAB'
#Apply-PatchesToIndexImage -PatchesDir 'E:\downloads2' -ImagePath Q:\install.wim -ImageIndex 1


