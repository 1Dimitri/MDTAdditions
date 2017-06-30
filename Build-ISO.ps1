Function Get-PatchesFromURL {
param(
[string[]] $patchURLs,
[string] $TargetPath
)

$i = 0

$patchURLs | ForEach-Object {
   $url = $_
   $i+=1;
   $filename =  ("{0:000}-" -f $i)+$url.Split('/')[-1]

   $filepath = Join-Path $TargetPath $filename
   Start-BitsTransfer -Priority Foreground -Source $url -Destination $filepath -DisplayName $filename
}

}


Function Apply-PatchesToIndexImage {
param(
[string] $PatchesDir,
[string] $ImagePath,
[int] $ImageIndex
)
  Set-ItemProperty -Path $ImagePath -name IsReadOnly -value $false
  $MountedPath = 'E:\MountDir'
  $LogDir ='Q:\Logs'
  $TempDir = 'Q:\Temp'
 
  $MountedPath,$LogDir,$TempDir | Remove-Item -Force -Recurse
  $MountedPath,$LogDir,$TempDir | ForEach-Object { New-Item  -ItemType Directory -Path $_ }

  # Passing all packages in the folder at once doesn't work e.g. KB 3125574 after servicing stack
  # Add-WindowsPackage -PackagePath $PatchesDir -Path $MountedPath -LogPath $LogPath -NoRestart
  Get-ChildItem $PatchesDir -REcurse | Sort-Object Name | ForEach-Object {
   $LogPath = Join-Path $LogDir "$($_.Name).log"
    Mount-WindowsImage -ImagePath $ImagePath -Index $ImageIndex -Path $MountedPath -LogPath $LogPath -ScratchDirectory $TempDir
    Add-WindowsPackage -PackagePath $_.FullName -Path $MountedPath -LogPath $LogPath -NoRestart -ScratchDirectory $TempDir
    Dismount-WindowsImage -Path $MountedPath -Save -LogPath $LogPath -ScratchDirectory $TempDir

  }
  
}

$urls = (Import-CSV 'E:\patchesround\win2008r2-patches-201706.csv' -Header URL).URL
Get-PatchesFromURL $urls -TargetPath 'E:\downloads2'
Apply-PatchesToIndexImage -PatchesDir 'E:\downloads2' -ImagePath Q:\install.wim -ImageIndex 1


