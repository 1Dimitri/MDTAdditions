param(
 [System.IO.DirectoryInfo] $SourceFolder
)

if ($SourceFolder -ne $null) {
    Write-Host "Copying package providers from [$SourceFolder]"
    $pkgdir =  $Env:ProgramFiles
    Copy-Item -Path "$SourceFolder" -Destination $pkgdir -Recurse -Container
}

Write-Host "Importing NuGet Package provider in current session"
Import-PackageProvider -Name "Nuget" -WA SilentlyContinue

Write-Host "Registering Package Sources as NUGet Sources"

$NUGetLocalFeeds = $TSEnvList:NUGetLocalFeeds
$ChocoLocalFeeds = $TSEnvList:ChocoLocalFeeds

if (($ChocolocalFeeds -eq $null) -and ($NUGetLocalFeeds -eq $null)) {
  Write-Host "No ChocoLocalFeeds, no NuGetLocalFeeds, nothing to do"
  exit
}
if ($NUGetLocalFeeds -eq $null)
{
 $NUGetLocalFeeds = $ChocoLocalFeeds
}


$NUGetLocalFeeds | Foreach-Object {
$feedurl = $_
$feedname = $_ -replace 'https?://','' -replace '[^\d\w]','_'
Write-Host "Adding feed $feedurl [$feedname]"

Register-PackageSource $feedurl -Name $feedname  -Trusted -ProviderName NuGet
}

if ($ChocoLocalFeeds -ne $null) {
    Write-Host "Installing Chocolatey from one of these sources"
    Find-Package Chocolatey | Install-Package
    Write-Host "Chocolatey installed"

    $installed = Get-Package 'Chocolatey'
    if ($installed -ne $null) {
     $installpath = Split-Path $installed.Source -Parent
     $cmdpath = Join-Path $installpath 'tools'
     $cmd = Join-Path $cmdpath 'chocolateyInstall.ps1'
     & $cmd
     $ec = $LasrExitCode
     Write-Host "ChocolateyInstall return code: $ec" 
    }
}


