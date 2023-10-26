#Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
[CmdletBinding()]
param (
    [string]$ChromeDir="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
  )

if (-Not (Test-Path $ChromeDir -PathType Leaf)) {
  Write-Output "Chrome not found in '$ChromeDir'. Please, install chrome or specify custom chrome location with -ChromeDir argument."
  Exit 1
}

[string]$thisScriptRoot = Join-Path $env:UserProfile "AppData\Local\Programs\Python\Python310"

$chromeDriverRelativeDir = "Scripts"
$chromeDriverDir = $(Join-Path $thisScriptRoot $chromeDriverRelativeDir)
$chromeDriverFileLocation = $(Join-Path $chromeDriverDir "chromedriver.exe")
$chromeVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($ChromeDir).FileVersion
$chromeMajorVersion = $chromeVersion.split(".")[0]

if (-Not (Test-Path $chromeDriverDir -PathType Container)) {
  New-Item -ItemType directory -Path $chromeDriverDir
}

if (Test-Path $chromeDriverFileLocation -PathType Leaf) {
  # get version of curent chromedriver.exe
  $chromeDriverFileVersion = (& $chromeDriverFileLocation --version)
  $chromeDriverFileVersionHasMatch = $chromeDriverFileVersion -match "ChromeDriver (\d+\.\d+\.\d+(\.\d+)?)"
  $chromeDriverCurrentVersion = $matches[1]

  if (-Not $chromeDriverFileVersionHasMatch) {
    Exit
  }
}
else {
  # if chromedriver.exe not found, will download it
  $chromeDriverCurrentVersion = ''
}

if ($chromeMajorVersion -lt 73) {
  # for chrome versions < 73 will use chromedriver v2.46 (which supports chrome v71-73)
  $chromeDriverExpectedVersion = "2.46"
  $chromeDriverVersionUrl = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_STABLE"
}
else {
  $chromeDriverExpectedVersion = $chromeVersion.split(".")[0..2] -join "."
  $chromeDriverVersionUrl = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_STABLE_" + $chromeDriverExpectedVersion
}

$chromeDriverLatestVersion = Invoke-RestMethod -Uri $chromeDriverVersionUrl

Write-Output "chrome version:       $chromeVersion"
Write-Output "chromedriver version: $chromeDriverCurrentVersion"
Write-Output "chromedriver latest:  $chromeDriverLatestVersion"

# will update chromedriver.exe if MAJOR.MINOR.PATCH
$needUpdateChromeDriver = $chromeDriverCurrentVersion -ne $chromeDriverLatestVersion
if ($needUpdateChromeDriver) {
  $chromeDriverZipLink = "https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/" + $chromeDriverLatestVersion + "/win32/chromedriver-win32.zip"
  Write-Output "Will download $chromeDriverZipLink"

  $chromeDriverZipFileLocation = $(Join-Path $chromeDriverDir "chromedriver-win32.zip")

  Invoke-WebRequest -Uri $chromeDriverZipLink -OutFile $chromeDriverZipFileLocation
  Expand-Archive $chromeDriverZipFileLocation -DestinationPath $(Join-Path $thisScriptRoot $chromeDriverRelativeDir) -Force
  Remove-Item -Path $chromeDriverZipFileLocation -Force

  $chromeDriverPath = Join-Path $thisScriptRoot "chromedriver.exe"
  if (Test-Path $chromeDriverPath -PathType Leaf) {
    $chromeDriverFileVersion = & $chromeDriverPath --version
  }

  Move-Item -Path $(Join-Path $thisScriptRoot "Scripts/chromedriver-win32/chromedriver.exe") -Destination $(Join-Path $thisScriptRoot "Scripts")
  Move-Item -Path $(Join-Path $thisScriptRoot "Scripts/chromedriver-win32/LICENSE.chromedriver") -Destination $(Join-Path $thisScriptRoot "Scripts")
  
  Remove-Item -Path $(Join-Path $thisScriptRoot "Scripts/chromedriver-win32") -Force

  Write-Output "chromedriver updated to version $chromeDriverFileVersion"
}
else {
  Write-Output "chromedriver is actual"
}
