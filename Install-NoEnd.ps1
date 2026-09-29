<#
.SYNOPSIS
  Builds NoEnd.ppam from the VBA sources in .\src and installs it into PowerPoint.
.PARAMETER Uninstall
  Unloads and removes the add-in instead.
#>
param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'

$AddinName   = 'NoEnd'
$InstallDir  = Join-Path $env:APPDATA 'NoEnd'
$AddinPath   = Join-Path $InstallDir "$AddinName.ppam"
$SrcDir      = Join-Path $PSScriptRoot 'src'
$SecurityKey = 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security'
$TrustKey    = Join-Path $SecurityKey 'Trusted Locations\NoEndAddin'

$ppSaveAsOpenXMLAddin = 30
$msoTrue  = -1
$msoFalse = 0

if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    Write-Host 'Please close PowerPoint first, then run this again.' -ForegroundColor Yellow
    exit 1
}

function Remove-RegisteredAddin($app) {
    for ($i = $app.AddIns.Count; $i -ge 1; $i--) {
        $a = $app.AddIns.Item($i)
        if ($a.FullName -like "*\$AddinName.ppam") {
            $a.Loaded = $msoFalse
            $app.AddIns.Remove($i)
        }
    }
}

if ($Uninstall) {
    $app = New-Object -ComObject PowerPoint.Application
    try { Remove-RegisteredAddin $app } finally { $app.Quit() }
    Remove-Item $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $TrustKey -Force -ErrorAction SilentlyContinue
    Write-Host 'NoEnd add-in removed.' -ForegroundColor Green
    exit 0
}

# Temporarily allow scripted access to the VBA project so the modules can be imported.
New-Item -Force -Path $SecurityKey -ErrorAction SilentlyContinue | Out-Null
$oldVbom = (Get-ItemProperty -Path $SecurityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty -Path $SecurityKey -Name AccessVBOM -Value 1 -Type DWord

# Trust only the install folder, so the add-in loads without a macro prompt.
New-Item -Force -Path $TrustKey | Out-Null
Set-ItemProperty -Path $TrustKey -Name Path -Value "$InstallDir\"
Set-ItemProperty -Path $TrustKey -Name Description -Value 'NoEnd PowerPoint add-in'
Set-ItemProperty -Path $TrustKey -Name AllowSubfolders -Value 0 -Type DWord

$app = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    Remove-RegisteredAddin $app

    New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
    if (Test-Path $AddinPath) { Remove-Item $AddinPath -Force }

    $pres = $app.Presentations.Add($msoFalse)
    $proj = $pres.VBProject
    $proj.Name = $AddinName
    $proj.VBComponents.Import((Join-Path $SrcDir 'modNoEnd.bas')) | Out-Null
    $proj.VBComponents.Import((Join-Path $SrcDir 'clsNoEndEvents.cls')) | Out-Null
    $pres.SaveAs($AddinPath, $ppSaveAsOpenXMLAddin)
    $pres.Close()

    $addin = $app.AddIns.Add($AddinPath)
    $addin.AutoLoad = $msoTrue
    $addin.Loaded   = $msoTrue

    Write-Host "NoEnd installed: $AddinPath" -ForegroundColor Green
    Write-Host 'It loads automatically every time PowerPoint starts.'
}
finally {
    if ($app) {
        $app.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    }
    if ($null -eq $oldVbom) {
        Remove-ItemProperty -Path $SecurityKey -Name AccessVBOM -ErrorAction SilentlyContinue
    } else {
        Set-ItemProperty -Path $SecurityKey -Name AccessVBOM -Value $oldVbom -Type DWord
    }
}
