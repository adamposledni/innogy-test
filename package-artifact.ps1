<#
.SYNOPSIS
    Packages this repo's Fabric items into a zip for the "Deploy solution" notebook to consume.

.DESCRIPTION
    Zips every top-level item folder (Lakehouse, notebooks, pipelines, variable library) except the
    deployment notebook's own folder and VCS/tooling directories. Upload the resulting zip to blob
    storage and generate a read-only SAS URL for it.

.PARAMETER OutputPath
    Path of the zip file to create. Defaults to artifact.zip in the repo root.
#>
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot "artifact.zip")
)

$excludeDirs = @(".git", ".claude", "Deploy solution.Notebook")

$itemsToZip = Get-ChildItem -Path $PSScriptRoot -Directory | Where-Object { $excludeDirs -notcontains $_.Name }

if ($itemsToZip.Count -eq 0) {
    throw "No item folders found to package."
}

if (Test-Path $OutputPath) {
    Remove-Item $OutputPath -Force
}

Compress-Archive -Path $itemsToZip.FullName -DestinationPath $OutputPath

Write-Host "Packaged $($itemsToZip.Count) item(s) into $OutputPath"
$itemsToZip | ForEach-Object { Write-Host "  - $($_.Name)" }
