param(
    [string]$BaseTemplate = (Join-Path $PSScriptRoot "base_template.dotm"),
    [string]$Template = (Join-Path $PSScriptRoot "..\build\base_template.dotm"),
    [string]$Source = (Join-Path $PSScriptRoot "VBA.bas"),
    [string]$ModuleName = "Module1"
)

# Copies the Word-authored base template into build\ and replaces its macro
# module with src\VBA.bas. Word only exposes VBProject when "Trust access to
# the VBA project object model" is on, so it is enabled for this run and
# restored.

$ErrorActionPreference = "Stop"
$Source = (Resolve-Path $Source).Path
$code = [IO.File]::ReadAllText($Source) -replace "`r?`n", "`r`n"

New-Item -ItemType Directory -Force (Split-Path $Template) | Out-Null
Copy-Item (Resolve-Path $BaseTemplate).Path $Template -Force
$Template = (Resolve-Path $Template).Path

$securityKey = "HKCU:\Software\Microsoft\Office\16.0\Word\Security"
if (-not (Test-Path $securityKey)) {
    New-Item $securityKey | Out-Null
}
$previousTrust = (Get-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
$runningWord = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = $null
$document = $null
$ownsWord = $false
$wordIds = @()

try {
    Set-ItemProperty $securityKey -Name AccessVBOM -Value 1 -Type DWord

    $word = New-Object -ComObject Word.Application
    $wordIds = @(Get-Process WINWORD | Where-Object { $runningWord -notcontains $_.Id } | ForEach-Object Id)
    $ownsWord = $wordIds.Count -gt 0
    if (-not $ownsWord) {
        throw "Word automation attached to an already-running Word instance. Close Word and retry."
    }
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3

    $document = $word.Documents.Open($Template, $false, $false, $false)
    foreach ($component in $document.VBProject.VBComponents) {
        $lines = $component.CodeModule.CountOfLines
        if ($component.Name -ne $ModuleName -and $lines -gt 0 -and
            $component.CodeModule.Lines(1, $lines) -match "Sub Ribbon_") {
            throw "Ribbon callbacks also exist in $($component.Name); they would be ambiguous."
        }
    }

    $module = $document.VBProject.VBComponents.Item($ModuleName).CodeModule
    if ($module.CountOfLines -gt 0) {
        $module.DeleteLines(1, $module.CountOfLines)
    }
    $module.AddFromString($code)
    $document.Save()
    Write-Output "vba=$Template ($($module.CountOfLines) lines in $ModuleName)"
}
finally {
    try {
        if ($document) {
            $document.Close(0)
            [Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
        }
        if ($word) {
            if ($ownsWord) {
                $word.Quit(0)
            }
            [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
    }
    catch {
        Write-Warning "Word cleanup failed: $_"
    }

    # Word writes its Trust Center state back on exit, so restore afterwards.
    # A hidden instance that ignores Quit is ours and is force-closed.
    if ($wordIds.Count -gt 0) {
        Wait-Process -Id $wordIds -Timeout 60 -ErrorAction SilentlyContinue
        Get-Process -Id $wordIds -ErrorAction SilentlyContinue | Stop-Process -Force
        Wait-Process -Id $wordIds -Timeout 10 -ErrorAction SilentlyContinue
    }
    if ($null -eq $previousTrust) {
        Remove-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue
    }
    else {
        Set-ItemProperty $securityKey -Name AccessVBOM -Value $previousTrust -Type DWord
    }
}
