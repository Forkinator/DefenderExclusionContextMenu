[CmdletBinding(DefaultParameterSetName = 'Install')]
param(
    [Parameter(ParameterSetName = 'Install')]
    [switch]$Install,

    [Parameter(ParameterSetName = 'Uninstall')]
    [switch]$Uninstall,

    [Parameter(ParameterSetName = 'AddExclusion', Mandatory = $true)]
    [string]$AddExclusion
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Add-ExplorerMenuEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseKey
    )

    # A dedicated verb name avoids collisions with other shell extensions.
    $verbKey = Join-Path -Path $BaseKey -ChildPath 'shell\AddToDefenderExclusions'
    $commandKey = Join-Path -Path $verbKey -ChildPath 'command'

    New-Item -Path $verbKey -Force | Out-Null
    New-Item -Path $commandKey -Force | Out-Null

    Set-ItemProperty -Path $verbKey -Name '(Default)' -Value 'Add to Windows Defender Exclusions' -Type String
    Set-ItemProperty -Path $verbKey -Name 'Icon' -Value 'imageres.dll,-5302' -Type String

    $escapedScriptPath = $PSCommandPath.Replace('"', '""')
    $command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$escapedScriptPath`" -AddExclusion `"%1`""
    Set-ItemProperty -Path $commandKey -Name '(Default)' -Value $command -Type String
}

function Remove-ExplorerMenuEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseKey
    )

    $verbKey = Join-Path -Path $BaseKey -ChildPath 'shell\AddToDefenderExclusions'
    if (Test-Path -LiteralPath $verbKey) {
        Remove-Item -LiteralPath $verbKey -Recurse -Force
    }
}

function Ensure-ElevatedForAdd {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    if (Test-IsAdministrator) {
        return $true
    }

    Write-Host 'Elevation required. Re-launching with admin rights...'
    $escapedScriptPath = $PSCommandPath.Replace('"', '""')
    $escapedTargetPath = $TargetPath.Replace('"', '""')
    $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$escapedScriptPath`" -AddExclusion `"$escapedTargetPath`""

    try {
        Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Verb RunAs -WindowStyle Normal | Out-Null
    }
    catch {
        throw "Unable to elevate process. $_"
    }

    return $false
}

function Add-DefenderExclusionPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    if (-not (Test-Path -LiteralPath $TargetPath)) {
        throw "Path does not exist: $TargetPath"
    }

    # Resolve to a stable absolute path before adding exclusion.
    $resolvedPath = (Resolve-Path -LiteralPath $TargetPath).Path
    $existing = @((Get-MpPreference).ExclusionPath)

    if ($existing -contains $resolvedPath) {
        Write-Host "Already excluded: $resolvedPath"
        return
    }

    Add-MpPreference -ExclusionPath $resolvedPath
    Write-Host "Added Defender exclusion: $resolvedPath"
}

switch ($PSCmdlet.ParameterSetName) {
    'Install' {
        # Default behavior if no switch is provided.
        Add-ExplorerMenuEntry -BaseKey 'Registry::HKEY_CURRENT_USER\Software\Classes\*\'
        Add-ExplorerMenuEntry -BaseKey 'Registry::HKEY_CURRENT_USER\Software\Classes\Directory\'
        Write-Host 'Installed context menu for files and folders.'
        break
    }

    'Uninstall' {
        Remove-ExplorerMenuEntry -BaseKey 'Registry::HKEY_CURRENT_USER\Software\Classes\*\'
        Remove-ExplorerMenuEntry -BaseKey 'Registry::HKEY_CURRENT_USER\Software\Classes\Directory\'
        Write-Host 'Removed context menu for files and folders.'
        break
    }

    'AddExclusion' {
        if (Ensure-ElevatedForAdd -TargetPath $AddExclusion) {
            Add-DefenderExclusionPath -TargetPath $AddExclusion
        }
        break
    }
}
