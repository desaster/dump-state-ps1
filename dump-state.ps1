param(
    [string]$Prefix = $env:computerName,
    [string]$Postfix = "",
    [string]$OutputPath = ""
)

# these settings should make the script stop on error
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$PSDefaultParameterValues['*:ErrorAction']='Stop'

# listing drivers requires administrator privileges
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
        [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    break
}

# the module(s) and output folder should be in the same directory as the main
# script, therefore they can be accessed by $PSScriptRoot
Import-Module $PSScriptRoot\installedapplication.psm1
if ([string]::IsNullOrEmpty($OutputPath) -eq $true) {
    $path = Join-Path -Path $PSScriptRoot -ChildPath "output"
} else {
    $path = $OutputPath
}

# create unique folder
$stamp = $(get-date -f yyyyMMdd) + $(get-date -f HHmmss)
if ([string]::IsNullOrEmpty($Prefix) -eq $false) {
    $stamp = "$($Prefix.ToLower())_$stamp"
}
if ([string]::IsNullOrEmpty($Postfix) -eq $false) {
    $stamp = "$($stamp)_$Postfix"
}
$path = Join-Path -Path $path -ChildPath $stamp
New-Item -ItemType directory -Path $path | Out-Null

# List drivers by using the built-in Get-WindowsDriver call
$driversFile = Join-Path $path -ChildPath "drivers.csv"
Write-Host "Gathering drivers"
Get-WindowsDriver -Online |
    Sort-Object -Property ClassGuid, Driver |
    Export-Csv -NoTypeInformation -Path $driversFile

# List installed applications by
# using the provided Get-InstalledApplication cmdlet
$appsFile = Join-Path $path -ChildPath "applications.csv"
Write-Host "Gathering applications"
Get-InstalledApplication -OutputType CSV -outfile $appsFile

# List App Packages
$appPackagesFile = Join-Path $path -ChildPath "apppackages.csv"
Write-Host "Gathering app packages"
Get-AppxPackage -AllUsers |
    Sort-Object -Property Name, Version |
    Select-Object -Property Name, Version |
    Export-Csv -NoTypeInformation -Path $appPackagesFile

# List services by using the built-in Get-Service call
$servicesFile = Join-Path $path -ChildPath "services.csv"
Write-Host "Gathering services"
Get-Service |
    Sort-Object -Property Name |
    Export-Csv -NoTypeInformation -Path $servicesFile

# List programs & programdata in the windows default folders
$foldersFile = Join-Path $path -ChildPath "folders.csv"
Write-Host "Gathering application folders"
Get-Item 'C:\Program Files\*' |
    Select-Object -Property Parent, Name, CreationTime |
    Sort-Object -Property Parent, Name |
    Export-Csv -NoTypeInformation -Path $foldersFile
Get-Item 'C:\Program Files (x86)\*' |
    Select-Object -Property Parent, Name, CreationTime |
    Sort-Object -Property Parent, Name |
    Export-Csv -NoTypeInformation -Append -Path $foldersFile
Get-Item 'C:\ProgramData\*' |
    Select-Object -Property Parent, Name, CreationTime |
    Sort-Object -Property Parent, Name |
    Export-Csv -NoTypeInformation -Append -Path $foldersFile

# List startup programs
$startupFile = Join-Path $path -ChildPath "startup.csv"
Write-Host "Gathering startup programs"
Get-CimInstance Win32_StartupCommand |
    Select-Object Name, command, Location, User |
    Sort-Object -Property Name, Command |
    Export-Csv -NoTypeInformation -Path $startupFile

# List scheduled tasks, excluding the ones in \Microsoft\
$tasksFile = Join-Path $path -ChildPath "tasks.csv"
Write-Host "Gathering scheduled tasks"
Get-ScheduledTask |
    Where-Object {$_.TaskPath -notlike "\Microsoft\*" } |
    Sort-Object -Property URI |
    Select-Object -Property TaskPath, Author, TaskName, State, Triggers |
    Export-Csv -NoTypeInformation -Path $tasksFile

# List start menu folders
$startmenuFile = Join-Path $path -ChildPath "startmenu.csv"
Write-Host "Gathering start menu folders"
Get-ChildItem -Recurse -Directory -Path "$([Environment]::GetFolderPath('StartMenu'))" |
    Select-Object -Property Name, Parent, FullName, CreationTime |
    Sort-Object -Property FullName |
    Export-Csv -NoTypeInformation -Path $startmenuFile
Get-ChildItem -Recurse -Directory -Path "$([Environment]::GetFolderPath('CommonStartMenu'))" |
    Select-Object -Property Name, Parent, FullName, CreationTime |
    Sort-Object -Property FullName |
    Export-Csv -NoTypeInformation -Append -Path $startmenuFile

# BIOS version
$biosFile = Join-Path $path -ChildPath "bios.csv"
Write-Host "Gathering BIOS version"
Get-WmiObject win32_bios | Export-Csv -NoTypeInformation -Path $biosFile

# Output the unique folder name so it can be copied over to a memo
Write-Host $path

# Pause so the powershell window doesn't disappear until the user has seen the
# folder name
Read-Host -Prompt "Press Enter to continue"
