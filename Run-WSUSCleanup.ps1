<#
	.NOTES
	===========================================================================
     Created on:   	2018-10-10
     Created by:   	Tony Racci
	 Organization: 	GitHub
	 Filename:      Run-WSUSCleanup.ps1
	===========================================================================
	.DESCRIPTION
    Runs WSUS cleanup. Requires 'UpdateServices' PS Module. Adjust parameters accordingly and run on a schedule.
#>

# Variables
$Server = "WSUS.domain.com"
[int]$Port = 8530
$WSUS = Get-WsusServer -Name $Server -PortNumber $Port
$LogPath = "\\path\to\logs"

# Check WSUS module
if (!(Get-Module -ListAvailable -Name UpdateServices)) {
    Write-Host "Update Services module not found`r`n" -f red
    Write-Host "Please install the module and re-run the script." -f red
    exit
}

# Parameters
$Params = @{
    CleanupObsoleteComputers    = $true
    CleanupObsoleteUpdates      = $true
    CleanupUnneededContentFiles = $true
    DeclineExpiredUpdates       = $true
    DeclineSupersededUpdates    = $true
    CompressUpdates             = $true
}

# Turn on logging
Get-ChildItem "$LogPath\*.log" | Where-Object LastWriteTime -LT (Get-Date).AddDays(-15) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'MM-dd-yyyy').log"
Start-Transcript $LogPathName -Append

$StartTime = (Get-Date)
Invoke-WsusServerCleanup -UpdateServer $WSUS @Params
$EndTime = (Get-Date)

Write-Output "`r`nElapsed Time:"
New-TimeSpan -Start $StartTime -End $EndTime | Format-List Hours, Minutes, Seconds

Stop-Transcript
