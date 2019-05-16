<#
	.NOTES
	===========================================================================
     Created on:   	2018-08-14
     Created by:   	Tony Racci
	 Organization: 	GitHub
	 Filename:      Set-AccountExpiration.ps1
	===========================================================================
	.DESCRIPTION
    Sets AD account expiration to specific date/time.
#>

# AD module check
if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Active Directory module not found`r`n" -f red
    Write-Host "Please install the module and re-run the script." -f red
    exit
}

# Variables
$DC = "domaincontroller.domain.com"
$Creds = Get-Credential -Message "Please enter your network credentials"
$LogPath = "\\path\to\logs"

# Turn on logging
Get-ChildItem "$LogPath\*.log" | Where-Object LastWriteTime -LT (Get-Date).AddDays(-30) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'yyyy-MM-ddTHHmmss').log"
Start-Transcript $LogPathName

# Prompt for username and time
$Username = Read-Host -Prompt 'Enter username of account to expire'
$ADCheck = Get-ADUser -Server $DC -Credential $Creds -Filter { SamAccountName -eq $Username }

If ($ADCheck) {
    $DateInput = Read-Host "Enter expiration date"
    $Date = Get-Date $DateInput -Format "MM/dd/yyyy"
    $TimeInput = Read-Host "Enter expiration time"
    $Time = Get-Date $TimeInput -Format "HH:mm"
    $Selection = Get-Date "$Date $Time"
    $OutDate = $Selection.DateTime
    $Name = "$($ADCheck.GivenName) $($ADCheck.Surname)"
}
Else {
    Write-Host "`nScript failed! $Username is not in Active Directory." -ForegroundColor Red
    break
}

# Confirm before writing to AD
$Message = Write-Host "`r`nSetting $Name's account to expire: $OutDate" -ForegroundColor Blue
$Question = "Are you sure you want to proceed?"
$Choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$Choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList "&Yes", "Yes, account will be set to expire"))
$Choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList "&No", "No, account unmodified and script will end"))
$Decision = $Host.UI.PromptForChoice($Message, $Question, $Choices, 0)
if ($Decision -eq 0) {
    Write-Host "Confirmed, setting account expiration.`r`n"
    Set-ADUser -Server $DC -Credential $Creds -Identity $Username -AccountExpirationDate $Selection
}
else {
    Write-Host 'Operation cancelled.'
    break
}
