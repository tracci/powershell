<#	
	.NOTES
	===========================================================================
	 Created on:   	2018-06-22
	 Created by:   	Tony Raccioppi
	 Organization: 	GitHub
	 Filename:      Get-MailboxFolderSize.ps1 	
	===========================================================================
	.DESCRIPTION
    Exports a user's mailbox stats by folder to determine size.
    Opens CSV file if requested.
#>

# Variables
$Timestamp = Get-Date -Format "yyyy-MM-dd_THHmmss"
$DC = "domaincontroller.domain.com"
$ExchangeServer = "exchange.domain.com"
$Creds = Get-Credential -Message "Please enter your network credentials"
$ExportRoot = "\\path\to\exports"

# Check for AD module
if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Active Directory module not found`r`n" -f red
    Write-Host "Please install the module and re-run the script." -f red
    exit
}

# Connect to Exchange
$Params = @{
    ConfigurationName = "Microsoft.Exchange"
    ConnectionUri     = "http://$ExchangeServer/Powershell/"
    Authentication    = "Kerberos"
    Credential        = $Creds
}

if (!(Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" })) {
    $RemoteEx2013Session = New-PSSession @Params
    Import-PSSession $RemoteEx2013Session -AllowClobber
}

# Validate username and proceed with export
$UserInput = Read-Host -Prompt "`r`nEnter AD username of mailbox to be checked"
$ADcheck = Get-ADUser -Server $DC -Filter { SamAccountName -eq $UserInput }
If ($ADcheck) {
    $Stats = Get-MailboxFolderStatistics $UserInput | Sort-Object -Property ItemsInFolder -Desc | Select-Object Name, FolderSize, ItemsInFolder
    $ExportPath = "$ExportRoot\$Timestamp-$UserInput-MailboxFolders.csv"
    $Stats | Export-CSV $ExportPath -NoTypeInformation
    Write-Host "`r`n$UserInput's mailbox folder stats have been exported to " -NoNewLine; Write-Host "$ExportPath`r`n" -f Blue
    Write-Host "Would you like to open this file now? Default is YES."
    $Readhost = Read-Host "( Y / N ) "
    Switch ($ReadHost) { 
        Y { Invoke-Item $ExportPath }
        N { continue }
        Default { Invoke-Item $ExportPath }
    }
}
Else {
    Write-Host "`nScript failed! $UserInput is not in Active Directory." -ForegroundColor White -BackgroundColor Red
}

Remove-PSSession $RemoteEx2013Session
