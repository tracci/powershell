<#
	.NOTES
	===========================================================================
	 Created on:   	2018-07-19
     Created by:   	Tony Racci
	 Organization: 	GitHub
	 Filename:      Remove-MailMessage.ps1
	===========================================================================
	.DESCRIPTION
    Removes a mail message from ALL MAILBOXES. Run with caution!
#>

# Connect to Exchange PowerShell
$ExchangeServer = "exchange.domain.com"     # Exchange server
$Creds = Get-Credential -Message "Please enter your network credentials"
$Params = @{
    ConfigurationName = "Microsoft.Exchange"
    ConnectionUri     = "http://$ExchangeServer/Powershell/"
    Authentication    = "Kerberos"
    Credential        = $Creds
}
$RemoteExSession = New-PSSession @Params
Import-PSSession $RemoteExSession

# Variables
$Timestamp = Get-Date -Format "yyyy-MM-dd_THHmmss"
$ExportRoot = "\\path\to\exports"
$FileName = "$Timestamp-DeletedEmails.csv"
$Export = "$ExportRoot\$FileName"
$LogPath = "\\path\to\logs"

# Enter email criteria, perferably as specific as possible.
$Subject = "Email subject"
$From = "Sender's email address"
$Sent = "MM/DD/YYYY"

# Turn on logging
Get-ChildItem "$LogPath\*.log" | Where-Object LastWriteTime -LT (Get-Date).AddDays(-30) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'yyyy-MM-ddTHHmmss').log"
Start-Transcript $LogPathName

Get-Mailbox -ResultSize Unlimited | Search-Mailbox -SearchQuery Subject:$Subject, From:$From, Sent:$Sent -DeleteContent | Export-Csv -Path $Export -Append

Stop-Transcript
