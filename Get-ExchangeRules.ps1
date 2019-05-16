<#
	.NOTES
	===========================================================================
	 Created on:   	2018-05-18
	 Created by:   	Tony Raccioppi
	 Organization: 	GitHub
     Filename:      Get-ExchangeRules.ps1
	===========================================================================
	.DESCRIPTION
    Gets a list of a user's server-side rules (on-prem Exchange)
    Outputs to HTML or CSV. HTML displays easier.
#>

# Variables
$Timestamp = Get-Date -Format "yyyy-MM-dd_THHmmss"
$ExchangeServer = "exchange.domain.com"
$Creds = Get-Credential -Message "Please enter your network credentials"
$ExportRoot = "\\path\to\exports"

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

$Email = Read-Host "Enter email address to check"
$ExportPath = "$ExportRoot\$Timestamp-$Email-MailboxRules.htm"
# $ExportPathCSV = "$ExportRoot\$Timestamp-$Email-MailboxRules.csv"             # Uncomment to use the CSV output below.
$Rules = Get-InboxRule -Mailbox $Email | Select-Object Name, Enabled, Priority, Descriptionget-
$Array = @()

foreach ($Rule in $Rules) {
    $Array += [PSCustomObject]@{Name = $item.Name; Priority = $item.Priority; Description = $item.Description; Enabled = $item.Enabled }
}

# Use below for CSV but note that the rows need to be expanded vertically 
# to see the full descriptions, otherwise they will look truncated.

# $Array | Export-CSV $ExportPathCSV

$Array | ConvertTo-HTML -Property Name, Priority, Description, Enabled | Out-File $ExportPath

Write-Host "$Email's rules have been output to $ExportPath.`r`n"
Write-Host "Would you like to open this file now? Default is YES."
$Readhost = Read-Host "( Y / N ) "
Switch ($ReadHost) {
    Y { Invoke-Item $ExportPath }
    N { continue }
    Default { Invoke-Item $ExportPath }
}

Remove-PSSession $RemoteEx2013Session
