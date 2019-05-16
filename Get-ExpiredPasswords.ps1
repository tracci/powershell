<#
	.NOTES
	===========================================================================
     Created on:   	2018-10-29
     Created by:   	Tony Racci
	 Organization: 	GitHub
	 Filename:      Get-ExpiredPasswords.ps1
	===========================================================================
	.DESCRIPTION
    Lists AD passwords currently expired.
#>

# Variables
$DC = "domaincontroller.domain.com"
$Timestamp = Get-Date -Format "yyyy-MM-dd_THHmmss"
$ExportRoot = "\\path\to\exports"
$FileName = "$Timestamp-ExpiredPWs.csv"
$Export = "$ExportRoot\$FileName"
$Creds = Get-Credential -Message "Please enter your network credentials"

# Check AD module
if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Active Directory module not found`r`n" -f red
    Write-Host "Please install the module and re-run the script." -f red
    exit
}

Get-ADUser -Server $DC -Credential $Creds -Filter * -Properties Name, PasswordExpired, PasswordLastSet |
Where-Object { $_.Enabled -eq $true -and $_.PasswordExpired -eq $true } |
Select-Object UserPrincipalName, PasswordLastSet |
Export-Csv -Path $Export -NoTypeInformation -UseCulture

# Export and prompt to open file
Write-Host "Expired password report has been exported to $Export`r`n'"
Write-Host "Would you like to open this file now? Default is YES."
$Readhost = Read-Host "( Y / N ) "
Switch ($ReadHost) { 
    Y { Invoke-Item $Export }
    N { continue }
    Default { Invoke-Item $Export }
}
