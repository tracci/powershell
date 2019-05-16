<#
	.NOTES
	===========================================================================
	 Created on:   	2018-03-07
     Created by:   	Tony Racci
	 Organization: 	GitHub
	 Filename:      Add-ExchangeContact.ps1
	===========================================================================
	.DESCRIPTION
	Creates email contact.
#>

# Variables
$ExchangeServer = "exchange.domain.com"
$OU = "OU=contacts,DC=domain,DC=com"

# Enter Exchange session
$Creds = Get-Credential -Message "Enter your network credentials."
$Params = @{
    ConfigurationName = "Microsoft.Exchange"
    ConnectionUri     = "http://$ExchangeServer/Powershell/"
    Authentication    = "Kerberos"
    Credential        = $Creds
}
$RemoteExSession = New-PSSession @Params
Import-PSSession $RemoteExSession

# Get contact info
$FirstName = Read-Host "Enter first name"
$LastName = Read-Host "Enter last name"
$ExtMail = Read-Host "Enter external email address"

# Change case for style
$LowerFirst = $FirstName.ToLower()
$LowerLast = $LastName.ToLower()
$CapFirst = (Get-Culture).TextInfo.ToTitleCase($LowerFirst)
$CapLast = (Get-Culture).TextInfo.ToTitleCase($LowerLast)

# Create contact
New-MailContact -Name "$CapLast, $CapFirst" -ExternalEmailAddress $ExtMail -OrganizationalUnit $ou
Write-Host "$CapFirst $CapLast's contact has been created."
