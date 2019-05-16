<#	
	.NOTES
	===========================================================================
	 Created on:   	2018-06-05
     Created by:   	Tony Racci
	 Organization: 	GitHub
	 Filename:      Add-AzureGuest.ps1 	
	===========================================================================
	.DESCRIPTION
    Adds user as guest in Azure AD and sends invitation email.
#>

# Variables
$Tenant = "your_O365_tenant"
$RedirectURL = "https://$tenant-my.sharepoint.com/"

# Azure AD module check
if (!(Get-Module -ListAvailable -Name AzureAD)) {
    Write-Host "AzureAD module not found`r`n" -f Red
    Write-Host "Please install the module and re-run the script." -f red
    exit
}
$AADCred = Get-Credential -Message "Enter your Office 365 credentials. Use email address for username"
Connect-AzureAD -Credential $AADCred

# Get email address and send invitation
$Email = Read-Host "Enter guest user's email address"
Write-Host "`r`nAdding $Email as an Azure AD guest user`r`n" -f blue
Write-Host "$Email will be receiving an invitation email. Proceed? Default is NO."
$Readhost = Read-Host "( Y / N ) "
Switch ($ReadHost) { 
    Y {
        New-AzureADMSInvitation -InvitedUserEmailAddress $Email -SendInvitationMessage $True -InviteRedirectURL $RedirectURL
        Write-Host "Guest invitation has been sent to $Email" -f green
    }
    N { Write-Host "`r`nInvitation has not been sent.`r`n" -f red }
    Default { Write-Host "`r`nInvitation has not been sent.`r`n" -f red } 
}

Disconnect-AzureAD
