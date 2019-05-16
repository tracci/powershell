<#
	.NOTES
	===========================================================================
	 Created on:   	2019-04-15
     Created by:   	Tony Racci
	 Organization: 	GitHub
     Filename:      Get-AccountExpirations.ps1
	===========================================================================
	.DESCRIPTION
    Finds accounts that either have no expiration date set
    or have already expired

    This is written to run scheduled from a secure system allowed to send SMTP
    without authentication. Adjust as needed.
#>

# Check for AD module
if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Output "Active Directory module not found. Exiting."
    exit
}

# Variables
$DC = "domaincontroller.domain.com"
$SMTPServer = "mail.domain.com"
$Today = Get-Date -Format "yyyy-MM-dd"
$LogPath = "\\path\to\transcript\logs"
$OU = "OU=Users,DC=domain,DC=com"

# Turn on logging
Get-ChildItem "$LogPath\*.log" | Where-Object LastWriteTime -LT (Get-Date).AddDays(-30) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'yyyy-MM-ddTHHmmss').log"
Start-Transcript $LogPathName

# Get user account expiration info
$Users = Get-ADUser -Server $DC -SearchBase $OU -Filter * -Properties AccountExpirationDate, pwdLastSet
foreach ($User in $Users) {
    if ($null -eq $User.AccountExpirationDate -and $User.pwdLastSet -ne 0) {
        $NoExp += "$($User.Name)<br/>"
    }
    if ($User.AccountExpirationDate -LT (Get-Date) -and $User.pwdLastSet -ne 0) {
        $Exp += "$($User.Name)<br/>"
    }
}

# Set email titles and send message if any variables are populated
if ($null -ne $Exp) {
    $ExpTitle = "Accounts that have expired:<br/>"
}
if ($null -ne $NoExp) {
    $NoExpTitle = "Accounts without an expiration date:<br/>"
}
if (!($null -eq $Exp -and $null -eq $NoExp)) {
    $From = "noreply@domain.com"
    $To = "alerts@domain.com"
    $Subject = "Employee AD report: $today"
    $Body = "<b>Report of users in the organization:</b><br/><br/>

        $ExpTitle <font color=""blue"">$Exp</font><br/>

        $NoExpTitle <font color=""blue"">$NoExp</font><br/>"

    $SMTPPort = "25"
    $MailMessage = @{
        To         = $To
        From       = $From
        Subject    = $Subject
        Body       = $Body
        SmtpServer = $SMTPServer
        Port       = $SMTPPort
        BodyAsHtml = $true
    }
    Send-MailMessage @MailMessage
}

Stop-Transcript
