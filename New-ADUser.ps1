<#
	.NOTES
	===========================================================================
	 Created on:   	2018-04-04
     Created by:   	Tony Racci
	 Organization: 	GitHub
	 Filename:      New-ADUser.ps1
	===========================================================================
	.DESCRIPTION
    Creates new AD user account from data entered manually
    Adds to base security and distribution groups
    Adds Logon script
    Assigns O365 license
    Sends confirmation email to alert address
#>

# Check for MSO and AD modules
$Modules = @("MSOnline", "ActiveDirectory")
foreach ($Module in $Modules) {
    if (!(Get-Module -ListAvailable -Name $Module)) {
        Write-Host "$Module not found. Please install and re-run the script." -f red
        exit
    }
}

# Variables /  AD & Exchange
$DC = "domaincontroller.domain.com"         # Primary DC
$ExchangeServer = "exchange.domain.com"     # Exchange server
$MailboxDB = "Mailbox Database name"        # Mailbox database
$ScriptPath = "LogonScript.bat"             # New user logon script field
$Company = "Company, LLC"                   # Company
$TestGroup = "Domain Guests"                # Built-in group to validate AD credentials against
$SecGroup = "Default security group"        # Add default security group addition here
$DistGroup = "Default distribution group"   # Add default distribution group addition here
$Creds = Get-Credential -Message "Please enter your network credentials"
$Date = Get-Date -format yyyy-MM-dd
$LogPath = "\\path\to\logs"
$AlertEmail = "alerts@domain.com"

# Variables / O365 and Azure AD Connect
# O365 username is formatted as username@domain.com from credentials above
$AADServer = "AADC.domain.com"              # Azure AD Connect server
$DomainUser = $Creds.UserName
$DomainPass = $Creds.GetNetworkCredential().SecurePassword
$O365U = $DomainUser + "@domain.com"
$LiveCred = New-Object System.Management.Automation.PSCredential ($O365U, $DomainPass)
$Tenant = "O365 Tenant"
$SKU = $Tenant + ":ENTERPRISEPACK"          # Sets E3 license
$ISO3166 = "US"                             # Licenses for US. Adjust as needed according to ISO3166 nomenclature.

# Validate AD credentials
try {
    Get-ADGroup -Credential $creds -Server $DC -Identity $testgroup | Out-Null
}
catch [System.Security.Authentication.AuthenticationException] {
    Write-Host "Invalid credentials. Re-run the script." -f Red
    exit
}

# Turn on logging
Get-ChildItem "$LogPath\*.log" | Where-Object LastWriteTime -LT (Get-Date).AddDays(-30) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'yyyy-MM-ddTHHmmss').log"
Start-Transcript $LogPathName

# Create remote sessions to Exchange and O365
$ExcParams = @{
    ConfigurationName = "Microsoft.Exchange"
    ConnectionUri     = "http://$exchangeserver/Powershell/"
    Authentication    = "Kerberos"
    Credential        = $creds
}
Try {
    $RemoteEx2013Session = New-PSSession @ExcParams -ErrorAction Stop
}
Catch { 
    $ErrMsg = $_.Exception.Message
    Write-Output "Failed to connect to Exchange. Ending script."
    Write-Output "Error: $ErrMsg"
    exit
}

$EOparams = @{
    ConfigurationName = "Microsoft.Exchange"
    ConnectionUri     = "https://outlook.office365.com/powershell-liveid/"
    Authentication    = "Basic"
    Credential        = $LiveCred
}
Try {
    $EOSession = New-PSSession @EOparams -ErrorAction Stop
}
Catch { 
    $ErrMsg = $_.Exception.Message
    Write-Output "Failed to connect to Exchange Online. Ending script."
    Write-Output "Error: $ErrMsg"
    exit
}

Try { 
    $AADSession = New-PSSession -ComputerName $AADServer -ErrorAction Stop
}
Catch { 
    $ErrMsg = $_.Exception.Message
    Write-Output "Failed to connect to Azure AD Connect. Ending script."
    Write-Output "Error: $ErrMsg"
    exit
}

Import-PSSession $EOSession -DisableNameChecking -AllowClobber
Import-PSSession $RemoteEx2013Session -AllowClobber
Import-Module ActiveDirectory
Import-Module MSOnline

##########################################
##  Map variables and data from         ##
##  input                               ##
##########################################

$First = Read-Host "Enter new employee first name"
$Last = Read-Host "Enter new employee last name"
$LowerFirst = $First.ToLower()
$CapFirst = (Get-Culture).TextInfo.ToTitleCase($LowerFirst)
$LowerLast = $Last.ToLower()
$CapLast = (Get-Culture).TextInfo.ToTitleCase($LowerLast)
$Username = $LowerFirst + '.' + $LowerLast                          # Username format as firstname.lastname
$Username = $Username -replace '\s', ''
$ValidUser = Get-ADUser -LDAPFilter "(SAMAccountName=$Username)"    # If name is taken, change to firstinitial.lastname
If ($null -ne $ValidUser) {
    Write-Host "$Username already exists. Changing username." -ForegroundColor red
    $Firstinit = $LowerFirst[0]
    $Username = $Firstinit + '.' + $LowerLast
}
$Displayname = $CapLast + ', ' + $CapFirst                          # Sets display name as Last, First.
$Department = Read-Host "Enter department"
$JobTitle = Read-Host "Enter job title"
$Description = $JobTitle
do {
    $SupervisorMail = Read-Host "Enter supervisor's email address"
    $ValidSupMail = (Get-ADUser -Filter "EmailAddress -like '$SupervisorMail'")
}
until ($null -ne $ValidSupMail)
$Supervisor = (Get-ADUser -Filter "EmailAddress -eq '$SupervisorMail'" -Properties SAMAccountName, Department, DistinguishedName)
$SupervisorName = $Supervisor.SamAccountName
$SupervisorDN = $Supervisor.DistinguishedName
$ITEmail = $Username + "@domain.com"
$SecurePW = Read-Host "Enter initial password" -AsSecureString
$OU = (([ADSI]"LDAP://$SupervisorDN").parent).substring(7)           # Places new user in same OU as their manager. 

$variables = @{

    Enabled               = $True
    Server                = $DC
    AccountPassword       = $SecurePW
    Department            = $Department
    Company               = $Company
    Manager               = $SupervisorName
    Country               = $ISO3166
    Description           = $Description
    EmployeeID            = $EmpNumber
    DisplayName           = $Displayname
    Path                  = $OU
    Title                 = $JobTitle
    SamAccountName        = $Username
    UserPrincipalName     = $ITEmail
    GivenName             = $CapFirst
    Surname               = $CapLast
    ChangePasswordAtLogon = $True
    ScriptPath            = $ScriptPath
}

##########################################
##  Confirm, then write data            ##
##  to Active Directory                 ##
##########################################

Write-Host "`r`n===== Data Entered =====`r`n"
Write-Host "First Name: " -NoNewLine; Write-Host "$CapFirst" -f blue
Write-Host "Last Name: " -NoNewLine; Write-Host "$CapLast" -f blue
Write-Host "User Name: " -NoNewLine; Write-Host "$Username" -f blue
Write-Host "Job Title: " -NoNewLine; Write-Host "$JobTitle" -f blue
Write-Host "Description: " -NoNewLine; Write-Host "$Description" -f blue
Write-Host "Department: " -NoNewLine; Write-Host "$Department" -f blue
Write-Host "Country: " -NoNewLine; Write-Host "$ISO3166" -f blue
Write-Host "Manager: " -NoNewLine; Write-Host "$SupervisorName" -f blue
Write-Host "Company: " -NoNewLine; Write-Host "$Company" -f blue
Write-Host "Display Name: " -NoNewLine; Write-Host "$Displayname" -f blue
Write-Host "OU: " -NoNewline; Write-Host "$OU" -f blue
Write-Host "O365 License: " -NoNewLine; Write-Host "$SKU" -f blue

$message = "`r`nWriting to AD"
$question = "Are you sure you want to proceed?"
$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList "&Yes", "Yes, changes will be written to AD"))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList "&No", "No, changes will not be written to AD and script will end"))
$decision = $Host.UI.PromptForChoice($message, $question, $choices, 0)
if ($decision -eq 0) {
    Write-Host "Confirmed, creating new user in AD.`r`n"
    New-ADUser $Username @variables
}
else {
    Write-Host 'Operation cancelled.'
    break
}

##########################################
##  Enable mailbox                      ##
##  Select on-prem or O365              ##
##########################################

# Enable on-prem mailbox and set default calendar permissions to 'Reviewer'

Enable-Mailbox -Identity $Username -Database $MailboxDB -DomainController $DC
Write-Host "Waiting for mailbox to register." -f Cyan
$MailboxTimeout = $null
$ExchTimer = Get-Date
do {
    Start-Sleep -s 3
    $MailboxTimeout = Get-MailboxFolderStatistics -identity $Username -FolderScope Calendar -ErrorAction SilentlyContinue
}
while ($null -eq $MailboxTimeout -and $ExchTimer.AddMinutes(5) -gt (Get-Date))
Write-Host "`r`nMailbox registered, setting calendar permissions.`r`n" -f Green
Set-MailboxFolderPermission -Identity ${username}:\calendar -User Default -AccessRights Reviewer -DomainController $DC

# Enable O365 mailbox

#Enable-RemoteMailbox $Username -RemoteRoutingAddress $Username@$tenant.onmicrosoft.com -DomainController $DC

##########################################
##  Remotely execute Azure AD Sync      ##
##  and wait for completion.            ##
##########################################

Invoke-Command -Session $AADSession -ScriptBlock { Import-Module -Name 'ADSync' }
Invoke-Command -Session $AADSession -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }
Invoke-Command -Session $AADSession -ScriptBlock { $null = $ADSyncSched }

Write-Host "Waiting for Azure AD Sync to complete." -f Cyan
$aadtime = Get-Date
do {
    Start-Sleep -s 10
    Invoke-Command -Session $AADSession -ScriptBlock { $ADSyncSched = Get-ADSyncScheduler }
    Invoke-Command -Session $AADSession -ScriptBlock { $ADSyncStatus = $ADSyncSched.SyncCycleInProgress }
}
while ($ADSyncStatus -eq $true -and $aadtime.AddMinutes(10) -gt (Get-Date))
Write-Host "`r`nAzure AD Sync complete.`r`n" -f Green
Remove-Variable aadtime

Connect-MsolService -Credential $LiveCred
$O365timeout = $null
Write-Host "Waiting for $ITEmail to be available in O365." -f Cyan
Write-Host "This can take a minute or two." -f Cyan
$o365time = Get-Date
do {
    Start-Sleep -s 15
    $O365timeout = Get-MsolUser -UserPrincipalName $ITEmail -ErrorAction SilentlyContinue
}
while ($null -eq $O365timeout -and $o365time.AddMinutes(15) -gt (Get-Date))
Remove-Variable o365time
Start-Sleep -s 30
Write-Host "`r`n$ITEmail available, assigning O365 license.`r`n" -f Green

##########################################
##  Assign O365 license                 ##
##  Add to groups                       ##
##########################################

Set-MsolUser -UserPrincipalName $ITEmail -UsageLocation $ISO3166
Set-MsolUserLicense -UserPrincipalName $ITEmail -AddLicenses $SKU
Add-ADGroupMember -Identity $SecGroup -Members $Username -Server $DC
Add-DistributionGroupMember -Identity $DistGroup -Member $Username -BypassSecurityGroupManagerCheck

Start-Sleep -s 5

##########################################
##  Confirm AD info, rename object,     ##
##  and send email with info            ##
##########################################

$UserDetails = (Get-ADUser $Username -Properties * -Server $DC |
    Select-Object SamAccountName, Created, EmailAddress, DisplayName, EmployeeID, Title, Department, Country, Manager, CanonicalName, DistinguishedName, MemberOf)
$O365Details = (Get-MSOLuser -UserPrincipalName $ITEmail |
    Select-Object IsLicensed, Licenses)

$ADSam = $UserDetails.SamAccountName
$ADCreated = $UserDetails.Created
$ADEmail = $UserDetails.EmailAddress
$ADDisplay = $UserDetails.DisplayName
$ADEmpID = $UserDetails.EmployeeID
$ADTitle = $UserDetails.Title
$ADDept = $UserDetails.Department
$ADCountry = $UserDetails.Country
$ADManager = $UserDetails.Manager
$ADManagerSAM = (Get-ADUser $ADManager -Properties Manager).SamAccountName
$ADCanon = $UserDetails.CanonicalName
$ADDN = $UserDetails.DistinguishedName
$O365Lic = $O365Details.IsLicensed
$O365LicDet = $O365Details.Licenses.AccountSkuID
$O365LicSumm = (Get-MsolAccountSku | Where-Object { $_.AccountSkuID -eq $SKU })
$O365LicAvail = ($O365LicSumm.ActiveUnits - $O365LicSumm.ConsumedUnits)

## Get group membership and clean up for the output.
$MemberList = @()
$ListGroups = Get-ADUser -Identity $ADSam -Properties * | Select-Object -ExpandProperty MemberOf
foreach ($i in $ListGroups) {
    $i = ($i -split ',')[0]
    $MemberList += ($i -creplace 'CN=|}', '')
}

Rename-ADObject -Identity $ADDN -NewName $ADDisplay -Server $DC

$From = [adsisearcher]"(samaccountname=$env:USERNAME)"
$To = $AlertEmail
$Subject = "Account created: $Username"
$Body = "<b>Active Directory configuration</b>:<br/><br/>

        Username: <font color=""blue"">$ADSam</font><br/>
        Created: <font color=""blue"">$ADCreated</font><br/>
        Email: <font color=""blue"">$ADEmail</font><br/>
        Display: <font color=""blue"">$ADDisplay</font><br/>
        Employee ID: <font color=""blue"">$ADEmpID</font><br/>
        Title: <font color=""blue"">$ADTitle</font><br/>
        Department: <font color=""blue"">$ADDept</font><br/>
        Country: <font color=""blue"">$ADCountry</font><br/>
        Manager: <font color=""blue"">$ADManagerSAM</font><br/>
        Canonical Name: <font color=""blue"">$ADCanon</font><br/>

        Office 365 license assigned: <font color=""blue"">$O365Lic</font><br/>
        Office 365 licenses: <font color=""blue"">$O365LicDet</font><br/>
        Office 365 licenses available: <font color=""blue"">$O365LicAvail</font><br/>
        
        <br/>Added to these groups: <font color=""blue"">$MemberList</font><br/>"

$SMTPServer = $ExchangeServer
$SMTPPort = "587"

$MailMessage = @{
    To         = $To
    From       = $From.FindOne().Properties.mail
    Subject    = $Subject
    Body       = $Body
    SmtpServer = $SMTPServer
    Port       = $SMTPPort
    Credential = $creds
    BodyAsHtml = $true
    UseSsl     = $true
}
Send-MailMessage @MailMessage

Write-Output "$Username created $Date"

Remove-PSSession $EOSession
Remove-PSSession $RemoteEx2013Session
Remove-PSSession $AADSession

Stop-Transcript
