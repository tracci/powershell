<#	
	.NOTES
	===========================================================================
	 Created on:   	2018-05-14
     Created by:   	Tony Racci
	 Organization: 	GitHub
     Filename:      AccountTermination-O365.ps1	
	===========================================================================
	.DESCRIPTION
    Rewritten version of the termination script.
    Modified for O365 tasks.
    Original script written by asus892 at https://github.com/asus892/TerminatedScript2.
#>

# Variables
$DC = "domaincontroller.domain.com"
$ExchangeServer = "exchange.domain.com"
$Creds = Get-Credential -Message "Please enter your network credentials"
$TestGroup = "Domain Guests"        # Built-in group to validate AD credentials against
$LogPath = "\\path\to\transcript\logs"
$DeactivatedOU = "OU=Terminated,DC=domain,DC=com"
$Date = Get-Date -format yyyy-MM-dd
$DomainUser = $Creds.UserName
$Tenant = "O365_Tenant"
$License = $Tenant + ":ENTERPRISEPACK"
$Domain = "domain.com"

# Format username to email address for O365 login - assuming username is email prefix.
$DomainPass = $Creds.GetNetworkCredential().SecurePassword
$O365U = $DomainUser + $Domain
$LiveCred = New-Object System.Management.Automation.PSCredential ($O365U, $DomainPass)


# Validate AD credentials
try {
    Get-ADGroup -Credential $Creds -Server $DC -Identity $TestGroup | Out-Null
}
catch [System.Security.Authentication.AuthenticationException] {
    Write-Host "Invalid credentials. Re-run the script." -f Red
    exit
}

# Turn on logging
Get-ChildItem "$LogPath\*.log" | Where-Object LastWriteTime -LT (Get-Date).AddDays(-30) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'yyyy-MM-dd').log"
Start-Transcript $LogPathName -Append

# Get terminated username and validate
$UserInput = Read-Host "Enter username to be deactivated"
$ADcheck = Get-ADUser -Server $dc -Filter { SamAccountName -eq $UserInput }

# Confirm before committing
If ($ADcheck) {
    Write-Host "`r`n===== Confirm Action =====`r`n"
    Write-Host "User to be deactivated: " -NoNewLine; Write-Host "$UserInput" -f blue
    $Message = "`r`nDeactivating AD account."
    $Question = "Are you sure you want to proceed?"
    $Choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
    $Choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList "&Yes", "Yes, changes will be written to AD"))
    $Choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList "&No", "No, changes will not be written to AD and script will end"))
    $Decision = $Host.UI.PromptForChoice($Message, $Question, $Choices, 0)
    If ($Decision -eq 0) {
        Write-Host "Confirmed, deactivating account.`r`n"
            
        ##########################################
        ##  Active Directory                    ##
        ##  actions                             ##
        ##########################################

        # Get user manager
        $Manager = (Get-ADUser (Get-ADUser $UserInput -Properties Manager).manager -Properties SamAccountName).SamAccountName

        # Hide from GAL
        # $UserInput.msExchHideFromAddressLists = $true

        # Deactivate account
        Disable-ADAccount -Identity $UserInput -Server $DC
        $Disabled = $UserInput + " has been deactivated"
 
        # Get AD group memberships
        $User = $UserInput
        $List = @()
        $Groups = Get-ADUser -Identity $User -Server $DC -Properties * | Select-Object -ExpandProperty MemberOf
        ForEach ($Group in $Groups) {
            $Group = ($Group -split ',')[0]
            $List += "<br/>" + ($Group -creplace 'CN=|}', '')
        }
    
        # Clear groups (minus 'Domain Users')
        (Get-ADuser $UserInput -Server $DC -Properties MemberOf).MemberOf | Remove-ADGroupMember -Member $UserInput -Server $DC -Confirm:$False

        # Change user attributes accordingly
        $ADParams = @{
            Identity    = $UserInput
            Server      = $DC
            Title       = $null
            Company     = $null
            Manager     = $null
            Department  = $null
            Description = "Account deactivated $Date"
        }
        Set-ADUser @ADParams

        # Reset password
        $NewPwd = ConvertTo-SecureString -String "T3Rm^&#12E4" -AsPlainText -Force
        Set-ADAccountPassword $UserInput -Server $DC -NewPassword $NewPwd -Reset

        # Move to deactivated OU
        Get-ADUser -Server $DC -Filter { SamAccountName -like $UserInput } | Move-ADObject -Server $DC -TargetPath $DeactivatedOU

        ##########################################
        ##  Office 365 actions                  ##
        ##                                      ##
        ##########################################

        Connect-MsolService -Credential $LiveCred
        $EOLparams = @{
            ConfigurationName = "Microsoft.Exchange"
            ConnectionUri     = "https://outlook.office365.com/powershell-liveid/"
            Authentication    = "Basic"
            Credential        = $LiveCred
            AllowRedirection  = $true
        }
        $EOLSession = New-PSSession @EOLparams
        Import-PSSession $EOLSession -AllowClobber

        # Give manager full access to calendar
        $CalName = (Get-MailboxFolderStatistics -Identity $UserInput | Where-Object { $_.FolderType -eq "Calendar" }).Name
        $Cal = $UserInput + ":\$CalName"
        Add-MailboxFolderPermission -Identity $Cal -User $Manager -AccessRights Owner

        # Convert to shared mailbox
        Set-Mailbox $UserInput -Type Shared

        # Remove O365 license (E3)
        Set-MsolUserLicense -UserPrincipalName $UserInput$Domain -RemoveLicenses $License

        Remove-PSsession $EOLSession
        Start-Sleep -s 2

        ##########################################
        ##  Compose email                       ##
        ##  confirmation                        ##
        ##########################################

        $CurrentUser = [adsisearcher]"(SamAccountName=$env:USERNAME)"
        $MailMessage = @{
            To         = "alert.email@domain.com"
            From       = $currentuser.FindOne().Properties.mail
            Subject    = "Account deactivated via script: $UserInput"
            BodyAsHtml = $true
            Body       = "<b>Active Directory changes</b>:<br/><br/>
                         $Disabled<br/>
                         AD password has been changed<br/>
                         AD description changed to:  Account Deactivated $Date<br/>
                         AD title, department, company, and manager have all been cleared<br/>
                         Account moved to 'Terminated' OU in AD<br/>
                         Account was removed from the following AD groups:<br/>
                         -----<br/>
                         $List<br/>
                         -----<br/><br/>
                         <b>O365 changes</b>:<br/><br/>
                         $UserInput's mailbox has been converted to shared<br/>
                         O365 license has been removed<br/>"
            SMTPServer = $ExchangeServer
            Port       = "587"
            Credential = $Creds
            UseSsl     = $true
        }
        Send-MailMessage @MailMessage
        Write-Output "$UserInput deactivated $Date"
    }
    Else {
        Write-Host 'Operation cancelled.'
        Break
    }
}
Else {
    Write-Host "`nScript failed! $UserInput is not in Active Directory." -ForegroundColor White -BackgroundColor Red
}

Stop-Transcript
