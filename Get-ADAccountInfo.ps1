<#
	.NOTES
	===========================================================================
     Created on:   	2018-03-08
     Created by:   	Tony Racci
	 Organization: 	GitHub
	 Filename:      Get-ADAccountInfo
	===========================================================================
	.DESCRIPTION
    Checks AD account and password expiration status.
    Unlocks account if requested.
#>

# Import AD module
if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Active Directory module not found`r`n" -f red
    Write-Host "Please install the module and re-run the script." -f red
    exit
}
Import-Module ActiveDirectory

# Variables
$DC = "domaincontroller.domain.com"
$Username = Read-Host -Prompt 'Enter username of account to check'
$ValidUser = Get-ADUser -Server $DC -LDAPFilter "(SamAccountName=$Username)" -Properties *, "msDS-UserPasswordExpiryTimeComputed"

# Validate username and gather AD info
if ($null -ne $ValidUser) {
    if ($null -eq $ValidUser.PasswordLastSet) {
        $PwdChange = $true
        $PwdLastSet = "Change at next login."
    }
    else {
        $PwdChange = $false
        $PwdLastSet = Get-Date -Date $ValidUser.PasswordLastSet
    }
    $BadPwCount = $ValidUser.BadLogonCount
    $LastBadPwTime = $ValidUser.LastBadPasswordAttempt
    $Lockout = $ValidUser.LockedOut
    $IsExpired = $ValidUser.PasswordExpired
    $AcctExpire = $ValidUser.AccountExpirationDate
    $NeverExpires = $ValidUser.PasswordNeverExpires
    if ($NeverExpires -eq $false) {
        $ExpireTime = $ValidUser."msDS-UserPasswordExpiryTimeComputed"
        $FormatExpire = Get-Date -Date ([DateTime]::FromFileTime([Int64]::Parse($ExpireTime))) -Format "MM/dd/yyyy HH:mm:ss"
    }
    else { 
        $FormatExpire = "Account never expires."
    }
    
    # Lockout check
    Write-Host "========== Details for $Username ==========`r`n"
    if ($Lockout -eq $true) {
        Write-Host "Account IS currently locked out." -ForegroundColor Red
    }
    else { Write-Host "Account IS NOT currently locked out." -ForegroundColor Green }

    # Password expiration check
    if ($NeverExpires -eq $true) { 
        Write-Host "Password is set to never expire.`r`n" -ForegroundColor Blue 
    }
    elseif ($PwdChange -eq $true) { 
        Write-Host "Password is set to be changed at next login.`r`n" -ForegroundColor Red 
    }
    elseif ($IsExpired -eq $true) { 
        Write-Host "Password IS currently expired.`r`n" -ForegroundColor Red 
    }
    else { Write-Host "Password HAS NOT expired.`r`n" -ForegroundColor Green }

    # Account expiration check
    if ($null -ne $AcctExpire) { 
        Write-Host "User account expiration:            $AcctExpire" -ForegroundColor Blue 
    }

    # Output info
    Write-Host "Password last set:                  $PwdLastSet"
    Write-Host "Password expiration:                $FormatExpire`r`n"
    Write-Host "Number of bad password attempts:    $BadPwCount"
    Write-Host "Time of last bad password attempt:  $LastBadPwTime`r`n"

    # Ask to unlock a locked account
    if ($Lockout -eq $true) {
        Write-Host "Would you like to unlock the account? Default is NO."
        $Readhost = Read-Host "( Y / N ) "
        Switch ($ReadHost) {
            Y {
                Unlock-ADAccount -Identity $Username -Server $DC 
                Write-Host "$Username's account has been unlocked"
            }
            N { Write-Host "`r`n$Username's account has not been unlocked" }
            Default { Write-Host "`r`n$Username's account has not been unlocked" }
        }
    }
}
else { "`r`n$Username was not found in AD`r`n" }
