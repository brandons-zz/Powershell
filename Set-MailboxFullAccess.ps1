<#
.SYNOPSIS
    Give access to an O365 Mailbox to single user
.DESCRIPTION
    This script connects to Office 365 and gives a user full access rights to another user's mailbox. Requires MSOnline Powershell for Active Directory
    https://www.powershellgallery.com/packages/MSOnline/1.1.183.17
.PARAMETER MailboxUPN
    The mailbox we want to give someone access to.
.PARAMETER AccessUPN
    The user we are giving mailbox access to. 
.EXAMPLE
    C:\PS> Set-MailboxFullAccess user@domain.com user2@domain.com -Credential $cred
.NOTES
    Author: Brandon Steili
    Date:   10/10/2018   
#>
[CmdletBinding()]
param (
    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter the UPN of the mailbox you are giving access to",
        Position = 0)]
    [string]$MailboxUPN,
    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter the UPN of the user you are giving access to",
        Position = 1)]
    [string]$AccessUPN,
    [Parameter()]
    [ValidateNotNull()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential = [System.Management.Automation.PSCredential]::Empty 
)

function IsValidEmail { 
    param([string]$EmailAddress)
    Write-Verbose "Validating $EmailAddress is an email address"
    Write-Verbose "---------"
    try {
        $null = [mailaddress]$EmailAddress
        return $true
    }
    catch {
        return $false
    }
}

Write-Verbose "---------"
Write-Verbose "---------"

if (!(IsValidEmail $MailboxUPN)) {
    do { 
        $MailboxUPN = Read-Host -Prompt "Entry appears invalid. Please enter the UPN of the MAILBOX you are giving access to" 
    } 
    until (IsValidEmail $MailboxUPN)
}

if (!(IsValidEmail $AccessUPN)) {
    do { 
        $AccessUPN = Read-Host -Prompt "Entry appears invalid. Please enter the UPN of the USER you are giving access to" 
    } 
    until (IsValidEmail $AccessUPN)
}

Write-Verbose "Validated Email Addresses. Connecting to O365"
Write-Verbose "---------"
#Import the Local Microsoft Online PowerShell Module Cmdlets and Connect to O365 Online
if($Credential -eq [System.Management.Automation.PSCredential]::Empty) {
    $Credential = Get-Credential
}
Import-Module MSOnline
Connect-MsolService -Credential $Credential
#Establish an Remote PowerShell Session to Exchange Online
$msoExchangeURL = "https://ps.outlook.com/powershell/"
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $msoExchangeURL -Credential $Credential -Authentication Basic -AllowRedirection
Import-PSSession $session
Write-Verbose "Modifying the mailbox $MailboxUPN"
Write-Verbose "---------"
Add-MailboxPermission -Identity $MailboxUPN -User $AccessUPN -AccessRights FullAccess -InheritanceType All

# Yes/No From the command line - Send on Behalf
Write-host "Would you like to add Send on Behalf Permissions" -ForegroundColor Yellow 
    $Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost) 
     { 
       Y {$SendOnBehalf=$true} 
       N {$SendOnBehalf=$false} 
       Default {$SendOnBehalf=$false} 
     } 

IF ($SendOnBehalf) {
    Set-Mailbox $MailboxUPN -GrantSendOnBehalfTo $AccessUPN
    Write-Verbose "Adding Send On Behalf to the mailbox $MailboxUPN"
    Write-Verbose "---------"
}

# Yes/No From the command line - Send As
Write-host "Would you like to add Send As Permissions" -ForegroundColor Yellow 
    $Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost) 
     { 
       Y {$SendAs=$true} 
       N {$SendAs=$false} 
       Default {$SendAs=$false} 
     } 

IF ($SendAs) {
    Add-RecipientPermission $MailboxUPN -AccessRights SendAs -Trustee $AccessUPN -Confirm:$false 
    Write-Verbose "Adding Send As to the mailbox $MailboxUPN"
    Write-Verbose "---------"
}

Write-Verbose "Exiting Office 365"
Write-Verbose "---------"
Exit-PSSession