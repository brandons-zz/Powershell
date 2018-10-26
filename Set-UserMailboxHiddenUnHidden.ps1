<#
.SYNOPSIS
    Hide or Unhide a User Mailbox on Office365
.DESCRIPTION
    This script connects to your Hybrid Exchange Server and hides or unhides a user mailbox from the GAL.
.PARAMETER MailboxUPN
    The full email address of the user we are performing this action on.
.PARAMETER SetHidden
    Specify if you want the mailbox hidden or not
.
PARAMETER Credential
    This account MUST be an Exchange Admin!
.EXAMPLE
    C:\PS> Set-UserMailboxHiddenUnhidden user@domain.com -SetHidden True -Credential $cred
.NOTES
    Author: Brandon Steili
    Date:   10/26/2018   
#>
[CmdletBinding()]
param (
    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter the UPN of the mailbox you are setting hidden/unhidden",
        Position = 0)]
    [string]$MailboxUPN,
    [Parameter(
        Mandatory=$true, 
        HelpMessage = "Please enter the UPN of the mailbox you are setting hidden/unhidden",
        Position = 1)]
        [ValidateSet("true", "false")]
        [string]$SetHidden="false",
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

Write-Verbose "Validated Email Address. Connecting to Exchange Server"
Write-Verbose "---------"
#Import the Local Microsoft Online PowerShell Module Cmdlets and Connect to O365 Online
if($Credential -eq [System.Management.Automation.PSCredential]::Empty) {
    $Credential = Get-Credential
}

if (!($SetHidden)) {
   # Yes/No From the command line  
Write-host "Would you like to set the mailbox hidden? (Default is No)" -ForegroundColor Yellow 
$Readhost = Read-Host " ( y / n ) " 
Switch ($ReadHost) 
 { 
   Y {$SetHidden=$true} 
   N {$SetHidden=$false} 
   Default {$SetHidden=$false} 
 } 
}
$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<ExchangeServerFQDN>/PowerShell/ -Authentication Kerberos -Credential $cred
Import-PSSession $s -DisableNameChecking
Write-Verbose "Modifying the mailbox $MailboxUPN"
Write-Verbose "---------"

if ($SetHidden.ToLower() -eq "true" ) {
    Write-Verbose "Hiding the mailbox $MailboxUPN"
    Write-Verbose "---------"
    Set-RemoteMailbox -Identity $MailboxUPN -HiddenFromAddressListsEnabled $true
 } else {
    Write-Verbose "Un-Hiding the mailbox $MailboxUPN"
    Write-Verbose "---------"
    Set-RemoteMailbox -Identity $MailboxUPN -HiddenFromAddressListsEnabled $false
 }

 Write-Verbose "Exiting the Exchange Server"
 Write-Verbose "---------"
 Exit-PSSession
 Write-Verbose "Connecting to Sync Server"
 Write-Verbose "---------"
 $s = New-PSSession -ComputerName '<ADConnectSyncServerFQDN>' -Credential $cred
 Import-Module ADSync -PSSession $s
 Write-Verbose "Syncing to Azure AD"
 Write-Verbose "---------"
 Start-ADSyncSyncCycle -PolicyType Delta
 Exit-PSSession

 Write-Verbose "Done. Please allow a couple minutes for sync to complete."
 Write-Verbose "---------"
