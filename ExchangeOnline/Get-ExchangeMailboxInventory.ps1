
<#
.SYNOPSIS
    Exports an Exchange Online mailbox inventory.

.DESCRIPTION
    Retrieves mailbox information from Exchange Online
    and exports the results to a CSV file.

.NOTES
    Author: KikeCorp
    Website: https://kikecorp.eu/
    Requires: ExchangeOnlineManagement module
    PowerShell: 7 or later

.EXAMPLE
    .\Get-ExchangeMailboxInventory.ps1

.EXAMPLE
    .\Get-ExchangeMailboxInventory.ps1 -OutputPath "C:\Reports\Mailboxes.csv"
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$OutputPath = ".\MailboxInventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Check required module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "The ExchangeOnlineManagement module is not installed."
}

# Check Exchange Online connection
$connection = Get-ConnectionInformation -ErrorAction SilentlyContinue

if (-not $connection) {
    throw "No active Exchange Online connection. Run Connect-ExchangeOnline first."
}

try {
    Write-Host "Retrieving mailbox inventory..." -ForegroundColor Cyan

    $mailboxes = Get-EXOMailbox -ResultSize Unlimited |
        Select-Object DisplayName,
                      UserPrincipalName,
                      PrimarySmtpAddress,
                      RecipientTypeDetails

    $mailboxes | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Host "Export completed successfully." -ForegroundColor Green
    Write-Host "File: $OutputPath"
    Write-Host "Mailboxes: $($mailboxes.Count)"
}
catch {
    Write-Error "Unable to retrieve mailbox inventory: $_"
    throw
}
