# Exchange Online

PowerShell scripts for Exchange Online administration and reporting.

## Get-ExchangeMailboxInventory.ps1

### Description

Retrieves mailbox information from Exchange Online and exports
the inventory to a CSV file.

### Requirements

- PowerShell 7 or later.
- ExchangeOnlineManagement module.
- Active Exchange Online connection.

### Usage

Connect to Exchange Online:

```powershell
Connect-ExchangeOnline

Run the script:
.\Get-ExchangeMailboxInventory.ps1

Specify an output file:
.\Get-ExchangeMailboxInventory.ps1
```

### Output
A CSV file containing:
- DisplayName
- UserPrincipalName
- PrimarySmtpAddress
- RecipientTypeDetails
