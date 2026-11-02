# Exchange Online

PowerShell scripts for Exchange Online administration and reporting.

## Scripts available

### [Get-ExchangeMailboxInventory.ps1](./Get-ExchangeMailboxInventory.ps1)

Retrieves mailbox information from Exchange Online and exports the inventory to a CSV file.

#### Requirements

- PowerShell 7 or later.
- ExchangeOnlineManagement module.
- Active Exchange Online connection.

#### Usage

Connect to Exchange Online:

```powershell
Connect-ExchangeOnline
```

Run the script:

```powershell
.\Get-ExchangeMailboxInventory.ps1
```

Specify an output file:

```powershell
.\Get-ExchangeMailboxInventory.ps1 -OutputPath "C:\Reports\Mailboxes.csv"
```

#### Output

A CSV file containing:

- DisplayName
- UserPrincipalName
- PrimarySmtpAddress
- RecipientTypeDetails
