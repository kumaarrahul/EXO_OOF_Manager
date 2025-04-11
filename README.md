# Exchange Online OOF Manager

A streamlined PowerShell script for managing Out of Office (OOF) settings in Exchange Online. This tool allows administrators to easily backup existing OOF configurations and deploy new OOF messages in bulk for multiple users.

## Features

- **Backup OOF Settings**: Export current OOF configurations for all users to a CSV file
- **Deploy OOF Messages**: Set new OOF messages for multiple users at once
- **Flexible Scheduling**: Configure scheduled, enabled, or disabled OOF states
- **Message Templates**: Read OOF messages from text files for easy standardization
- **Detailed Reporting**: Generate comprehensive CSV reports of operations

## Prerequisites

- Exchange Online PowerShell module installed
  ```powershell
  Install-Module -Name ExchangeOnlineManagement
  ```
- Appropriate permissions to manage mailboxes (at least Exchange Administrator role)
- CSV file with user list (containing 'UserPrincipalName' or 'Email' column)

## Installation

1. Clone this repository or download the script
2. Make sure you have the required prerequisites installed
3. Prepare your user list CSV file

## Usage

### Backup OOF Settings

```powershell
.\EXO_OOF_Manager.ps1 -Action Backup -UserListPath UserList.csv
```

This will:
- Connect to Exchange Online
- Retrieve OOF settings for all users in the CSV
- Export the settings to a timestamped CSV file

### Deploy OOF Messages

```powershell
.\EXO_OOF_Manager.ps1 -Action SetOOF -UserListPath UserList.csv
```

You can also specify message file paths:

```powershell
.\EXO_OOF_Manager.ps1 -Action SetOOF -UserListPath UserList.csv -InternalMessageFile "OOF_Internal.txt" -ExternalMessageFile "OOF_External.txt"
```

The script will:
- Connect to Exchange Online
- Prompt for OOF configuration (state, dates if scheduled)
- Load messages from files or prompt for input
- Apply settings to all users in the CSV
- Generate a results report

## CSV File Format

The user list CSV should include at least one of these columns:
- `UserPrincipalName` (preferred)
- `Email`

Example:
```
UserPrincipalName
user1@contoso.com
user2@contoso.com
```

## OOF Message Files

Create two text files in the same directory as the script:
- `OOF_Internal.txt` - Contains the message for internal recipients
- `OOF_External.txt` - Contains the message for external recipients

HTML formatting is supported in these files for formatted OOF messages.

## Output Files

The script generates output files in the specified output folder (default: `.\OOF_Output`):
- Backup files: `OOF_Backup_[timestamp].csv`
- Results files: `OOF_Results_[timestamp].csv`

## Examples

### Enable OOF for All Users

```powershell
.\EXO_OOF_Manager.ps1 -Action SetOOF -UserListPath UserList.csv
```
Then select "Enabled" when prompted for OOF state.

### Schedule OOF for a Future Date

```powershell
.\EXO_OOF_Manager.ps1 -Action SetOOF -UserListPath UserList.csv
```
Then select "Scheduled" when prompted and enter your desired date range.

### Disable OOF for All Users

```powershell
.\EXO_OOF_Manager.ps1 -Action SetOOF -UserListPath UserList.csv
```
Then select "Disabled" when prompted for OOF state.

## Notes

- The script automatically handles time zone information for scheduled OOF
- When using "Scheduled" mode, date/time values use your local system time zone
- Always review the generated CSV reports to verify successful operations

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Author

Rahul Kumaar