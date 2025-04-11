<#
# ===================================================================================================
# Script Name: Exchange Online OOF Manager
# Created: Apr 2025
# Author: Rahul Kumaar
# Version: 2.0
# 
# Synopsis:
#   A streamlined PowerShell script for managing Out of Office (OOF) settings in Exchange Online.
#   Handles backing up existing settings and deploying new OOF messages.
#
# Prerequisites:
#   - Exchange Online PowerShell module installed (Install-Module -Name ExchangeOnlineManagement)
#   - Appropriate permissions to manage mailboxes
#   - CSV file with user list (must contain a column named 'UserPrincipalName' or 'Email')
#
# Usage:
#   .\OOF-Manager.ps1 -Action Backup -UserListPath UserList.csv
#   .\OOF-Manager.ps1 -Action SetOOF -UserListPath UserList.csv -InternalMessageFile "OOF_Internal.txt" -ExternalMessageFile "OOF_External.txt"
# ===================================================================================================
#>

param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("Backup", "SetOOF")]
    [string]$Action,
    
    [string]$UserListPath = ".\UserList.csv",
    
    [string]$OutputFolder = ".\OOF_Output",
    
    [string]$InternalMessageFile = ".\OOF_Internal.txt",
    
    [string]$ExternalMessageFile = ".\OOF_External.txt"
)

# Create output folder if it doesn't exist
if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory | Out-Null
    Write-Host "Created output folder: $OutputFolder" -ForegroundColor Green
}

# Function to validate user list file
function Test-UserListFile {
    if (-not (Test-Path -Path $UserListPath)) {
        Write-Host "Error: User list file not found at path: $UserListPath" -ForegroundColor Red
        return $false
    }
    
    try {
        $users = Import-Csv -Path $UserListPath
        if ($users.Count -eq 0) {
            Write-Host "Error: User list file is empty: $UserListPath" -ForegroundColor Red
            return $false
        }
        
        # Check for required column
        $firstUser = $users[0]
        if (-not ($firstUser.PSObject.Properties.Name -contains "UserPrincipalName" -or 
                  $firstUser.PSObject.Properties.Name -contains "Email")) {
            Write-Host "Error: User list file must contain 'UserPrincipalName' or 'Email' column: $UserListPath" -ForegroundColor Red
            return $false
        }
        
        return $true
    }
    catch {
        Write-Host "Error validating user list file: $_" -ForegroundColor Red
        return $false
    }
}

# Main script execution
Write-Host "Exchange Online OOF Manager" -ForegroundColor Cyan
Write-Host "Action: $Action" -ForegroundColor Yellow
Write-Host "User List: $UserListPath" -ForegroundColor Yellow
Write-Host "Internal Message File: $InternalMessageFile" -ForegroundColor Yellow
Write-Host "External Message File: $ExternalMessageFile" -ForegroundColor Yellow

# Validate user list
if (-not (Test-UserListFile)) {
    exit
}

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
}
catch {
    Write-Host "Error connecting to Exchange Online: $_" -ForegroundColor Red
    exit
}

try {
    # Load users
    $users = Import-Csv -Path $UserListPath
    $userColumn = if ($users[0].PSObject.Properties.Name -contains "UserPrincipalName") { "UserPrincipalName" } else { "Email" }
    $userCount = $users.Count
    
    if ($Action -eq "Backup") {
        # Perform backup of OOF settings
        Write-Host "Starting backup of OOF settings for $userCount users..." -ForegroundColor Yellow
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupFile = Join-Path -Path $OutputFolder -ChildPath "OOF_Backup_$timestamp.csv"
        $backupData = @()
        $success = 0
        $failed = 0
        
        foreach ($user in $users) {
            $userEmail = $user.$userColumn
            Write-Host "Backing up $userEmail..." -ForegroundColor Cyan -NoNewline
            
            try {
                $oofSettings = Get-MailboxAutoReplyConfiguration -Identity $userEmail
                
                $backupObject = [PSCustomObject]@{
                    UserPrincipalName = $userEmail
                    AutoReplyState = $oofSettings.AutoReplyState
                    StartTime = $oofSettings.StartTime
                    EndTime = $oofSettings.EndTime
                    InternalMessage = $oofSettings.InternalMessage
                    ExternalMessage = $oofSettings.ExternalMessage
                    ExternalAudience = $oofSettings.ExternalAudience
                    BackupTime = (Get-Date)
                }
                
                $backupData += $backupObject
                $success++
                Write-Host " Success" -ForegroundColor Green
            }
            catch {
                $failed++
                Write-Host " Failed: $_" -ForegroundColor Red
            }
        }
        
        # Export backup data to CSV
        $backupData | Export-Csv -Path $backupFile -NoTypeInformation
        Write-Host "Backup completed. Saved to: $backupFile" -ForegroundColor Green
        Write-Host "Summary: Success = $success, Failed = $failed" -ForegroundColor Yellow
    }
    else {
        # Set new OOF messages
        Write-Host "Configure new OOF settings for $userCount users..." -ForegroundColor Yellow
        
        # Get OOF configuration from user
        Write-Host "`nEnter OOF configuration:" -ForegroundColor Cyan
        $oofState = Read-Host "OOF State (Scheduled/Enabled/Disabled)"
        
        if ($oofState -ne "Scheduled" -and $oofState -ne "Enabled" -and $oofState -ne "Disabled") {
            Write-Host "Invalid OOF state. Please use 'Scheduled', 'Enabled', or 'Disabled'." -ForegroundColor Red
            exit
        }
        
        $startTime = $null
        $endTime = $null
        
        # Handle scheduling if selected
        if ($oofState -eq "Scheduled") {
            try {
                Write-Host "`nEnter start date and time in your local time zone:" -ForegroundColor Cyan
                $startDateStr = Read-Host "Start Date (MM/DD/YYYY)"
                $startTimeStr = Read-Host "Start Time (HH:MM AM/PM)"
                $startTime = [DateTime]::Parse("$startDateStr $startTimeStr")
                
                Write-Host "`nEnter end date and time in your local time zone:" -ForegroundColor Cyan
                $endDateStr = Read-Host "End Date (MM/DD/YYYY)"
                $endTimeStr = Read-Host "End Time (HH:MM AM/PM)"
                $endTime = [DateTime]::Parse("$endDateStr $endTimeStr")
                
                # Basic validation
                if ($endTime -le $startTime) {
                    Write-Host "Error: End time must be after start time." -ForegroundColor Red
                    exit
                }
                
                # Show time zone information
                $currentTZ = [System.TimeZoneInfo]::Local
                Write-Host "`nUsing time zone: $($currentTZ.DisplayName)" -ForegroundColor Yellow
                Write-Host "Start: $startTime" -ForegroundColor Yellow
                Write-Host "End: $endTime" -ForegroundColor Yellow
                
                $confirm = Read-Host "`nAre these times correct? (Y/N)"
                if ($confirm -ne "Y" -and $confirm -ne "y") {
                    Write-Host "Operation cancelled." -ForegroundColor Yellow
                    exit
                }
            }
            catch {
                Write-Host "Error parsing date/time: $_" -ForegroundColor Red
                exit
            }
        }
        
        # Get OOF messages from files if they exist, otherwise prompt user
        $internalMessage = ""
        $externalMessage = ""
        
        # Try to read internal message from file with absolute path resolution
        $internalFilePath = Resolve-Path -Path $InternalMessageFile -ErrorAction SilentlyContinue
        if (-not $internalFilePath) {
            $internalFilePath = Join-Path -Path $PSScriptRoot -ChildPath (Split-Path -Path $InternalMessageFile -Leaf)
        }
        Write-Host "`nLooking for internal message file at: $internalFilePath" -ForegroundColor Cyan
        
        if (Test-Path -Path $internalFilePath -PathType Leaf) {
            try {
                $internalMessage = Get-Content -Path $internalFilePath -Raw -ErrorAction Stop
                Write-Host "Successfully read internal message from file: $internalFilePath" -ForegroundColor Green
                Write-Host "Internal message preview (first 100 chars):" -ForegroundColor Cyan
                Write-Host ($internalMessage.Substring(0, [Math]::Min(100, $internalMessage.Length)) + "...") -ForegroundColor Gray
            }
            catch {
                Write-Host "Error reading internal message file: $_" -ForegroundColor Yellow
                Write-Host "`nEnter the internal OOF message (press Enter when done):" -ForegroundColor Cyan
                $internalMessage = Read-Host
            }
        }
        else {
            Write-Host "Internal message file not found at: $internalFilePath" -ForegroundColor Yellow
            
            # Try in current directory as fallback
            $altPath = ".\OOF_Internal.txt"
            Write-Host "Trying alternative path: $altPath" -ForegroundColor Yellow
            if (Test-Path -Path $altPath -PathType Leaf) {
                try {
                    $internalMessage = Get-Content -Path $altPath -Raw -ErrorAction Stop
                    Write-Host "Successfully read internal message from file: $altPath" -ForegroundColor Green
                    Write-Host "Internal message preview (first 100 chars):" -ForegroundColor Cyan
                    Write-Host ($internalMessage.Substring(0, [Math]::Min(100, $internalMessage.Length)) + "...") -ForegroundColor Gray
                }
                catch {
                    Write-Host "Error reading alternative internal message file: $_" -ForegroundColor Yellow
                    Write-Host "`nEnter the internal OOF message (press Enter when done):" -ForegroundColor Cyan
                    $internalMessage = Read-Host
                }
            }
            else {
                Write-Host "Alternative internal message file also not found" -ForegroundColor Yellow
                Write-Host "Enter the internal OOF message (press Enter when done):" -ForegroundColor Cyan
                $internalMessage = Read-Host
            }
        }
        
        # Try to read external message from file with absolute path resolution
        $externalFilePath = Resolve-Path -Path $ExternalMessageFile -ErrorAction SilentlyContinue
        if (-not $externalFilePath) {
            $externalFilePath = Join-Path -Path $PSScriptRoot -ChildPath (Split-Path -Path $ExternalMessageFile -Leaf)
        }
        Write-Host "`nLooking for external message file at: $externalFilePath" -ForegroundColor Cyan
        
        if (Test-Path -Path $externalFilePath -PathType Leaf) {
            try {
                $externalMessage = Get-Content -Path $externalFilePath -Raw -ErrorAction Stop
                Write-Host "Successfully read external message from file: $externalFilePath" -ForegroundColor Green
                Write-Host "External message preview (first 100 chars):" -ForegroundColor Cyan
                Write-Host ($externalMessage.Substring(0, [Math]::Min(100, $externalMessage.Length)) + "...") -ForegroundColor Gray
            }
            catch {
                Write-Host "Error reading external message file: $_" -ForegroundColor Yellow
                Write-Host "`nEnter the external OOF message (press Enter when done):" -ForegroundColor Cyan
                $externalMessage = Read-Host
            }
        }
        else {
            Write-Host "External message file not found at: $externalFilePath" -ForegroundColor Yellow
            
            # Try in current directory as fallback
            $altPath = ".\OOF_External.txt"
            Write-Host "Trying alternative path: $altPath" -ForegroundColor Yellow
            if (Test-Path -Path $altPath -PathType Leaf) {
                try {
                    $externalMessage = Get-Content -Path $altPath -Raw -ErrorAction Stop
                    Write-Host "Successfully read external message from file: $altPath" -ForegroundColor Green
                    Write-Host "External message preview (first 100 chars):" -ForegroundColor Cyan
                    Write-Host ($externalMessage.Substring(0, [Math]::Min(100, $externalMessage.Length)) + "...") -ForegroundColor Gray
                }
                catch {
                    Write-Host "Error reading alternative external message file: $_" -ForegroundColor Yellow
                    Write-Host "`nEnter the external OOF message (press Enter when done):" -ForegroundColor Cyan
                    $externalMessage = Read-Host
                }
            }
            else {
                Write-Host "Alternative external message file also not found" -ForegroundColor Yellow
                Write-Host "Enter the external OOF message (press Enter when done):" -ForegroundColor Cyan
                $externalMessage = Read-Host
            }
        }
        
        # Get external audience setting
        Write-Host "`nExternal audience options:" -ForegroundColor Cyan
        Write-Host "None - Don't send external replies" -ForegroundColor Gray
        Write-Host "Known - Only send replies to known senders (contacts)" -ForegroundColor Gray
        Write-Host "All - Send replies to all external senders" -ForegroundColor Gray
        $externalAudience = Read-Host "External Audience (None/Known/All)"
        
        if ($externalAudience -ne "None" -and $externalAudience -ne "Known" -and $externalAudience -ne "All") {
            Write-Host "Invalid external audience. Using 'All' as default." -ForegroundColor Yellow
            $externalAudience = "All"
        }
        
        # Final confirmation
        Write-Host "`nReady to apply these OOF settings to $userCount users." -ForegroundColor Yellow
        $finalConfirm = Read-Host "Proceed? (Y/N)"
        
        if ($finalConfirm -ne "Y" -and $finalConfirm -ne "y") {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            exit
        }
        
        # Apply OOF settings to each user
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $resultFile = Join-Path -Path $OutputFolder -ChildPath "OOF_Results_$timestamp.csv"
        $results = @()
        $success = 0
        $failed = 0
        
        foreach ($user in $users) {
            $userEmail = $user.$userColumn
            Write-Host "Setting OOF for $userEmail..." -ForegroundColor Cyan -NoNewline
            
            try {
                $params = @{
                    Identity = $userEmail
                    AutoReplyState = $oofState
                    ExternalAudience = $externalAudience
                    InternalMessage = $internalMessage
                    ExternalMessage = $externalMessage
                }
                
                if ($oofState -eq "Scheduled") {
                    $params.StartTime = $startTime
                    $params.EndTime = $endTime
                }
                
                Set-MailboxAutoReplyConfiguration @params
                
                $result = [PSCustomObject]@{
                    UserPrincipalName = $userEmail
                    AutoReplyState = $oofState
                    StartTime = if ($startTime) { $startTime.ToString() } else { "N/A" }
                    EndTime = if ($endTime) { $endTime.ToString() } else { "N/A" }
                    ExternalAudience = $externalAudience
                    Status = "Success"
                }
                
                $results += $result
                $success++
                Write-Host " Success" -ForegroundColor Green
            }
            catch {
                $result = [PSCustomObject]@{
                    UserPrincipalName = $userEmail
                    AutoReplyState = $oofState
                    StartTime = if ($startTime) { $startTime.ToString() } else { "N/A" }
                    EndTime = if ($endTime) { $endTime.ToString() } else { "N/A" }
                    ExternalAudience = $externalAudience
                    Status = "Failed: $_"
                }
                
                $results += $result
                $failed++
                Write-Host " Failed: $_" -ForegroundColor Red
            }
        }
        
        # Export results
        $results | Export-Csv -Path $resultFile -NoTypeInformation
        Write-Host "`nOOF deployment completed. Results saved to: $resultFile" -ForegroundColor Green
        Write-Host "Summary: Success = $success, Failed = $failed" -ForegroundColor Yellow
    }
}
finally {
    # Always disconnect from Exchange Online
    Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
}