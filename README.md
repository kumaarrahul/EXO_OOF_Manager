
<#
# ===================================================================================================
# Script Name: Exchange Online OOF Manager
# Created: Apr 2025
# Author: Rahul Kumaar
# Version: 1.0
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
