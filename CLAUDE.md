# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This repository contains an interactive PowerShell verification script for configuring SMTP OAuth 2.0 authentication between OutSystems ODC (OutSystems Developer Cloud) and Microsoft Exchange Online. The script validates and configures all required settings for secure email sending via Microsoft 365.

## Script Execution

```powershell
# Run the main verification script
.\SMTP-OAuth20-MicrosoftAzure-Testscript.ps1
```

The script runs interactively and will prompt for:
- Azure Tenant ID
- Application (Client) ID
- Enterprise Application Object ID (different from App Registration Object ID)
- Mailbox address for sending

## Key Architecture

### Verification Flow

The script performs sequential verification steps:

1. **Step 0**: Connect to Exchange Online (requires ExchangeOnlineManagement module)
2. **Step 1**: Verify Service Principal exists in Exchange Online
3. **Step 2**: Retrieve mailbox information and aliases
4. **Step 3**: Check mailbox permissions for Service Principal
5. **Step 4**: Verify SMTP AUTH settings (both mailbox-level and tenant-level)
6. **Step 5**: Check SendFromAliasEnabled configuration

### Interactive Remediation

At each step, if a problem is detected, the script:
- Displays the current state in color-coded output (Red=Error, Yellow=Warning, Green=OK)
- Shows the PowerShell command needed to fix the issue
- Offers to automatically fix the problem if the user confirms

### Critical Distinctions

- **Enterprise Application Object ID** vs **App Registration Object ID**: These are different values from different Azure Portal locations. The script requires the Enterprise Application Object ID.
- **Primary SMTP Address** vs **Alias**: The script detects if an alias is being used and warns that SendFromAliasEnabled must be true for aliases to work.
- **Mailbox-level** vs **Tenant-level** SMTP AUTH: Both need to be enabled for OAuth to work.

## PowerShell Commands Used

The script uses Exchange Online Management cmdlets:
- `Connect-ExchangeOnline`: Authenticate to Exchange Online
- `Get-ServicePrincipal` / `New-ServicePrincipal`: Manage service principals
- `Get-Mailbox`: Retrieve mailbox information
- `Get-MailboxPermission` / `Add-MailboxPermission`: Manage mailbox access
- `Get-CASMailbox` / `Set-CASMailbox`: Configure client access settings
- `Get-TransportConfig` / `Set-TransportConfig`: Configure tenant-wide transport settings
- `Get-OrganizationConfig` / `Set-OrganizationConfig`: Configure organization settings

## OAuth 2.0 Configuration Details

The script validates configuration for OAuth 2.0 Client Credentials flow with:
- **Token URL**: `https://login.microsoftonline.com/{TenantId}/oauth2/v2.0/token`
- **Scope**: `https://outlook.office365.com/.default`
- **Required API Permission**: `SMTP.SendAsApp` (application permission in Azure AD)
- **SMTP Endpoint**: smtp.office365.com:587

## Output Summary

At completion, the script provides a formatted summary with all values needed for OutSystems ODC Portal configuration, including server settings, authentication parameters, and the correct sender email address.
