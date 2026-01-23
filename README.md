# OutSystems ODC Email OAuth 2.0 Verification Script

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://docs.microsoft.com/en-us/powershell/)
[![Exchange Online](https://img.shields.io/badge/Exchange%20Online-Management-0078D4.svg)](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

🌐 **Language / Taal:** [Nederlands](#-doel) | [English](#-purpose)

---

Een interactief PowerShell script om alle vereiste Microsoft 365 / Exchange Online instellingen te verifiëren voor SMTP OAuth 2.0 authenticatie met OutSystems ODC (OutSystems Developer Cloud).

## 🎯 Doel

Dit script verifieert alle benodigde instellingen om e-mail te versturen vanuit OutSystems ODC via Microsoft Exchange Online met OAuth 2.0 authenticatie (Client Credentials Flow). Bij problemen biedt het script de mogelijkheid om deze interactief op te lossen (na gebruikersbevestiging). Het vervangt de verouderde Basic Authentication methode met een veilige en compliant setup.

## ✅ Wat wordt gecontroleerd?

| Stap | Controle | Kan oplossen (na bevestiging) |
|------|----------|-------------------------------|
| 0 | Verbinding met Exchange Online | - |
| 1 | Service Principal registratie in Exchange | ✅ |
| 2 | Mailbox informatie & aliassen | - |
| 3 | Mailbox permissions voor Service Principal | ✅ |
| 4 | SMTP AUTH instellingen (mailbox & tenant) | ✅ |
| 5 | SendFromAliasEnabled instelling | ✅ |

## 📋 Vereisten

### Software
- PowerShell 5.1 of hoger (Windows) of PowerShell 7+ (Cross-platform)
- [ExchangeOnlineManagement module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell) (wordt automatisch geïnstalleerd indien nodig)

### Azure / Microsoft 365
- Azure AD App Registration met:
  - `SMTP.SendAsApp` API permission (Office 365 Exchange Online)
  - Admin consent verleend
  - Client Secret aangemaakt
- Exchange Online mailbox voor verzenden (**Een normale (regular) mailbox. Geen shared mailbox!! Heeft ook een licentie nodig, anders wordt er geen mailbox aangemaakt in Exchange.**
- Administrator rechten in Exchange Online

### Benodigde gegevens
Zorg dat je de volgende gegevens bij de hand hebt voordat je het script uitvoert:

| Gegeven | Waar te vinden |
|---------|----------------|
| **Tenant ID** | Azure Portal → Microsoft Entra ID → Overview |
| **Application (Client) ID** | Azure Portal → App registrations → [Jouw App] → Overview |
| **Enterprise App Object ID** | Azure Portal → Enterprise applications → [Jouw App] → Overview |
| **Mailbox adres** | Het e-mailadres van de mailbox die gebruikt wordt voor verzenden |

> ⚠️ **Let op:** Het Object ID in Enterprise applications is **anders** dan het Object ID in App registrations!

## 🚀 Gebruik

### Stap 1: Script downloaden

```powershell
# Optie A: Clone de repository
git clone https://github.com/[jouw-username]/SMTP-OAuth20-MicrosoftAzure-Testscript.git

# Optie B: Download alleen het script
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/[jouw-username]/SMTP-OAuth20-MicrosoftAzure-Testscript/main/SMTP-OAuth20-MicrosoftAzure-Testscript.ps1" -OutFile "SMTP-OAuth20-MicrosoftAzure-Testscript.ps1"
```

### Stap 2: Execution Policy instellen (indien nodig)

```powershell
# Optie A: Voor huidige sessie (aanbevolen)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Optie B: Unblock het gedownloade bestand
Unblock-File -Path ".\SMTP-OAuth20-MicrosoftAzure-Testscript.ps1"
```

### Stap 3: Script uitvoeren

```powershell
.\SMTP-OAuth20-MicrosoftAzure-Testscript.ps1
```

Het script vraagt interactief om alle benodigde gegevens. Bij elke verificatiestap waar een probleem wordt gedetecteerd, krijg je de mogelijkheid om dit direct op te lossen (na bevestiging).

## 📸 Voorbeelduitvoer

```
=============================================
 OutSystems ODC Email OAuth 2.0 Verificatie
=============================================

[CONFIGURATIE] Voer de volgende gegevens in:

  Azure Tenant ID
  (Te vinden in Azure Portal > Microsoft Entra ID > Overview)
  Tenant ID: 12bb5b87-f62b-4d8a-9011-97b0a0b1bba6

  ...

[STAP 1] Service Principal in Exchange Online controleren...
  [OK] Service Principal gevonden:
       DisplayName: OutSystems ODC Mail Service
       ObjectId:    84aa014f-3fba-4cda-b80e-2390f6cb3e01
       AppId:       ee610a20-c543-4a05-970e-a5df69cf84f6

[STAP 2] Mailbox informatie ophalen...
  [OK] Mailbox gevonden:
       DisplayName:          OutSystems Mail Send account
       RecipientTypeDetails: UserMailbox
       PrimarySmtpAddress:   outsystems-noreply@contoso.com

  Email adressen (aliassen):
       [PRIMAIR] outsystems-noreply@contoso.com
       [ALIAS]   OutSystemsMailSendaccount@contoso.com

  ...
```

## ⚙️ OutSystems ODC Configuratie

Na succesvolle verificatie geeft het script een samenvatting met de exacte waarden voor OutSystems ODC Portal:

| Veld | Waarde |
|------|--------|
| **SMTP Server** | `smtp.office365.com` |
| **SMTP Port** | `587` |
| **Authentication** | `OAuth 2.0 - Client credentials` |
| **Server Token URL** | `https://login.microsoftonline.com/{TenantId}/oauth2/v2.0/token` |
| **Client ID** | `{Application ID uit Azure}` |
| **Client Secret** | `{Secret uit Azure App Registration}` |
| **Scope** | `https://outlook.office365.com/.default` |
| **Sender Email** | `{Primaire SMTP adres van de mailbox}` |

## 🔧 Handmatige configuratie

Als je de stappen handmatig wilt uitvoeren, zijn dit de PowerShell commando's:

### Verbinding maken
```powershell
Install-Module -Name ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -Organization <tenant-id>
```

### Service Principal aanmaken
```powershell
New-ServicePrincipal -AppId <APPLICATION_ID> -ObjectId <ENTERPRISE_APP_OBJECT_ID>
```

### Mailbox permissions toekennen
```powershell
Add-MailboxPermission -Identity "mailbox@contoso.com" -User <ENTERPRISE_APP_OBJECT_ID> -AccessRights FullAccess
```

### SMTP AUTH inschakelen
```powershell
# Per mailbox
Set-CASMailbox -Identity "mailbox@contoso.com" -SmtpClientAuthenticationDisabled $false

# Tenant-breed
Set-TransportConfig -SmtpClientAuthenticationDisabled $false
```

### Versturen vanuit aliassen toestaan (indien het adres afwijkt van het primaire SMTP adres van de mailbox!)
```powershell
Set-OrganizationConfig -SendFromAliasEnabled $True
```

## ❗ Veelvoorkomende problemen

### "Authentication unsuccessful" (535 5.7.3)
- Controleer of de scope correct is: `https://outlook.office365.com/.default`
- Controleer of SMTP AUTH is ingeschakeld voor de mailbox
- Controleer of de Service Principal is geregistreerd in Exchange Online

### "Cannot send from distribution group"
- Gebruik een mailbox adres, geen distributiegroep
- Controleer of je het primaire SMTP adres gebruikt (niet een alias)

### Script kan niet worden geladen (execution policy)
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

### Service Principal niet gevonden
- Gebruik het Object ID uit **Enterprise applications**, niet uit App registrations
- Wacht enkele minuten na het aanmaken van de app registration

## 📚 Referenties

- [OutSystems ODC - Configure SMTP Settings for Emails](https://success.outsystems.com/documentation/outsystems_developer_cloud/managing_outsystems_platform_and_apps/configure_smtp_settings_for_emails/)
- [Microsoft - Authenticate SMTP connection using OAuth](https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth)
- [Microsoft - Enable or disable SMTP AUTH](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/authenticated-client-smtp-submission)

## 📄 Licentie

Dit project is gelicentieerd onder de MIT License - zie het [LICENSE](LICENSE) bestand voor details.

## 🤝 Bijdragen

Bijdragen zijn welkom! Voel je vrij om een issue aan te maken of een pull request in te dienen.

---

# 🇬🇧 English Version

An interactive PowerShell script to verify all required Microsoft 365 / Exchange Online settings for SMTP OAuth 2.0 authentication with OutSystems ODC (OutSystems Developer Cloud).

## 🎯 Purpose

This script verifies all required settings to send email from OutSystems ODC via Microsoft Exchange Online with OAuth 2.0 authentication (Client Credentials Flow). When issues are detected, the script offers the option to resolve them interactively (after user confirmation). It replaces the deprecated Basic Authentication method with a secure and compliant setup.

## ✅ What is verified?

| Step | Check | Can fix (after confirmation) |
|------|-------|------------------------------|
| 0 | Connection to Exchange Online | - |
| 1 | Service Principal registration in Exchange | ✅ |
| 2 | Mailbox information & aliases | - |
| 3 | Mailbox permissions for Service Principal | ✅ |
| 4 | SMTP AUTH settings (mailbox & tenant) | ✅ |
| 5 | SendFromAliasEnabled setting | ✅ |

## 📋 Requirements

### Software
- PowerShell 5.1 or higher (Windows) or PowerShell 7+ (Cross-platform)
- [ExchangeOnlineManagement module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell) (automatically installed if needed)

### Azure / Microsoft 365
- Azure AD App Registration with:
  - `SMTP.SendAsApp` API permission (Office 365 Exchange Online)
  - Admin consent granted
  - Client Secret created
- Exchange Online mailbox for sending (**A normal (regular) mailbox. Not a shared mailbox!! Also requires a license, otherwise no mailbox is created in Exchange.**)
- Administrator rights in Exchange Online

### Required information
Make sure you have the following information ready before running the script:

| Information | Where to find |
|-------------|---------------|
| **Tenant ID** | Azure Portal → Microsoft Entra ID → Overview |
| **Application (Client) ID** | Azure Portal → App registrations → [Your App] → Overview |
| **Enterprise App Object ID** | Azure Portal → Enterprise applications → [Your App] → Overview |
| **Mailbox address** | The email address of the mailbox used for sending |

> ⚠️ **Note:** The Object ID in Enterprise applications is **different** from the Object ID in App registrations!

## 🚀 Usage

### Step 1: Download the script

```powershell
# Option A: Clone the repository
git clone https://github.com/[your-username]/SMTP-OAuth20-MicrosoftAzure-Testscript.git

# Option B: Download only the script
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/[your-username]/SMTP-OAuth20-MicrosoftAzure-Testscript/main/SMTP-OAuth20-MicrosoftAzure-Testscript.ps1" -OutFile "SMTP-OAuth20-MicrosoftAzure-Testscript.ps1"
```

### Step 2: Set Execution Policy (if needed)

```powershell
# Option A: For current session (recommended)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Option B: Unblock the downloaded file
Unblock-File -Path ".\SMTP-OAuth20-MicrosoftAzure-Testscript.ps1"
```

### Step 3: Run the script

```powershell
.\SMTP-OAuth20-MicrosoftAzure-Testscript.ps1
```

The script interactively prompts for all required information. At each verification step where an issue is detected, you get the option to resolve it directly (after confirmation).

## 📸 Example output

```
=============================================
 OutSystems ODC Email OAuth 2.0 Verificatie
=============================================

[CONFIGURATIE] Voer de volgende gegevens in:

  Azure Tenant ID
  (Te vinden in Azure Portal > Microsoft Entra ID > Overview)
  Tenant ID: 12bb5b87-f62b-4d8a-9011-97b0a0b1bba6

  ...

[STAP 1] Service Principal in Exchange Online controleren...
  [OK] Service Principal gevonden:
       DisplayName: OutSystems ODC Mail Service
       ObjectId:    84aa014f-3fba-4cda-b80e-2390f6cb3e01
       AppId:       ee610a20-c543-4a05-970e-a5df69cf84f6

[STAP 2] Mailbox informatie ophalen...
  [OK] Mailbox gevonden:
       DisplayName:          OutSystems Mail Send account
       RecipientTypeDetails: UserMailbox
       PrimarySmtpAddress:   outsystems-noreply@contoso.com

  Email adressen (aliassen):
       [PRIMAIR] outsystems-noreply@contoso.com
       [ALIAS]   OutSystemsMailSendaccount@contoso.com

  ...
```

## ⚙️ OutSystems ODC Configuration

After successful verification, the script provides a summary with the exact values for OutSystems ODC Portal:

| Field | Value |
|-------|-------|
| **SMTP Server** | `smtp.office365.com` |
| **SMTP Port** | `587` |
| **Authentication** | `OAuth 2.0 - Client credentials` |
| **Server Token URL** | `https://login.microsoftonline.com/{TenantId}/oauth2/v2.0/token` |
| **Client ID** | `{Application ID from Azure}` |
| **Client Secret** | `{Secret from Azure App Registration}` |
| **Scope** | `https://outlook.office365.com/.default` |
| **Sender Email** | `{Primary SMTP address of the mailbox}` |

## 🔧 Manual configuration

If you want to perform the steps manually, these are the PowerShell commands:

### Connect
```powershell
Install-Module -Name ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -Organization <tenant-id>
```

### Create Service Principal
```powershell
New-ServicePrincipal -AppId <APPLICATION_ID> -ObjectId <ENTERPRISE_APP_OBJECT_ID>
```

### Grant mailbox permissions
```powershell
Add-MailboxPermission -Identity "mailbox@contoso.com" -User <ENTERPRISE_APP_OBJECT_ID> -AccessRights FullAccess
```

### Enable SMTP AUTH
```powershell
# Per mailbox
Set-CASMailbox -Identity "mailbox@contoso.com" -SmtpClientAuthenticationDisabled $false

# Tenant-wide
Set-TransportConfig -SmtpClientAuthenticationDisabled $false
```

### Allow sending from aliases (if the address differs from the primary SMTP address of the mailbox!)
```powershell
Set-OrganizationConfig -SendFromAliasEnabled $True
```

## ❗ Common issues

### "Authentication unsuccessful" (535 5.7.3)
- Check if the scope is correct: `https://outlook.office365.com/.default`
- Check if SMTP AUTH is enabled for the mailbox
- Check if the Service Principal is registered in Exchange Online

### "Cannot send from distribution group"
- Use a mailbox address, not a distribution group
- Check if you are using the primary SMTP address (not an alias)

### Script cannot be loaded (execution policy)
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

### Service Principal not found
- Use the Object ID from **Enterprise applications**, not from App registrations
- Wait a few minutes after creating the app registration

## 📚 References

- [OutSystems ODC - Configure SMTP Settings for Emails](https://success.outsystems.com/documentation/outsystems_developer_cloud/managing_outsystems_platform_and_apps/configure_smtp_settings_for_emails/)
- [Microsoft - Authenticate SMTP connection using OAuth](https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth)
- [Microsoft - Enable or disable SMTP AUTH](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/authenticated-client-smtp-submission)

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

Contributions are welcome! Feel free to create an issue or submit a pull request.

---

**Made with ❤️ for the OutSystems community**
