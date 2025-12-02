#############################################################################
# OutSystems ODC - Microsoft 365 OAuth 2.0 Email Configuratie Verificatie
# 
# Dit script controleert alle vereiste instellingen voor SMTP OAuth 2.0
# authenticatie tussen OutSystems ODC en Microsoft Exchange Online.
#
# Gebruik: Voer het script uit en volg de prompts
#############################################################################

Clear-Host
Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " OutSystems ODC Email OAuth 2.0 Verificatie" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host " Dit script controleert alle vereiste instellingen voor" -ForegroundColor Gray
Write-Host " SMTP OAuth 2.0 authenticatie met Microsoft Exchange Online." -ForegroundColor Gray
Write-Host ""
Write-Host "---------------------------------------------" -ForegroundColor DarkGray
Write-Host ""

# =====================================================
# INVOER VERZAMELEN
# =====================================================

Write-Host "[CONFIGURATIE] Voer de volgende gegevens in:" -ForegroundColor Yellow
Write-Host ""

# Tenant ID
Write-Host "  Azure Tenant ID" -ForegroundColor White
Write-Host "  (Te vinden in Azure Portal > Microsoft Entra ID > Overview)" -ForegroundColor DarkGray
$TenantId = Read-Host "  Tenant ID"
if ([string]::IsNullOrWhiteSpace($TenantId)) {
    Write-Host "  [FOUT] Tenant ID is verplicht. Script wordt afgebroken." -ForegroundColor Red
    exit 1
}
Write-Host ""

# App ID (Client ID)
Write-Host "  Application (Client) ID" -ForegroundColor White
Write-Host "  (Te vinden in Azure Portal > App registrations > [Jouw App] > Overview)" -ForegroundColor DarkGray
$AppId = Read-Host "  App ID"
if ([string]::IsNullOrWhiteSpace($AppId)) {
    Write-Host "  [FOUT] App ID is verplicht. Script wordt afgebroken." -ForegroundColor Red
    exit 1
}
Write-Host ""

# Enterprise Application Object ID
Write-Host "  Enterprise Application Object ID" -ForegroundColor White
Write-Host "  (Te vinden in Azure Portal > Enterprise applications > [Jouw App] > Overview)" -ForegroundColor DarkGray
Write-Host "  LET OP: Dit is een ANDER Object ID dan in App registrations!" -ForegroundColor Yellow
$EnterpriseAppObjectId = Read-Host "  Object ID"
if ([string]::IsNullOrWhiteSpace($EnterpriseAppObjectId)) {
    Write-Host "  [FOUT] Enterprise App Object ID is verplicht. Script wordt afgebroken." -ForegroundColor Red
    exit 1
}
Write-Host ""

# Mailbox Address
Write-Host "  Mailbox adres voor verzenden" -ForegroundColor White
Write-Host "  (Het email adres van de mailbox die gebruikt wordt voor versturen)" -ForegroundColor DarkGray
$MailboxAddress = Read-Host "  Mailbox adres"
if ([string]::IsNullOrWhiteSpace($MailboxAddress)) {
    Write-Host "  [FOUT] Mailbox adres is verplicht. Script wordt afgebroken." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "---------------------------------------------" -ForegroundColor DarkGray
Write-Host ""
Write-Host "[OVERZICHT] Ingevoerde gegevens:" -ForegroundColor Yellow
Write-Host "  Tenant ID:      $TenantId" -ForegroundColor White
Write-Host "  App ID:         $AppId" -ForegroundColor White
Write-Host "  Object ID:      $EnterpriseAppObjectId" -ForegroundColor White
Write-Host "  Mailbox:        $MailboxAddress" -ForegroundColor White
Write-Host ""

$confirm = Read-Host "  Zijn deze gegevens correct? (J/N)"
if ($confirm -notmatch "^[JjYy]") {
    Write-Host "  Script afgebroken. Voer het script opnieuw uit." -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " VERIFICATIE STARTEN" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan

# =====================================================
# STAP 0: VERBINDING MAKEN
# =====================================================

Write-Host ""
Write-Host "[STAP 0] Verbinding maken met Exchange Online..." -ForegroundColor Yellow
Write-Host ""

try {
    # Controleer of ExchangeOnlineManagement module is geinstalleerd
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "  [WAARSCHUWING] ExchangeOnlineManagement module niet gevonden." -ForegroundColor Yellow
        Write-Host "  Installeren met: Install-Module -Name ExchangeOnlineManagement" -ForegroundColor Cyan
        $installModule = Read-Host "  Wil je de module nu installeren? (J/N)"
        if ($installModule -match "^[JjYy]") {
            Write-Host "  Module installeren..." -ForegroundColor Gray
            Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
            Import-Module ExchangeOnlineManagement
        } else {
            Write-Host "  [FOUT] Module is vereist. Script wordt afgebroken." -ForegroundColor Red
            exit 1
        }
    }

    # Controleer of er al een sessie is
    $existingSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
    
    if (-not $existingSession) {
        Write-Host "  Verbinden met Exchange Online (er opent een login venster)..." -ForegroundColor Gray
        Connect-ExchangeOnline -Organization $TenantId -ShowBanner:$false
    } else {
        Write-Host "  Bestaande Exchange Online sessie gevonden." -ForegroundColor Gray
    }
    Write-Host "  [OK] Verbonden met Exchange Online" -ForegroundColor Green
} catch {
    Write-Host "  [FOUT] Kan niet verbinden met Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "---------------------------------------------" -ForegroundColor DarkGray

# =====================================================
# STAP 1: SERVICE PRINCIPAL CONTROLE
# =====================================================

Write-Host ""
Write-Host "[STAP 1] Service Principal in Exchange Online controleren..." -ForegroundColor Yellow
Write-Host "  App ID: $AppId" -ForegroundColor Gray
Write-Host ""

$servicePrincipalFound = $false
try {
    $servicePrincipal = Get-ServicePrincipal | Where-Object { $_.AppId -eq $AppId }
    
    if ($servicePrincipal) {
        $servicePrincipalFound = $true
        Write-Host "  [OK] Service Principal gevonden:" -ForegroundColor Green
        Write-Host "       DisplayName: $($servicePrincipal.DisplayName)" -ForegroundColor White
        Write-Host "       ObjectId:    $($servicePrincipal.ObjectId)" -ForegroundColor White
        Write-Host "       AppId:       $($servicePrincipal.AppId)" -ForegroundColor White
    } else {
        Write-Host "  [FOUT] Service Principal NIET gevonden!" -ForegroundColor Red
        Write-Host ""
        Write-Host "  Oplossing - voer dit commando uit:" -ForegroundColor Yellow
        Write-Host "  New-ServicePrincipal -AppId $AppId -ObjectId $EnterpriseAppObjectId" -ForegroundColor Cyan
        Write-Host ""
        $createSP = Read-Host "  Wil je de Service Principal nu aanmaken? (J/N)"
        if ($createSP -match "^[JjYy]") {
            try {
                New-ServicePrincipal -AppId $AppId -ObjectId $EnterpriseAppObjectId
                Write-Host "  [OK] Service Principal aangemaakt!" -ForegroundColor Green
                $servicePrincipal = Get-ServicePrincipal | Where-Object { $_.AppId -eq $AppId }
                $servicePrincipalFound = $true
            } catch {
                Write-Host "  [FOUT] Kan Service Principal niet aanmaken: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
} catch {
    Write-Host "  [FOUT] Kan Service Principal niet controleren: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "---------------------------------------------" -ForegroundColor DarkGray

# =====================================================
# STAP 2: MAILBOX INFORMATIE
# =====================================================

Write-Host ""
Write-Host "[STAP 2] Mailbox informatie ophalen..." -ForegroundColor Yellow
Write-Host "  Mailbox: $MailboxAddress" -ForegroundColor Gray
Write-Host ""

$mailboxFound = $false
$primaryAddress = $null
try {
    $mailbox = Get-Mailbox -Identity $MailboxAddress -ErrorAction Stop
    $mailboxFound = $true
    
    Write-Host "  [OK] Mailbox gevonden:" -ForegroundColor Green
    Write-Host "       DisplayName:          $($mailbox.DisplayName)" -ForegroundColor White
    Write-Host "       RecipientTypeDetails: $($mailbox.RecipientTypeDetails)" -ForegroundColor White
    Write-Host "       PrimarySmtpAddress:   $($mailbox.PrimarySmtpAddress)" -ForegroundColor White
    
    $primaryAddress = $mailbox.PrimarySmtpAddress.ToString()
    
    # Waarschuwing als het een SharedMailbox is
    if ($mailbox.RecipientTypeDetails -eq "SharedMailbox") {
        Write-Host ""
        Write-Host "  [WAARSCHUWING] Dit is een Shared Mailbox!" -ForegroundColor Yellow
        Write-Host "  Shared mailboxes kunnen problemen geven met SMTP OAuth." -ForegroundColor Yellow
        Write-Host "  Overweeg een UserMailbox te gebruiken." -ForegroundColor Yellow
    }
    
    # Email aliassen tonen
    Write-Host ""
    Write-Host "  Email adressen (aliassen):" -ForegroundColor Cyan
    $mailbox.EmailAddresses | ForEach-Object {
        $addrString = $_.ToString()
        if ($addrString -cmatch "^SMTP:") {
            $address = $addrString.Replace("SMTP:", "")
            Write-Host "       [PRIMAIR] $address" -ForegroundColor Green
        } elseif ($addrString -cmatch "^smtp:") {
            $address = $addrString.Replace("smtp:", "")
            Write-Host "       [ALIAS]   $address" -ForegroundColor White
        } elseif ($addrString -cmatch "^SIP:") {
            $address = $addrString.Replace("SIP:", "")
            Write-Host "       [SIP]     $address" -ForegroundColor DarkGray
        }
    }
    
    # Controleer of het geconfigureerde adres het primaire adres is
    if ($MailboxAddress.ToLower() -ne $primaryAddress.ToLower()) {
        Write-Host ""
        Write-Host "  [LET OP] Het ingevoerde adres ($MailboxAddress)" -ForegroundColor Yellow
        Write-Host "           is een ALIAS, niet het primaire adres." -ForegroundColor Yellow
        Write-Host "           Primair adres is: $primaryAddress" -ForegroundColor Yellow
        Write-Host "           Zorg dat 'SendFromAliasEnabled' aan staat (zie stap 5)" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "  [FOUT] Mailbox niet gevonden: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "---------------------------------------------" -ForegroundColor DarkGray

# =====================================================
# STAP 3: MAILBOX PERMISSIONS CONTROLE
# =====================================================

Write-Host ""
Write-Host "[STAP 3] Mailbox permissions controleren..." -ForegroundColor Yellow
Write-Host ""

if ($mailboxFound) {
    try {
        $allPermissions = Get-MailboxPermission -Identity $MailboxAddress | Where-Object { $_.User -notlike "NT AUTHORITY\*" }
        
        $spPermission = $allPermissions | Where-Object { 
            $_.User -like "*$EnterpriseAppObjectId*" -or
            ($servicePrincipal -and $_.User -like "*$($servicePrincipal.DisplayName)*")
        }
        
        if ($spPermission) {
            Write-Host "  [OK] Service Principal heeft permissions op mailbox:" -ForegroundColor Green
            $spPermission | ForEach-Object {
                Write-Host "       User: $($_.User)" -ForegroundColor White
                Write-Host "       AccessRights: $($_.AccessRights -join ', ')" -ForegroundColor White
            }
        } else {
            Write-Host "  [WAARSCHUWING] Service Principal permission niet gevonden." -ForegroundColor Yellow
            Write-Host ""
            if ($allPermissions) {
                Write-Host "  Huidige permissions op deze mailbox:" -ForegroundColor Gray
                $allPermissions | ForEach-Object {
                    Write-Host "       $($_.User) - $($_.AccessRights -join ', ')" -ForegroundColor White
                }
            } else {
                Write-Host "  Geen speciale permissions geconfigureerd." -ForegroundColor Gray
            }
            Write-Host ""
            Write-Host "  Oplossing - voer dit commando uit:" -ForegroundColor Yellow
            Write-Host "  Add-MailboxPermission -Identity `"$MailboxAddress`" -User $EnterpriseAppObjectId -AccessRights FullAccess" -ForegroundColor Cyan
            Write-Host ""
            $addPerm = Read-Host "  Wil je de permission nu toevoegen? (J/N)"
            if ($addPerm -match "^[JjYy]") {
                try {
                    Add-MailboxPermission -Identity $MailboxAddress -User $EnterpriseAppObjectId -AccessRights FullAccess -ErrorAction Stop
                    Write-Host "  [OK] Permission toegevoegd!" -ForegroundColor Green
                } catch {
                    Write-Host "  [FOUT] Kan permission niet toevoegen: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
    } catch {
        Write-Host "  [FOUT] Kan permissions niet controleren: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "  [OVERGESLAGEN] Mailbox niet gevonden, kan permissions niet controleren." -ForegroundColor Gray
}

Write-Host ""
Write-Host "---------------------------------------------" -ForegroundColor DarkGray

# =====================================================
# STAP 4: SMTP AUTH INSTELLINGEN
# =====================================================

Write-Host ""
Write-Host "[STAP 4] SMTP AUTH instellingen controleren..." -ForegroundColor Yellow
Write-Host ""

if ($mailboxFound) {
    # Mailbox-niveau
    try {
        $casMailbox = Get-CASMailbox -Identity $MailboxAddress -ErrorAction Stop
        $smtpDisabled = $casMailbox.SmtpClientAuthenticationDisabled
        
        if ($smtpDisabled -eq $false) {
            Write-Host "  [OK] SMTP AUTH is INGESCHAKELD voor mailbox" -ForegroundColor Green
        } elseif ($smtpDisabled -eq $true) {
            Write-Host "  [FOUT] SMTP AUTH is UITGESCHAKELD voor mailbox" -ForegroundColor Red
            Write-Host ""
            Write-Host "  Oplossing:" -ForegroundColor Yellow
            Write-Host "  Set-CASMailbox -Identity `"$MailboxAddress`" -SmtpClientAuthenticationDisabled `$false" -ForegroundColor Cyan
            Write-Host ""
            $enableSmtp = Read-Host "  Wil je SMTP AUTH nu inschakelen voor deze mailbox? (J/N)"
            if ($enableSmtp -match "^[JjYy]") {
                try {
                    Set-CASMailbox -Identity $MailboxAddress -SmtpClientAuthenticationDisabled $false
                    Write-Host "  [OK] SMTP AUTH ingeschakeld voor mailbox!" -ForegroundColor Green
                } catch {
                    Write-Host "  [FOUT] Kan SMTP AUTH niet inschakelen: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "  [INFO] SMTP AUTH volgt tenant-instelling (niet expliciet ingesteld voor mailbox)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  [FOUT] Kan CAS mailbox niet controleren: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "  [OVERGESLAGEN] Mailbox niet gevonden." -ForegroundColor Gray
}

# Tenant-niveau
Write-Host ""
try {
    $transportConfig = Get-TransportConfig
    $tenantSmtpDisabled = $transportConfig.SmtpClientAuthenticationDisabled
    
    if ($tenantSmtpDisabled -eq $false) {
        Write-Host "  [OK] SMTP AUTH is INGESCHAKELD op tenant-niveau" -ForegroundColor Green
    } else {
        Write-Host "  [FOUT] SMTP AUTH is UITGESCHAKELD op tenant-niveau" -ForegroundColor Red
        Write-Host ""
        Write-Host "  Oplossing:" -ForegroundColor Yellow
        Write-Host "  Set-TransportConfig -SmtpClientAuthenticationDisabled `$false" -ForegroundColor Cyan
        Write-Host ""
        $enableTenantSmtp = Read-Host "  Wil je SMTP AUTH nu inschakelen op tenant-niveau? (J/N)"
        if ($enableTenantSmtp -match "^[JjYy]") {
            try {
                Set-TransportConfig -SmtpClientAuthenticationDisabled $false
                Write-Host "  [OK] SMTP AUTH ingeschakeld op tenant-niveau!" -ForegroundColor Green
            } catch {
                Write-Host "  [FOUT] Kan SMTP AUTH niet inschakelen: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
} catch {
    Write-Host "  [FOUT] Kan transport config niet controleren: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "---------------------------------------------" -ForegroundColor DarkGray

# =====================================================
# STAP 5: SEND FROM ALIAS INSTELLING
# =====================================================

Write-Host ""
Write-Host "[STAP 5] Send From Alias instelling controleren..." -ForegroundColor Yellow
Write-Host ""

try {
    $orgConfig = Get-OrganizationConfig
    $sendFromAlias = $orgConfig.SendFromAliasEnabled
    
    if ($sendFromAlias -eq $true) {
        Write-Host "  [OK] SendFromAliasEnabled is INGESCHAKELD" -ForegroundColor Green
        Write-Host "       (Versturen vanuit aliassen is toegestaan)" -ForegroundColor Gray
    } else {
        Write-Host "  [INFO] SendFromAliasEnabled is UITGESCHAKELD" -ForegroundColor Yellow
        Write-Host "       (Alleen versturen vanuit primair adres toegestaan)" -ForegroundColor Gray
        
        # Check of gebruiker een alias gebruikt
        if ($mailboxFound -and $primaryAddress -and ($MailboxAddress.ToLower() -ne $primaryAddress.ToLower())) {
            Write-Host ""
            Write-Host "  [WAARSCHUWING] Je gebruikt een alias als sender adres!" -ForegroundColor Red
            Write-Host "  Dit zal NIET werken zonder SendFromAliasEnabled." -ForegroundColor Red
            Write-Host ""
            Write-Host "  Opties:" -ForegroundColor Yellow
            Write-Host "  1. Gebruik het primaire adres: $primaryAddress" -ForegroundColor White
            Write-Host "  2. Schakel SendFromAliasEnabled in" -ForegroundColor White
            Write-Host ""
            $enableAlias = Read-Host "  Wil je SendFromAliasEnabled nu inschakelen? (J/N)"
            if ($enableAlias -match "^[JjYy]") {
                try {
                    Set-OrganizationConfig -SendFromAliasEnabled $true
                    Write-Host "  [OK] SendFromAliasEnabled ingeschakeld!" -ForegroundColor Green
                } catch {
                    Write-Host "  [FOUT] Kan SendFromAliasEnabled niet inschakelen: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        } else {
            Write-Host ""
            Write-Host "  Als je vanuit aliassen wilt versturen:" -ForegroundColor Yellow
            Write-Host "  Set-OrganizationConfig -SendFromAliasEnabled `$True" -ForegroundColor Cyan
        }
    }
} catch {
    Write-Host "  [FOUT] Kan organization config niet controleren: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "---------------------------------------------" -ForegroundColor DarkGray

# =====================================================
# SAMENVATTING
# =====================================================

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " SAMENVATTING - OutSystems ODC Configuratie" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Gebruik deze waarden in OutSystems ODC Portal > Configure > Emails:" -ForegroundColor White
Write-Host ""
Write-Host "  ----------------------------------------" -ForegroundColor DarkGray
Write-Host "  SERVER" -ForegroundColor Yellow
Write-Host "  ----------------------------------------" -ForegroundColor DarkGray
Write-Host "  SMTP Server:        smtp.office365.com" -ForegroundColor Green
Write-Host "  SMTP Port:          587" -ForegroundColor Green
Write-Host ""
Write-Host "  ----------------------------------------" -ForegroundColor DarkGray
Write-Host "  AUTHENTICATION" -ForegroundColor Yellow
Write-Host "  ----------------------------------------" -ForegroundColor DarkGray
Write-Host "  Type:               OAuth 2.0 - Client credentials" -ForegroundColor Green
Write-Host "  Server Token URL:   https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -ForegroundColor Green
Write-Host "  Client ID:          $AppId" -ForegroundColor Green
Write-Host "  Client Secret:      [Uit Azure App Registration > Certificates & secrets]" -ForegroundColor Green
Write-Host "  Scope:              https://outlook.office365.com/.default" -ForegroundColor Green
Write-Host ""
Write-Host "  ----------------------------------------" -ForegroundColor DarkGray
Write-Host "  SENDER" -ForegroundColor Yellow
Write-Host "  ----------------------------------------" -ForegroundColor DarkGray

# Bepaal welk adres te gebruiken
if ($mailboxFound -and $primaryAddress) {
    try {
        $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
        if ($orgConfig.SendFromAliasEnabled -eq $true) {
            Write-Host "  Email:              $MailboxAddress" -ForegroundColor Green
            Write-Host "                      (of gebruik primair: $primaryAddress)" -ForegroundColor Gray
        } else {
            if ($MailboxAddress.ToLower() -ne $primaryAddress.ToLower()) {
                Write-Host "  Email:              $primaryAddress" -ForegroundColor Green
                Write-Host "                      (Let op: $MailboxAddress is een alias en werkt NIET)" -ForegroundColor Red
            } else {
                Write-Host "  Email:              $primaryAddress" -ForegroundColor Green
            }
        }
    } catch {
        Write-Host "  Email:              $primaryAddress (primair adres aanbevolen)" -ForegroundColor Green
    }
} else {
    Write-Host "  Email:              [Gebruik het PRIMAIRE email adres van de mailbox]" -ForegroundColor Green
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " VERIFICATIE VOLTOOID" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""
