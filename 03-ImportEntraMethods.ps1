#requires -Version 5.1

<#
.SYNOPSIS
Adds, updates, or removes users' mobile-phone and authentication-email methods
in Entra ID.

.DESCRIPTION
Reads a CSV with UserPrincipalName, MobilePhone, and AuthenticationEmail columns.
By default, performs a preview and makes no changes. Supply -Apply to write the
authentication methods to Microsoft Entra ID.

A MobilePhone value adds or updates the mobile phone authentication method. A
blank MobilePhone cell removes the existing mobile phone authentication method
(this is how Script 02 propagates a cleared extensionAttribute2). Blank
AuthenticationEmail cells are skipped; this script never deletes an
authentication-email method.

.EXAMPLE
.\Set-EntraPhoneEmailAuthenticationMethods.ps1 -Apply

    Uses the default CSV (.\EntraMethods.csv) produced by 02-SyncPhoneAttribute.ps1.

.EXAMPLE
.\Set-EntraPhoneEmailAuthenticationMethods.ps1 -CsvPath .\EntraMethods-OU.csv -Apply -LogPath .\EntraMethods-results.csv

.NOTES
Permissions:
  - Delegated mode (default, device-code sign-in): requires the Graph delegated
    permission UserAuthenticationMethod.ReadWrite.All and an Authentication
    Administrator or Privileged Authentication Administrator role. Managing
    authentication methods for administrators generally requires Privileged
    Authentication Administrator.
  - App-only mode (supply -ClientId and -CertificateThumbprint together): the app
    registration must be granted the application permission
    UserAuthenticationMethod.ReadWrite.All with admin consent, and the certificate
    must be installed in Cert:\CurrentUser\My or Cert:\LocalMachine\My.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$CsvPath = ".\EntraMethods.csv",

    [switch]$Apply,

    [string]$LogPath,

    [string]$TenantId = '3ec35b81-463a-4801-864c-fedf6e44de2e',

    [string]$ClientId = 'e5ec9711-41fb-494a-86d8-2bdf70859e8c',

    [string]$CertificateThumbprint = 'a668ab98819e1be444d7d486029e4d4ad6cdb20c',

    [string]$RequiredVersion = '2.33.0'
)

$ErrorActionPreference = 'Stop'

$requiredModuleVersion = $RequiredVersion
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Where-Object { $_.Version -eq $requiredModuleVersion })) {
    throw @"
Microsoft.Graph.Authentication $requiredModuleVersion is not installed. Install it with:
  Install-Module Microsoft.Graph.Authentication -RequiredVersion $requiredModuleVersion -Scope AllUsers -Force
"@
}

Import-Module Microsoft.Graph.Authentication -RequiredVersion $requiredModuleVersion -Force

$requiredScope = 'UserAuthenticationMethod.ReadWrite.All'
$useAppOnly = $ClientId -and $CertificateThumbprint

if ($useAppOnly) {
    if (-not $TenantId) {
        throw 'TenantId is required when using certificate-based app-only authentication.'
    }

    $certFound = Get-ChildItem -Path Cert:\CurrentUser\My -ErrorAction SilentlyContinue |
        Where-Object { $_.Thumbprint -eq $CertificateThumbprint }
    if (-not $certFound) {
        $certFound = Get-ChildItem -Path Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $CertificateThumbprint }
    }
    if (-not $certFound) {
        Write-Host "Certificate with thumbprint '$CertificateThumbprint' not found. Falling back to delegated authentication." -ForegroundColor Yellow
        $useAppOnly = $false
    }
    else {
        $connectParameters = @{
            ClientId              = $ClientId
            TenantId              = $TenantId
            CertificateThumbprint = $CertificateThumbprint
            NoWelcome             = $true
            ContextScope          = 'Process'
        }
        Connect-MgGraph @connectParameters
    }
}

if (-not $useAppOnly) {
    $context = Get-MgContext
    $hasRequiredScope = $context -and ($context.Scopes -contains $requiredScope)
    $tenantMatches = (-not $TenantId) -or (
        $context -and (
            ($context.TenantId -eq $TenantId) -or
            ($context.Account -like "*@$TenantId")
        )
    )

    if (-not $hasRequiredScope -or -not $tenantMatches) {
        $connectParameters = @{
            Scopes        = $requiredScope
            NoWelcome     = $true
            UseDeviceCode = $true
            ContextScope  = 'Process'
        }
        if ($TenantId) {
            $connectParameters.TenantId = $TenantId
        }

        Write-Host 'Complete the device-code sign-in within two minutes.' -ForegroundColor Yellow
        Connect-MgGraph @connectParameters
    }
}

$rows = @(Import-Csv -LiteralPath $CsvPath)
if ($rows.Count -eq 0) {
    throw 'The input CSV contains no data rows.'
}

$requiredColumns = @('UserPrincipalName', 'MobilePhone', 'AuthenticationEmail')
$actualColumns = @($rows[0].PSObject.Properties.Name)
$missingColumns = @($requiredColumns | Where-Object { $_ -notin $actualColumns })
if ($missingColumns.Count) {
    throw "The CSV is missing required column(s): $($missingColumns -join ', ')"
}

Write-Host "`nProcessing $($rows.Count) users..." -ForegroundColor Cyan
if (-not $Apply) {
    Write-Host "PREVIEW MODE - No changes will be made. Use -Apply to commit changes.`n" -ForegroundColor Yellow
}

$total = $rows.Count
$current = 0
$successCount = 0
$failCount = 0
$skipCount = 0
$previewCount = 0

$results = foreach ($row in $rows) {
    $current++
    
    $upn = ([string]$row.UserPrincipalName).Trim()
    $mobilePhone = ([string]$row.MobilePhone).Trim()
    $authenticationEmail = ([string]$row.AuthenticationEmail).Trim()

    # Update progress bar
    $percentComplete = ($current / $total) * 100
    $statusLine = if ($Apply) {
        "Success: $successCount | Failed: $failCount | Skipped: $skipCount"
    } else {
        "Preview: $previewCount | Skipped: $skipCount"
    }
    
    Write-Progress -Activity "Processing Entra ID Authentication Methods" `
        -Status "User $current of $total - $statusLine" `
        -PercentComplete $percentComplete `
        -CurrentOperation "Processing $upn"

    if ([string]::IsNullOrWhiteSpace($upn)) {
        $skipCount++
        [pscustomobject]@{
            UserPrincipalName = ''
            Method            = ''
            Value             = ''
            Action            = 'Skipped'
            Status            = 'Error'
            Message           = 'UserPrincipalName is blank.'
        }
        continue
    }

    $encodedUser = [uri]::EscapeDataString($upn)
    $currentMethod = 'row'

    try {
        if ($mobilePhone) {
            $currentMethod = 'mobilePhone'
            if ($mobilePhone -notmatch '^\+\d{1,3} \d+(x\d+)?$') {
                throw "Invalid mobile-phone format '$mobilePhone'. Use +{country code} {number}, for example +1 9185551234."
            }
            $phoneResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$encodedUser/authentication/phoneMethods"
            $existingMobile = @($phoneResponse.value | Where-Object { $_.phoneType -eq 'mobile' }) | Select-Object -First 1
            $action = if ($existingMobile) { 'Update' } else { 'Add' }

            if (-not $Apply) {
                $previewCount++
                [pscustomobject]@{
                    UserPrincipalName = $upn
                    Method            = 'mobilePhone'
                    Value             = $mobilePhone
                    Action            = $action
                    Status            = 'Preview'
                    Message           = 'No change made. Run with -Apply to commit.'
                }
            }
            else {
                $changeApplied = $PSCmdlet.ShouldProcess($upn, "$action mobile authentication phone $mobilePhone")
                
                if ($changeApplied) {
                    $body = @{ phoneNumber = $mobilePhone; phoneType = 'mobile' } | ConvertTo-Json
                    if ($existingMobile) {
                        $phoneMethodId = [uri]::EscapeDataString([string]$existingMobile.id)
                        Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$encodedUser/authentication/phoneMethods/$phoneMethodId" -Body $body -ContentType 'application/json' | Out-Null
                    }
                    else {
                        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$encodedUser/authentication/phoneMethods" -Body $body -ContentType 'application/json' | Out-Null
                    }

                    $successCount++
                    [pscustomobject]@{
                        UserPrincipalName = $upn
                        Method            = 'mobilePhone'
                        Value             = $mobilePhone
                        Action            = $action
                        Status            = 'Success'
                        Message           = 'Authentication method saved.'
                    }
                }
                else {
                    $skipCount++
                    [pscustomobject]@{
                        UserPrincipalName = $upn
                        Method            = 'mobilePhone'
                        Value             = $mobilePhone
                        Action            = $action
                        Status            = 'WhatIf'
                        Message           = "Would $action authentication method."
                    }
                }
            }
        }
        else {
            # Blank MobilePhone = remove the existing mobile authentication method.
            $currentMethod = 'mobilePhone'
            $phoneResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$encodedUser/authentication/phoneMethods"
            $existingMobile = @($phoneResponse.value | Where-Object { $_.phoneType -eq 'mobile' }) | Select-Object -First 1

            if (-not $existingMobile) {
                $skipCount++
                [pscustomobject]@{
                    UserPrincipalName = $upn
                    Method            = 'mobilePhone'
                    Value             = ''
                    Action            = 'None'
                    Status            = 'No method'
                    Message           = 'No mobile authentication method to remove.'
                }
            }
            elseif (-not $Apply) {
                $previewCount++
                [pscustomobject]@{
                    UserPrincipalName = $upn
                    Method            = 'mobilePhone'
                    Value             = ''
                    Action            = 'Delete'
                    Status            = 'Preview'
                    Message           = 'No change made. Run with -Apply to remove.'
                }
            }
            else {
                $changeApplied = $PSCmdlet.ShouldProcess($upn, "Remove mobile authentication phone method")

                if ($changeApplied) {
                    $phoneMethodId = [uri]::EscapeDataString([string]$existingMobile.id)
                    Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/users/$encodedUser/authentication/phoneMethods/$phoneMethodId" | Out-Null

                    $successCount++
                    [pscustomobject]@{
                        UserPrincipalName = $upn
                        Method            = 'mobilePhone'
                        Value             = ''
                        Action            = 'Delete'
                        Status            = 'Success'
                        Message           = 'Authentication method removed.'
                    }
                }
                else {
                    $skipCount++
                    [pscustomobject]@{
                        UserPrincipalName = $upn
                        Method            = 'mobilePhone'
                        Value             = ''
                        Action            = 'Delete'
                        Status            = 'WhatIf'
                        Message           = 'Would remove authentication method.'
                    }
                }
            }
        }

        if ($authenticationEmail) {
            $currentMethod = 'email'
            try {
                $null = [System.Net.Mail.MailAddress]::new($authenticationEmail)
            }
            catch {
                throw "Invalid authentication-email format '$authenticationEmail'."
            }
            $emailResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$encodedUser/authentication/emailMethods"
            $existingEmail = @($emailResponse.value) | Select-Object -First 1
            $action = if ($existingEmail) { 'Update' } else { 'Add' }

            if (-not $Apply) {
                $previewCount++
                [pscustomobject]@{
                    UserPrincipalName = $upn
                    Method            = 'email'
                    Value             = $authenticationEmail
                    Action            = $action
                    Status            = 'Preview'
                    Message           = 'No change made. Run with -Apply to commit.'
                }
            }
            else {
                $changeApplied = $PSCmdlet.ShouldProcess($upn, "$action authentication email $authenticationEmail")
                
                if ($changeApplied) {
                    $body = @{ emailAddress = $authenticationEmail } | ConvertTo-Json
                    if ($existingEmail) {
                        $emailMethodId = [uri]::EscapeDataString([string]$existingEmail.id)
                        Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$encodedUser/authentication/emailMethods/$emailMethodId" -Body $body -ContentType 'application/json' | Out-Null
                    }
                    else {
                        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$encodedUser/authentication/emailMethods" -Body $body -ContentType 'application/json' | Out-Null
                    }

                    $successCount++
                    [pscustomobject]@{
                        UserPrincipalName = $upn
                        Method            = 'email'
                        Value             = $authenticationEmail
                        Action            = $action
                        Status            = 'Success'
                        Message           = 'Authentication method saved.'
                    }
                }
                else {
                    $skipCount++
                    [pscustomobject]@{
                        UserPrincipalName = $upn
                        Method            = 'email'
                        Value             = $authenticationEmail
                        Action            = $action
                        Status            = 'WhatIf'
                        Message           = "Would $action authentication method."
                    }
                }
            }
        }

    }
    catch {
        $failCount++
        $errorMessage = $_.Exception.Message
        if ($_.ErrorDetails.Message) {
            $errorMessage = "$errorMessage | Graph details: $($_.ErrorDetails.Message)"
        }

        [pscustomobject]@{
            UserPrincipalName = $upn
            Method            = $currentMethod
            Value             = ''
            Action            = 'Failed'
            Status            = 'Error'
            Message           = $errorMessage
        }
    }
}

Write-Progress -Activity "Processing Entra ID Authentication Methods" -Completed

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Processing Complete" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total users processed: $total" -ForegroundColor White

if ($Apply) {
    Write-Host "Successfully updated: $successCount" -ForegroundColor Green
    Write-Host "Failed: $failCount" -ForegroundColor Red
    Write-Host "Skipped: $skipCount" -ForegroundColor Yellow
} else {
    Write-Host "Preview operations: $previewCount" -ForegroundColor Yellow
    Write-Host "Skipped: $skipCount" -ForegroundColor Yellow
}
Write-Host "========================================`n" -ForegroundColor Cyan

if ($LogPath) {
    $resolvedLogPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($LogPath)
    $logDirectory = Split-Path -Parent $resolvedLogPath
    if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory)) {
        New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
    }
    $results | Export-Csv -LiteralPath $resolvedLogPath -NoTypeInformation -Encoding UTF8
    Write-Host "Results exported to: $resolvedLogPath" -ForegroundColor Green
}

$results | Format-Table -AutoSize