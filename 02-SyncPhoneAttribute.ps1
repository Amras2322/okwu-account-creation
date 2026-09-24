<#
   Pipeline to normalize phone numbers (extensionAttribute2 -> extensionAttribute10)
   and export both phone numbers and personal emails (extensionAttribute1) as
   Entra ID authentication methods.

   extensionAttribute2 is the source of truth for a user's personal mobile
   number (raw, as entered). extensionAttribute10 stores the normalized form
   ("+1 XXXXXXXXXX") that is exported to Entra ID. The AD "mobile" attribute is
   no longer used by this pipeline.

   Runs as a DELTA: only users whose account changed within the last
   -HoursBack hour(s) are evaluated (default 1 hour). Optionally scope to a
   specific OU with -SearchBase; without it the whole domain is searched.

   For every changed user with extensionAttribute2 or extensionAttribute10
   populated, this script will:
       1. Read extensionAttribute2 and strip all non-digit characters.
       2. If more than 10 digits remain, keep only the last 10 (drops leading
          country code / stray digits).
       3. If exactly 10 digits remain, format as "+1 XXXXXXXXXX" and write it to
          extensionAttribute10 ONLY if it differs (idempotent - safe to re-run
          without clobbering an already-correct value).
       4. If extensionAttribute2 is blank, clear extensionAttribute10. This
          exports a blank phone so Script 03 removes the Entra phone method.
       5. If extensionAttribute2 is malformed (not 10 digits), leave
          extensionAttribute10 (and the existing Entra method) untouched and
          send a warning email to support naming the affected users.
       6. Log every change and skip; write nothing if the value is already
          correct.

   Users with extensionAttribute1 (personal email) populated will also be
   exported to the CSV for Entra ID authentication-email import.

   Use -FullSync to process ALL users with phone/email data regardless of
   when they last changed. This is useful for initial setup or periodic
   full reconciliations. NOTE: a full sync will cause Script 03 to remove the
   phone authentication method for every user whose extensionAttribute2 is blank.


EXAMPLE USE (Delta past hour, whole domain, dry run)
    .\02-SyncPhoneAttribute.ps1 -WhatIf

EXAMPLE USE (Delta scoped to an OU)
    .\02-SyncPhoneAttribute.ps1 -SearchBase "OU=Bartlesville,DC=okwu,DC=edu" -HoursBack 1

EXAMPLE USE (Full sync, whole domain)
    .\02-SyncPhoneAttribute.ps1 -FullSync

EXAMPLE USE (Full sync, specific OU, dry run)
    .\02-SyncPhoneAttribute.ps1 -FullSync -SearchBase "OU=Bartlesville,DC=okwu,DC=edu" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$SearchBase,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 8760)]
    [int]$HoursBack = 1,

    [Parameter(Mandatory = $false)]
    [string]$CsvExportPath,

    [Parameter(Mandatory = $false)]
    [switch]$FullSync,

    [Parameter(Mandatory = $false)]
    [string]$SupportEmail = 'support@okwu.edu',

    [Parameter(Mandatory = $false)]
    [string]$FromAddress = 'entrasync@okwu.edu',

    [Parameter(Mandatory = $false)]
    [string]$SmtpServer = '10.0.5.16'
)

Import-Module ActiveDirectory

$cutoffDate = (Get-Date).ToUniversalTime().AddHours(-$HoursBack)
$cutoffLdap = $cutoffDate.ToString('yyyyMMddHHmmss.0Z')

$phoneFilter = "(|(extensionAttribute1=*)(extensionAttribute2=*)(extensionAttribute10=*))"
if ($FullSync) {
    $ldapFilter = "(&(objectCategory=person)(objectClass=user)$phoneFilter)"
}
else {
    $ldapFilter = "(&(objectCategory=person)(objectClass=user)$phoneFilter(whenChanged>=$cutoffLdap))"
}

$getParameters = @{
    LDAPFilter = $ldapFilter
    Properties = 'extensionAttribute1', 'extensionAttribute2', 'extensionAttribute10', 'UserPrincipalName'
}
if ($SearchBase) {
    $getParameters.SearchBase  = $SearchBase
    $getParameters.SearchScope = 'Subtree'
}

$users = Get-ADUser @getParameters

if (-not $users) {
    if ($FullSync) {
        Write-Host "`nNo users found with phone or email data. Nothing to do." -ForegroundColor Yellow
    }
    else {
        Write-Host "`nNo users changed within the last $HoursBack hour(s). Nothing to do." -ForegroundColor Yellow
    }
    return
}

$results        = [System.Collections.Generic.List[object]]::new()
$exportRows     = [System.Collections.Generic.List[object]]::new()
$malformedUsers = [System.Collections.Generic.List[object]]::new()

foreach ($user in $users) {

    # --- Phone processing ---
    $extAtt2       = $user.extensionAttribute2
    $phoneOriginal = $user.extensionAttribute10
    $phoneFinal    = $phoneOriginal

    if ([string]::IsNullOrWhiteSpace($extAtt2)) {
        # extensionAttribute2 is the source of truth. A blank value means the
        # user should have no phone, so clear extensionAttribute10 and let
        # Script 03 remove the Entra authentication method.
        if ([string]::IsNullOrWhiteSpace($phoneOriginal)) {
            $phoneFinal = $null
            $results.Add([PSCustomObject]@{
                SamAccountName = $user.SamAccountName
                Attribute      = "Phone"
                OriginalValue  = $phoneOriginal
                NewValue       = $null
                Status         = "NO CHANGE"
            })
        }
        else {
            $changeApplied = $PSCmdlet.ShouldProcess(
                $user.SamAccountName,
                "Clear extensionAttribute10 (extensionAttribute2 is blank)"
            )

            if ($changeApplied) {
                Set-ADUser -Identity $user -Clear 'extensionAttribute10'
            }

            $phoneFinal = if ($changeApplied) { $null } else { $phoneOriginal }
            $results.Add([PSCustomObject]@{
                SamAccountName = $user.SamAccountName
                Attribute      = "Phone"
                OriginalValue  = $phoneOriginal
                NewValue       = $null
                Status         = if ($changeApplied) { "CLEARED (extensionAttribute2 blank)" } else { "WHATIF - WOULD CLEAR" }
            })
        }
    }
    else {
        # Step 2: strip non-digits, keep only the last 10 if longer.
        $digitsOnly = $extAtt2 -replace '\D', ''
        if ($digitsOnly.Length -gt 10) {
            $digitsOnly = $digitsOnly.Substring($digitsOnly.Length - 10)
        }

        if ($digitsOnly.Length -ne 10) {
            # Malformed source value: leave extensionAttribute10 (and therefore
            # the existing Entra method) untouched and alert support.
            $malformedUsers.Add([PSCustomObject]@{
                SamAccountName      = $user.SamAccountName
                UserPrincipalName   = $user.UserPrincipalName
                extensionAttribute2 = $extAtt2
            })
            $results.Add([PSCustomObject]@{
                SamAccountName = $user.SamAccountName
                Attribute      = "Phone"
                OriginalValue  = $phoneOriginal
                NewValue       = $null
                Status         = "SKIPPED - invalid/short data ($extAtt2)"
            })
        }
        else {
            # Step 4: format as "+1 XXXXXXXXXX"
            $phoneNew = "+1 $digitsOnly"

            if ($phoneNew -eq $phoneOriginal) {
                $phoneFinal = $phoneOriginal
                $results.Add([PSCustomObject]@{
                    SamAccountName = $user.SamAccountName
                    Attribute      = "Phone"
                    OriginalValue  = $phoneOriginal
                    NewValue       = $phoneNew
                    Status         = "NO CHANGE"
                })
            }
            else {
                $changeApplied = $PSCmdlet.ShouldProcess(
                    $user.SamAccountName,
                    "Set extensionAttribute10 to '$phoneNew'"
                )

                if ($changeApplied) {
                    Set-ADUser -Identity $user -Replace @{
                        extensionAttribute10 = $phoneNew
                    }
                }

                $phoneFinal = if ($changeApplied) { $phoneNew } else { $phoneOriginal }
                $results.Add([PSCustomObject]@{
                    SamAccountName = $user.SamAccountName
                    Attribute      = "Phone"
                    OriginalValue  = $phoneOriginal
                    NewValue       = $phoneNew
                    Status         = if ($changeApplied) { "UPDATED" } else { "WHATIF - WOULD UPDATE" }
                })
            }
        }
    }

    # --- Email processing ---
    $email = $user.extensionAttribute1
    $hasValidEmail = -not [string]::IsNullOrWhiteSpace($email) -and $email -match '@'

    $results.Add([PSCustomObject]@{
        SamAccountName = $user.SamAccountName
        Attribute      = "Email"
        OriginalValue  = if ($hasValidEmail) { $email } else { $null }
        NewValue       = if ($hasValidEmail) { $email } else { $null }
        Status         = if ($hasValidEmail) { "EXPORTED" } else { "SKIPPED" }
    })

    # --- Export row (a blank phone signals Script 03 to remove the method) ---
    if ($user.UserPrincipalName) {
        $exportRows.Add([PSCustomObject]@{
            UserPrincipalName   = $user.UserPrincipalName
            MobilePhone         = if (-not [string]::IsNullOrWhiteSpace($phoneFinal)) { $phoneFinal } else { '' }
            AuthenticationEmail = if ($hasValidEmail) { $email } else { '' }
        })
    }
}

Write-Host "`nResults:" -ForegroundColor Cyan
$results | Format-Table -AutoSize

# Warn support about malformed phone data (extensionAttribute10 left untouched)
if ($malformedUsers.Count -gt 0) {
    Write-Host "`n$($malformedUsers.Count) user(s) have malformed extensionAttribute2 data:" -ForegroundColor Yellow
    $malformedUsers | Format-Table SamAccountName, UserPrincipalName, extensionAttribute2 -AutoSize

    if (-not $WhatIfPreference) {
        $lines = $malformedUsers | ForEach-Object {
            " - $($_.SamAccountName) ($($_.UserPrincipalName)): '$($_.extensionAttribute2)'"
        }
        $body = @"
The following users have malformed extensionAttribute2 values. A valid phone
number needs exactly 10 digits after non-digit characters are removed.

extensionAttribute10 was left unchanged and any existing Entra ID phone
authentication method was preserved until the data is corrected.

$($lines -join "`n")
"@

        try {
            Send-MailMessage `
                -To $SupportEmail `
                -From $FromAddress `
                -Subject "WARNING: $($malformedUsers.Count) user(s) with malformed phone number in extensionAttribute2" `
                -Body $body `
                -SmtpServer $SmtpServer
            Write-Host "Warning email sent to $SupportEmail" -ForegroundColor Cyan
        }
        catch {
            Write-Warning "Failed to send malformed-phone warning email: $($_.Exception.Message)"
        }
    }
}

# Export to CSV for Entra ID import (only if not running -WhatIf)
if (-not $WhatIfPreference) {
    if (-not $CsvExportPath) {
        $CsvExportPath = ".\EntraMethods.csv"
    }

    Write-Host "`nExporting users to CSV..." -ForegroundColor Cyan

    $exportRows | Export-Csv `
        -Path $CsvExportPath `
        -NoTypeInformation `
        -Encoding UTF8

    Write-Host "Exported $($exportRows.Count) users to: $CsvExportPath" -ForegroundColor Green
    Write-Host "`nNote: AuthenticationEmail is populated from extensionAttribute1 (personal email)." -ForegroundColor Yellow
    Write-Host "Note: A blank MobilePhone tells Script 03 to remove the Entra phone method." -ForegroundColor Yellow
}
else {
    Write-Host "`n-WhatIf specified: CSV export skipped." -ForegroundColor Yellow
}
