<#
    Standalone readmit reactivator. Runs independently of the hourly
    01-CreateStudentAccounts.ps1 scheduled task.

    Reads NewStudentsAccounts-ReAdmit.csv and, for each row:
        1. Looks up the existing account by SamAccountName (Given.Surname).
           Missing accounts are reported to support@okwu.edu and skipped.
        2. Enables the account if it is disabled (passwords are never touched;
           students recover access through the MyOKWU forgot-password flow).
        3. Moves the account to the OU mapped from the department column
           (same mapping as 01-CreateStudentAccounts.ps1). Readmits normally
           sit in OU=Alumni-WebAccessOnly,OU=Alumni.
        4. Refreshes AD attributes from the CSV: extensionAttribute1
           (personal email), extensionAttribute2 (raw mobile - 02 derives
           extensionAttribute10 from it), Department, employeeID (CSV is the
           source of truth; fixes zero-padded IDs so 04's AD lookup matches)
           and employeeNumber (zeros stripped, for PaperCut).
        5. Reconciles student security groups to exactly what 01 would grant
           for the department: adds missing groups, removes student groups
           inherited from a previous department. Groups outside the managed
           set are never touched.

    Then runs the rest of the pipeline against the readmit CSV:
        02-SyncPhoneAttribute.ps1  (delta, past hour -> EntraMethods.csv)
        03-ImportEntraMethods.ps1  (only if 02 exported rows)
        04-ProvisionJenzabar.ps1   (-CsvPath pointing at the readmit CSV)

EXAMPLE USE (dry run - no AD writes, no pipeline run)
    .\05-ReAdmitAccounts.ps1 -WhatIf

EXAMPLE USE (real run)
    .\05-ReAdmitAccounts.ps1
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param()

Import-Module ActiveDirectory

$time = Get-Date -Format MMddyyyyHHmm
Start-Transcript -Path "C:\Account_Creation\Logs\ReAdmitAccounts_Output-$time.log" -NoClobber -IncludeInvocationHeader -WhatIf:$false

$ReAdmitCsv = "C:\Account_Creation\NewStudentsAccounts-ReAdmit.csv"
$ADUsers = Import-Csv -Path $ReAdmitCsv -Delimiter ","

# Student groups this script is allowed to add/remove. Anything else on the
# account (legacy/alumni/distribution groups) is left alone.
$managedGroups = @(
    "Concurrent",
    "TradStudents",
    "AGSStudents",
    "OnlineStudents",
    "CROSSTraining",
    "StudentWireless",
    "StudentGroup"
)

$missingCount = 0
$movedCount   = 0
$attrCount    = 0
$enableCount  = 0
$groupCount   = 0
$failedCount  = 0

foreach ($User in $ADUsers) {
    try {

        $SamAccountName = "$($User.givenName).$($User.SurName)"
        $Name = "$($User.GivenName) $($User.Surname)"

        Write-Host "[******************************************************************************]"
        Write-Host "Processing $Name ($SamAccountName)"

        # Assign the OU based on the department (same mapping as 01)
        if ($User.Department -eq "Concurrent Student") {
            $OU = "OU=Concurrent Students,OU=Traditional,OU=Students,DC=okwu,DC=edu"
        }
        elseif ($User.Department -eq "TRAD Student") {
            $OU = "OU=Traditional,OU=Students,DC=okwu,DC=edu"
        }
        elseif ($User.Department -eq "GPS Student") {
            $OU = "OU=AGS,OU=Students,DC=okwu,DC=edu"
        }
        elseif ($User.Department -eq "NonDegree Student") {
            $OU = "OU=NonDegree,OU=Students,DC=okwu,DC=edu"
        }
        elseif ($User.Department -eq "Alumni") {
            $OU = "OU=Previous Students-Alumni,OU=Students,DC=okwu,DC=edu"
        }
        elseif ($User.Department -eq "Online Student") {
            $OU = "OU=Online Students,DC=okwu,DC=edu"
        }
        elseif ($User.Department -eq "Cross Student") {
            $OU = "OU=Cross Training,OU=Online Students,DC=okwu,DC=edu"
        }
        else {
            # Default OU if department does not match specific OUs
            $OU = "OU=Students,DC=okwu,DC=edu"
        }

        # Desired security groups for the department (same mapping as 01)
        $desiredGroups = @("StudentGroup")
        if ($User.Department -eq "Concurrent Student") {
            $desiredGroups += "Concurrent"
            $desiredGroups += "StudentWireless"
        }
        elseif ($User.Department -eq "TRAD Student") {
            $desiredGroups += "TradStudents"
            $desiredGroups += "StudentWireless"
        }
        elseif ($User.Department -eq "GPS Student") {
            $desiredGroups += "AGSStudents"
            $desiredGroups += "StudentWireless"
        }
        elseif ($User.Department -eq "Online Student") {
            $desiredGroups += "OnlineStudents"
        }
        elseif ($User.Department -eq "Cross Student") {
            $desiredGroups += "CROSSTraining"
        }

        # Look up the existing account
        $existing = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } `
            -Properties Department, extensionAttribute1, extensionAttribute2, employeeID, employeeNumber, MemberOf, Enabled

        if (-not $existing) {
            Write-Host "No AD account found for $SamAccountName. Sending email to Support." -ForegroundColor Yellow
            $missingCount++
            if ($PSCmdlet.ShouldProcess($SamAccountName, "Send 'user does not exist' alert to support@okwu.edu")) {
                Send-MailMessage `
                    -To "support@okwu.edu" `
                    -From "accounts@okwu.edu" `
                    -Subject "Readmit user does not exist in AD - $SamAccountName" `
                    -Body "Readmit CSV ($ReAdmitCsv) contains $Name ($SamAccountName) but no matching AD account was found. Please review." `
                    -SmtpServer "10.0.5.16"
            }
            continue
        }

        # Re-enable disabled accounts. Passwords are left untouched.
        if (-not $existing.Enabled) {
            if ($PSCmdlet.ShouldProcess($SamAccountName, "Enable-ADAccount")) {
                Enable-ADAccount -Identity $SamAccountName
                Write-Host "Enabled disabled account $SamAccountName" -ForegroundColor Green
                $enableCount++
            }
            elseif ($WhatIfPreference) {
                $enableCount++
            }
        }

        # Move to the OU mapped from the department
        $currentOU = ($existing.DistinguishedName -replace '^[^,]+,', '')
        if ($currentOU -ne $OU) {
            if ($PSCmdlet.ShouldProcess($SamAccountName, "Move from $currentOU to $OU")) {
                Move-ADObject -Identity $existing.DistinguishedName -TargetPath $OU
                Write-Host "Moved $SamAccountName to $OU" -ForegroundColor Green
                $movedCount++
            }
            elseif ($WhatIfPreference) {
                $movedCount++
            }
        }

        # Refresh attributes from the CSV (only write what actually differs)
        $attrChanges = @{}
        if ($existing.extensionAttribute1 -ne $User.PersonalEmail) {
            $attrChanges["extensionAttribute1"] = $User.PersonalEmail
        }
        if ($existing.extensionAttribute2 -ne $User.MobilePhone) {
            $attrChanges["extensionAttribute2"] = $User.MobilePhone
        }
        if ($existing.Department -ne $User.Department) {
            $attrChanges["Department"] = $User.Department
        }
        if ($existing.employeeID -ne $User.employeeID) {
            $attrChanges["employeeID"] = $User.employeeID
        }
        $employeeNumber = $User.employeeID -replace '^0+', ''
        if ($existing.employeeNumber -ne $employeeNumber) {
            $attrChanges["employeeNumber"] = $employeeNumber
        }

        if ($attrChanges.Count -gt 0) {
            if ($PSCmdlet.ShouldProcess($SamAccountName, "Set-ADUser -Replace $($attrChanges.Keys -join ', ')")) {
                Set-ADUser -Identity $SamAccountName -Replace $attrChanges
                Write-Host "Updated attributes for ${SamAccountName}: $($attrChanges.Keys -join ', ')" -ForegroundColor Green
                $attrCount++
            }
            elseif ($WhatIfPreference) {
                $attrCount++
            }
        }

        # Reconcile student security groups with the department mapping
        $currentGroups = @($existing.MemberOf | ForEach-Object { ($_ -split ',')[0] -replace '^CN=' })

        $groupsToAdd = @($desiredGroups | Where-Object { $currentGroups -notcontains $_ })
        $groupsToRemove = @($currentGroups | Where-Object { ($managedGroups -contains $_) -and ($desiredGroups -notcontains $_) })

        foreach ($group in $groupsToAdd) {
            if ($PSCmdlet.ShouldProcess($SamAccountName, "Add-ADGroupMember $group")) {
                Add-ADGroupMember -Identity $group -Members $SamAccountName
                Write-Host "Added $SamAccountName to $group" -ForegroundColor Green
                $groupCount++
            }
            elseif ($WhatIfPreference) {
                $groupCount++
            }
        }
        foreach ($group in $groupsToRemove) {
            if ($PSCmdlet.ShouldProcess($SamAccountName, "Remove-ADGroupMember $group")) {
                Remove-ADGroupMember -Identity $group -Members $SamAccountName -Confirm:$false
                Write-Host "Removed $SamAccountName from $group" -ForegroundColor Green
                $groupCount++
            }
            elseif ($WhatIfPreference) {
                $groupCount++
            }
        }

        if ($groupsToAdd.Count -eq 0 -and $groupsToRemove.Count -eq 0) {
            Write-Host "Group membership already matches $($User.Department) for $SamAccountName"
        }

    }
    catch {
        $failedCount++
        Write-Host "Failed to process $SamAccountName - $_" -ForegroundColor Red
    }
}


write-host "
-
-
-
-
-
Readmit summary: $missingCount missing, $movedCount moved, $attrCount attribute updates, $enableCount enabled, $groupCount group changes, $failedCount failed.
-
-
-
-
-"


if ($WhatIfPreference) {
    write-host "
-
-
-
-
-
#-#-#-#-#-# WhatIf was used - skipping Entra sync and the 02/03/04 pipeline. #-#-#-#-#-#
-
-
-
-
-"
    try { Stop-Transcript } catch { Write-Warning "Could not stop transcript: $($_.Exception.Message)" }
    return
}


#Take a quick nap
sleep 5

#Run delta sync so enabled-state and refreshed attributes reach Entra.
#A sync failure (e.g. access denied when not elevated) must not abort the run -
#02/03/04 do not depend on the sync cycle.
try {
    Start-ADSyncSyncCycle -PolicyType delta -ErrorAction Stop
}
catch {
    Write-Warning "Start-ADSyncSyncCycle failed: $($_.Exception.Message)"
    Write-Warning "Continuing without the sync cycle - trigger a delta sync manually if needed."
}

#Readmit accounts already exist in Entra, so only a short settle delay is needed.
Start-Sleep -Seconds 30


#Run phone normalization and Entra ID authentication-method import (delta, past hour).
#Each script runs synchronously, so nothing proceeds until the previous one finishes.
cd C:\Account_Creation

$PhoneCsv = "C:\Account_Creation\EntraMethods.csv"

# Remove any stale CSV so only this run's output is imported
Remove-Item -LiteralPath $PhoneCsv -ErrorAction SilentlyContinue

write-host "
-
-
-
-
-
#-#-#-#-#-# Now starting the 02-SyncPhoneAttribute.ps1 script. #-#-#-#-#-#
-
-
-
-
-"
.\02-SyncPhoneAttribute.ps1 -HoursBack 1 -CsvExportPath $PhoneCsv

# Only import if this run produced rows
if ((Test-Path -LiteralPath $PhoneCsv) -and @(Import-Csv -LiteralPath $PhoneCsv).Count -gt 0) {
    write-host "
-
-
-
-
-
#-#-#-#-#-# Now starting the 03-ImportEntraMethods.ps1 script. #-#-#-#-#-#
-
-
-
-
-"
    .\03-ImportEntraMethods.ps1 `
        -CsvPath $PhoneCsv `
        -Apply `
        -Confirm:$false
}
else {
    Write-Host "No phone methods to import (02-SyncPhoneAttribute.ps1 exported no users)." -ForegroundColor Yellow
}


#Trigger the next script to run
write-host "
-
-
-
-
-
#-#-#-#-#-# Now starting the 04-ProvisionJenzabar.ps1 script. The log for that starts now. #-#-#-#-#-#
-
-
-
-
-"
cd C:\Account_Creation
.\04-ProvisionJenzabar.ps1 -CsvPath $ReAdmitCsv


Start-Sleep -Seconds 5


# End logging
Write-Host "Readmit processing complete."
cd C:\Account_Creation
try { Stop-Transcript } catch { Write-Warning "Could not stop transcript: $($_.Exception.Message)" }
