# Import Active Directory module for running AD cmdlets
Import-Module ActiveDirectory
Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -Force

#Start the logs
$time = Get-Date -Format MMddyyyyHHmm
Start-Transcript -Path "C:\Account_Creation\Logs\CreateStudentAccounts_Output-$time.log" -NoClobber -IncludeInvocationHeader

# Store the data from a csv in the $ADUsers variable
$ADUsers = Import-Csv "C:\Account_Creation\NewStudentsAccounts-auto.csv" -Delimiter ","

# Define User Principle Name
$UPN = "okwustudents.edu"

# Track the UPNs created in this run so we can wait for AAD Connect
# to provision them in Entra before importing phone methods below.
$newUpnList = @()

# --- Username collision handling -------------------------------------------------
# Collected during the run and reported to support at the end so that a name
# collision is never a silent drop.
$collisions        = [System.Collections.Generic.List[object]]::new()
$readmitCandidates = [System.Collections.Generic.List[object]]::new()
$reviewCandidates  = [System.Collections.Generic.List[object]]::new()

# Names handed out earlier in this run. Get-ADUser can miss an account created
# moments ago if the write and the read land on different domain controllers,
# so keep our own list of names that are already spoken for.
$reservedSamNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

# Employee IDs are compared with leading zeros stripped: the CSV sends
# "190699" while AD may hold "000190699".
function Get-NormalizedEmployeeId {
    param([string]$EmployeeId)
    if ([string]::IsNullOrWhiteSpace($EmployeeId)) { return $null }
    $normalized = ($EmployeeId.Trim() -replace '^0+', '')
    if ([string]::IsNullOrWhiteSpace($normalized)) { return $null }
    return $normalized
}

# The IDs an AD account can be identified by. employeeNumber is written by this
# script (employeeID with the zeros stripped), so check both.
function Get-ADUserIdentityIds {
    param($ADUser)
    $ids = @()
    foreach ($value in @($ADUser.employeeID, $ADUser.employeeNumber)) {
        $normalized = Get-NormalizedEmployeeId $value
        if ($normalized -and ($ids -notcontains $normalized)) { $ids += $normalized }
    }
    return $ids
}

# AD may hold the employeeID unpadded (190699) or zero-padded (000190699) -
# the same forms 04/05 search on.
function Get-EmployeeIdForms {
    param([string]$NormalizedId)
    $forms = @($NormalizedId)
    if ($NormalizedId -match '^\d+$') {
        $forms += ('{0:D9}' -f [int]$NormalizedId)
    }
    return ($forms | Select-Object -Unique)
}

# A candidate username is usable only if the SamAccountName, the UPN and the CN
# in the target OU are all free.
function Test-ADAccountNameAvailable {
    param(
        [string]$SamAccountName,
        [string]$UserPrincipalName,
        [string]$Name,
        [string]$OU
    )

    if ($reservedSamNames.Contains($SamAccountName)) { return $false }
    if (Get-ADUser -Filter { SamAccountName -eq $SamAccountName }) { return $false }
    if (Get-ADUser -Filter { UserPrincipalName -eq $UserPrincipalName }) { return $false }

    # CN only has to be unique inside the target OU, but two students with the
    # same name in the same OU would otherwise break New-ADUser.
    $cn = $Name -replace '([\\,+"<>;=])', '\$1'
    try {
        $existingObject = Get-ADObject -Identity "CN=$cn,$OU" -ErrorAction Stop
    }
    catch {
        $existingObject = $null
    }
    return (-not $existingObject)
}

# Loop through each row containing user details in the CSV file
foreach ($User in $ADUsers) {
    try {

        # Assign the OU based on the department
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
        

        # Resolve the username for this row.
        #
        # Identity is employeeID, not the name: two students can share a name, and
        # a student can already own a numbered username from an earlier run.
        #
        #   1. If any account already carries this employeeID, the student already
        #      has an account - never create a second one. Report it so a returning
        #      student's row can go through 05-ReAdmitAccounts.ps1.
        #   2. Otherwise, if the base username (GivenName.Surname) is taken, the
        #      holder is a namesake - a different person. Append a number:
        #      John.Smith, John.Smith1, John.Smith2, ...
        #   3. If the existing holder has no employeeID to compare against, do not
        #      guess: skip and report for manual review.
        $BaseSamAccountName = "$($User.givenName).$($User.SurName)"
        $BaseName = "$($User.GivenName) $($User.Surname)"

        $SamAccountName = $BaseSamAccountName
        $Name = $BaseName

        $csvEmployeeId = Get-NormalizedEmployeeId $User.employeeID

        # Step 1: does this student already have an account, under any username?
        $idMatches = @()
        if ($csvEmployeeId) {
            foreach ($idForm in (Get-EmployeeIdForms $csvEmployeeId)) {
                $idMatches += @(Get-ADUser -Filter { employeeID -eq $idForm } `
                    -Properties employeeID, employeeNumber, DistinguishedName, Enabled, UserPrincipalName)
                $idMatches += @(Get-ADUser -Filter { employeeNumber -eq $idForm } `
                    -Properties employeeID, employeeNumber, DistinguishedName, Enabled, UserPrincipalName)
            }
            $idMatches = @($idMatches | Where-Object { $_ } | Sort-Object -Property DistinguishedName -Unique)
        }

        if ($idMatches.Count -gt 1) {
            # Duplicate IDs in AD - picking one would be a guess.
            $reviewCandidates.Add([PSCustomObject]@{
                Name           = $BaseName
                SamAccountName = (($idMatches | ForEach-Object { $_.SamAccountName }) -join ', ')
                CSVEmployeeID  = $csvEmployeeId
                ExistingDN     = (($idMatches | ForEach-Object { $_.DistinguishedName }) -join '; ')
                Reason         = "employeeID $csvEmployeeId matches $($idMatches.Count) AD accounts"
            })
            Write-Host "employeeID $csvEmployeeId matches $($idMatches.Count) AD accounts ($(($idMatches | ForEach-Object { $_.SamAccountName }) -join ', ')) - skipping $BaseName. See the collision report." -ForegroundColor Yellow
            continue
        }

        if ($idMatches.Count -eq 1) {
            $existing = $idMatches[0]
            $readmitCandidates.Add([PSCustomObject]@{
                Name           = $BaseName
                SamAccountName = $existing.SamAccountName
                EmployeeID     = $csvEmployeeId
                ExistingDN     = $existing.DistinguishedName
                Enabled        = $existing.Enabled
            })
            Write-Host "employeeID $csvEmployeeId already has account $($existing.SamAccountName) ($($existing.DistinguishedName)) - same student, not creating a duplicate. See the collision report." -ForegroundColor Yellow
            continue
        }

        # Step 2: the base username is free unless a namesake holds it.
        $existingHolder = Get-ADUser -Filter { SamAccountName -eq $BaseSamAccountName } `
            -Properties employeeID, employeeNumber, DistinguishedName, Enabled, UserPrincipalName

        if ($existingHolder) {
            $existingIds = @(Get-ADUserIdentityIds $existingHolder)

            if ($existingIds.Count -eq 0) {
                # No ID on the existing account - cannot prove it is a namesake.
                $reviewCandidates.Add([PSCustomObject]@{
                    Name           = $BaseName
                    SamAccountName = $BaseSamAccountName
                    CSVEmployeeID  = "$($User.employeeID)"
                    ExistingDN     = $existingHolder.DistinguishedName
                    Reason         = "existing account has no employeeID/employeeNumber to compare"
                })
                Write-Host "Cannot tell whether $BaseSamAccountName ($($existingHolder.DistinguishedName)) is the same student as $BaseName - it has no employeeID. Skipping this row; see the collision report." -ForegroundColor Yellow
                continue
            }

            # A different person holds the base username.
            # Walk Given.Surname1 -> Given.Surname2 -> Given.Surname3 ...
            # (the unsuffixed Given.Surname is held by the other student)
            $suffix = 1
            $resolved = $false
            while (-not $resolved) {
                $candidateSam = "$BaseSamAccountName$suffix"
                $candidateName = "$BaseName $suffix"
                if (Test-ADAccountNameAvailable -SamAccountName $candidateSam -UserPrincipalName "$candidateSam@$UPN" -Name $candidateName -OU $OU) {
                    $SamAccountName = $candidateSam
                    $Name = $candidateName
                    $resolved = $true
                }
                elseif ($suffix -ge 999) {
                    throw "No available username found for $BaseSamAccountName after 999 suffixes."
                }
                else {
                    $suffix++
                }
            }

            $collisions.Add([PSCustomObject]@{
                Name             = $BaseName
                CSVEmployeeID    = $csvEmployeeId
                BaseUsername     = $BaseSamAccountName
                AssignedUsername = $SamAccountName
                ExistingID       = ($existingIds -join ', ')
                ExistingDN       = $existingHolder.DistinguishedName
            })
            Write-Host "Username $BaseSamAccountName is held by a different student (employeeID $($existingIds -join ', ') at $($existingHolder.DistinguishedName)). Creating $BaseName as $SamAccountName." -ForegroundColor Yellow
        }

        [void]$reservedSamNames.Add($SamAccountName)

        # $Name is the real name, with a numeric suffix appended when the base
        # username was already held by a different student.

        # A compliant password has to be set or the script fails
        $Pass = "T#8uW6$mK0p&sAG2y!erB27"


        # Define parameters for New-ADUser
        $UserParams = @{
            GivenName             = $User.GivenName           # First Name
            Surname               = $User.Surname             # Last Name (Surname)
            Department            = $User.department          # Department (adjust as needed for the OU)
            EmployeeID            = $User.employeeID          # Employee ID
            EmailAddress          = "$SamAccountName@$UPN"    # Uni Email
            UserPrincipalName     = "$SamAccountName@$UPN"    # UPN (firstname.lastname@okwustudents.edu)
            SamAccountName        = $SamAccountName
            Enabled               = $true                    # Change to $True to enable accounts upon creation
            Name                  = $Name
            DisplayName           = "$($User.GivenName) $($User.Surname)"  # always the real name, never the suffix
            Country               = "US"
            Path                  = $OU                       # Determines which OU user is placed in
            AccountPassword       = (ConvertTo-SecureString $Pass -AsPlainText -Force)
            Company               = "Oklahoma Wesleyan University"
            ChangePasswordAtLogon = $true                     # Account must be enabled for this to take effect
        }

        # Debugging: Output all parameters to verify
        Write-Host "Creating user with the following parameters:"
        $UserParams.GetEnumerator() | ForEach-Object { Write-Host "$($_.Key): $($_.Value)" }

        # Safety net: the username was resolved above, but re-check right before
        # the write in case another process claimed it in the meantime.
        $raceHolder = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -Properties DistinguishedName
        if ($raceHolder) {
            $reviewCandidates.Add([PSCustomObject]@{
                Name           = $Name
                SamAccountName = $SamAccountName
                CSVEmployeeID  = "$($User.employeeID)"
                ExistingDN     = $raceHolder.DistinguishedName
                Reason         = "username became unavailable during this run"
            })
            Write-Host "A user with username $SamAccountName already exists in Active Directory - skipping this row; see the collision report." -ForegroundColor Yellow
        }
        else {
            # Create the new user account
            New-ADUser @UserParams

            # Set custom attributes after creating the user
            $customAttributes = @{
                extensionAttribute1 = $User.PersonalEmail
                extensionAttribute2 = $User.MobilePhone
            }

            # extensionAttribute2 is the source of truth for the personal phone
            # number. 02-SyncPhoneAttribute.ps1 derives/normalizes
            # extensionAttribute10 from it and writes that separately.
            Set-ADUser -Identity $SamAccountName -Add $customAttributes

            #Adding employeeNumber for PaperCut login capability
            $employeeNumber = $User.employeeID -replace '^0+', ''
            Set-ADUser -Identity $SamAccountName -Replace @{employeeNumber = $employeeNumber }
            


            # Check department and add user to the appropriate security group
            if ($User.Department -eq "Concurrent Student") {
                Add-ADGroupMember -Identity "Concurrent" -Members $SamAccountName
                Add-ADGroupMember -Identity "StudentWireless" -Members $SamAccountName
            }
            elseif ($User.Department -eq "TRAD Student") {
                Add-ADGroupMember -Identity "TradStudents" -Members $SamAccountName
                Add-ADGroupMember -Identity "StudentWireless" -Members $SamAccountName
            }
            elseif ($User.Department -eq "GPS Student") {
                Add-ADGroupMember -Identity "AGSStudents" -Members $SamAccountName
                Add-ADGroupMember -Identity "StudentWireless" -Members $SamAccountName
            }
            elseif ($User.Department -eq "Online Student") {
                Add-ADGroupMember -Identity "OnlineStudents" -Members $SamAccountName
            }
            elseif ($User.Department -eq "Cross Student") {
                Add-ADGroupMember -Identity "CROSSTraining" -Members $SamAccountName
            }

            #Adds all students created to the StudentGroup security group
            Add-ADGroupMember -Identity "StudentGroup" -Members $SamAccountName 

            write-host "EmployeeNumber: $employeeNumber"
           
# If user is created, show message.
            Write-Host "The user $SamAccountName was created and attributes are set.
            -
            -
            -
            -" -ForegroundColor Green

            # Remember the UPN so we can wait for it to sync to Entra below.
            $newUpnList += "$SamAccountName@$UPN"


        }
    }
    catch {
        # Handle any errors that occur during account creation
        Write-Host "Failed to create user $SamAccountName - $_" -ForegroundColor Red
    }
}



# Report username collisions so none of them are silently dropped.
if ($collisions.Count -gt 0 -or $readmitCandidates.Count -gt 0 -or $reviewCandidates.Count -gt 0) {
    $bodyLines = @()
    $bodyLines += "Username collision report from 01-CreateStudentAccounts.ps1"
    $bodyLines += "Run: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    $bodyLines += ""

    if ($collisions.Count -gt 0) {
        $bodyLines += "SUFFIXED - a different student already owns the base username. These accounts were created with a numeric suffix:"
        foreach ($c in $collisions) {
            $bodyLines += " - $($c.Name) (employeeID $($c.CSVEmployeeID)): $($c.BaseUsername) belongs to employeeID $($c.ExistingID) at $($c.ExistingDN). Created as $($c.AssignedUsername)."
        }
        $bodyLines += ""
    }

    if ($readmitCandidates.Count -gt 0) {
        $bodyLines += "NOT CREATED - the same student (matching employeeID) already has an account. If this student is returning, the row belongs in NewStudentsAccounts-ReAdmit.csv so 05-ReAdmitAccounts.ps1 can reactivate it:"
        foreach ($r in $readmitCandidates) {
            $bodyLines += " - $($r.Name) (employeeID $($r.EmployeeID)): existing account $($r.SamAccountName) at $($r.ExistingDN), Enabled=$($r.Enabled)."
        }
        $bodyLines += ""
    }

    if ($reviewCandidates.Count -gt 0) {
        $bodyLines += "NOT CREATED - needs manual review:"
        foreach ($v in $reviewCandidates) {
            $bodyLines += " - $($v.Name) (CSV employeeID '$($v.CSVEmployeeID)'): $($v.Reason). Existing: $($v.SamAccountName) at $($v.ExistingDN)."
        }
        $bodyLines += ""
    }

    Write-Host ($bodyLines -join "`n") -ForegroundColor Cyan

    try {
        Send-MailMessage `
            -To "support@okwu.edu" `
            -From "accounts@okwu.edu" `
            -Subject "Username collisions in student account creation ($($collisions.Count) suffixed, $($readmitCandidates.Count) readmit, $($reviewCandidates.Count) review)" `
            -Body ($bodyLines -join "`n") `
            -SmtpServer "10.0.5.16"
        Write-Host "Collision report emailed to support@okwu.edu" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Failed to send collision report email: $($_.Exception.Message)"
    }
}

#Take a quick nap
sleep 5


#Run delta sync
Start-ADSyncSyncCycle -PolicyType delta

# If this run created accounts, wait for AAD Connect to provision them in Entra
# before importing authentication methods. Otherwise keep a short settle delay.
if ($newUpnList.Count -gt 0) {
    Write-Host "Waiting for AAD Connect to provision $($newUpnList.Count) new account(s) in Entra..." -ForegroundColor Cyan
    try {
        Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.33.0 -Force -ErrorAction Stop

        $pollClientId   = 'e5ec9711-41fb-494a-86d8-2bdf70859e8c'
        $pollTenantId   = '3ec35b81-463a-4801-864c-fedf6e44de2e'
        $pollThumbprint = 'a668ab98819e1be444d7d486029e4d4ad6cdb20c'

        $cert = Get-ChildItem -Path Cert:\CurrentUser\My -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $pollThumbprint }
        if (-not $cert) {
            $cert = Get-ChildItem -Path Cert:\LocalMachine\My |
                Where-Object { $_.Thumbprint -eq $pollThumbprint }
        }

        if (-not $cert) {
            Write-Warning "Entra polling certificate not found; skipping the sync wait."
        }
        else {
            Connect-MgGraph -ClientId $pollClientId -TenantId $pollTenantId -CertificateThumbprint $pollThumbprint -NoWelcome -ContextScope Process -ErrorAction Stop | Out-Null

            $notSeen  = @($newUpnList)
            $deadline = (Get-Date).AddMinutes(15)
            while ($notSeen.Count -gt 0 -and (Get-Date) -lt $deadline) {
                $stillMissing = @()
                foreach ($upn in $notSeen) {
                    $encoded = [uri]::EscapeDataString($upn)
                    try {
                        Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/${encoded}?`$select=id" -ErrorAction Stop | Out-Null
                        Write-Host "Provisioned: $upn" -ForegroundColor Green
                    }
                    catch {
                        $statusCode = $null
                        if ($_.Exception.Response) {
                            $statusCode = [int]$_.Exception.Response.StatusCode
                        }
                        Write-Warning "Graph API check for $upn failed (HTTP $statusCode): $($_.Exception.Message)"
                        $stillMissing += $upn
                    }
                }
                $notSeen = $stillMissing
                if ($notSeen.Count -gt 0) {
                    Write-Host "AAD Connect still provisioning $($notSeen.Count) of $($newUpnList.Count) account(s)..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 10
                }
            }

            if ($notSeen.Count -gt 0) {
                Write-Warning "Timed out waiting for these account(s) to sync to Entra (check AAD Connect sync errors): $($notSeen -join ', ')"

                $timeoutBody = @"
Entra sync timed out after 15 minutes.

The following account(s) were not visible in Entra ID:
$($notSeen -join "`n")

Check AAD Connect sync errors for these UPNs.
"@

                try {
                    Send-MailMessage `
                        -To "kringnalda@okwu.edu" `
                        -From "entrasync@okwu.edu" `
                        -Subject "WARNING: Entra sync timeout - $($notSeen.Count) account(s) not provisioned" `
                        -Body $timeoutBody `
                        -SmtpServer "10.0.5.16"
                    Write-Host "Timeout notification email sent to kringnalda@okwu.edu" -ForegroundColor Cyan
                }
                catch {
                    Write-Warning "Failed to send timeout notification email: $($_.Exception.Message)"
                }
            }
            else {
                Write-Host "All $($newUpnList.Count) new account(s) are visible in Entra." -ForegroundColor Green
            }
        }
    }
    catch {
        Write-Warning "Entra sync-wait failed ($($_.Exception.Message)); continuing anyway."
    }
    finally {
        if ((Get-Module Microsoft.Graph.Authentication) -and (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue)) {
            Disconnect-MgGraph | Out-Null
        }
    }
}
else {
    Start-Sleep -Seconds 30
}



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
.\04-ProvisionJenzabar.ps1


# Delete old log files


# Set the directory path where you want to delete old files
$target_folder = Join-Path -Path $PSScriptRoot -ChildPath "Logs"


# Get the current date and subtract x many days to get the threshold date
$days_old = 60
$threshold_date = (Get-Date).AddDays(-$days_old)

# Get all the files in the target folder and its subfolders that are older than the threshold date
Get-ChildItem -Path $target_folder -Recurse | Where-Object {
    $_.LastWriteTime -lt $threshold_date
} | ForEach-Object {
    # Remove the file; a locked file (e.g. an open transcript) is skipped and retried next run
    Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue -ErrorVariable deleteErrors
    if ($deleteErrors) {
        Write-Host "Skipped (in use): $($_.FullName)" -ForegroundColor Yellow
    }
    else {
        Write-Host "Deleted: $($_.FullName)"
    }
}


Write-Host "Files older than $days_old days have been deleted from $target_folder"

write-host "
-
-
-
- 
-
('-')<This concludes the deletion log.
- 
-
-
-
-"


Start-Sleep -Seconds 5


# End logging
Write-PSFMessage -Level Output -Message "Script ended."

cd C:\Account_Creation
Stop-Transcript