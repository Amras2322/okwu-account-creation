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
        

        # Generate SamAccountName based on givenName and SurName columns
        $SamAccountName = "$($User.givenName).$($User.SurName)"
        
        # Creates Name of user
        $Name = "$($User.GivenName) $($User.Surname)"      

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
            DisplayName           = $Name
            Country               = "US"
            Path                  = $OU                       # Determines which OU user is placed in
            AccountPassword       = (ConvertTo-SecureString $Pass -AsPlainText -Force)
            Company               = "Oklahoma Wesleyan University"
            ChangePasswordAtLogon = $true                     # Account must be enabled for this to take effect
        }

        # Debugging: Output all parameters to verify
        Write-Host "Creating user with the following parameters:"
        $UserParams.GetEnumerator() | ForEach-Object { Write-Host "$($_.Key): $($_.Value)" }

        # Check if the user already exists in AD
        if (Get-ADUser -Filter { SamAccountName -eq $SamAccountName }) {
            # Give a warning if user exists
            Write-Host "A user with username $SamAccountName already exists in Active Directory." -ForegroundColor Yellow
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