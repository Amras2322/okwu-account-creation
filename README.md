# OKWU Student Account Creation

PowerShell pipeline that turns the student feed into working OKWU accounts: Active Directory, Entra ID authentication methods, and Jenzabar One access.

## Where it lives

- **Runs on:** the `adsync` server at `C:\Account_Creation`
- **Source:** `https://github.com/Amras2322/okwu-account-creation`
- **Logs:** `C:\Account_Creation\Logs` (transcripts older than 60 days are deleted by the pipeline)
- **Inputs/outputs:** CSV files in `C:\Account_Creation` (see below)

Deploy by copying the scripts to `C:\Account_Creation` on `adsync`. `01-CreateStudentAccounts.ps1` is run by the hourly scheduled task.

## What it does

```
NewStudentsAccounts-auto.csv
        |
        v
01-CreateStudentAccounts.ps1   Create AD accounts, OUs, groups, attributes
        |                      Run AAD Connect delta sync, wait for Entra
        v
02-SyncPhoneAttribute.ps1      Normalize phones -> EntraMethods.csv
        |
        v
03-ImportEntraMethods.ps1      Add/update/remove Entra phone & email methods
        |
        v
04-ProvisionJenzabar.ps1       Provision Jenzabar SQL rows, email the student
```

### 01 - Create student accounts

Reads `NewStudentsAccounts-auto.csv`, creates enabled AD accounts and places them in the OU for their department (Traditional, Concurrent, AGS/GPS, NonDegree, Online, Cross Training, Alumni). Sets `extensionAttribute1` (personal email), `extensionAttribute2` (raw mobile), and `employeeNumber` (for PaperCut), and adds the department security group plus `StudentGroup`.

Identity is `employeeID`, not name. If the ID already has an account, no duplicate is created. If the base `First.Last` username belongs to a different student, a number is appended (`First.Last1`, `First.Last2`, ...). Ambiguous rows are skipped and emailed to support in a collision report. After creating accounts, it runs an AAD Connect delta sync and waits up to 15 minutes for the new users to appear in Entra.

### 02 - Sync phone and email attributes

Delta job (default: users changed in the last hour; use `-FullSync` for everything). Normalizes `extensionAttribute2` to `+1 XXXXXXXXXX` in `extensionAttribute10`, exports phone numbers and personal emails to `EntraMethods.csv`, and emails support about malformed numbers. A blank `extensionAttribute2` intentionally clears the exported phone.

### 03 - Import Entra authentication methods

Reads `EntraMethods.csv` and adds, updates, or removes mobile-phone and authentication-email methods in Entra ID. Runs in preview mode unless `-Apply` is passed; the pipeline always passes `-Apply`. Uses certificate app-only auth when available, otherwise delegated sign-in.

### 04 - Provision Jenzabar

Matches each student to their AD account by `employeeID`, then inserts the OKWU email into Jenzabar's `AlternateContactMethod`, `TW_WEB_SECURITY`, and `TW_GRP_MEMBERSHIP` tables (with retries). Sends the "MyOKWU Login created" welcome email to the student's personal address. Missing, duplicate, or alumni-only accounts are reported to support.

### 05 - Re-admit returning students

Separate script for `NewStudentsAccounts-ReAdmit.csv`, not part of the hourly run. It re-enables the existing account, moves it to the correct department OU, refreshes attributes, reconciles the managed student groups, then runs `02` -> `03` -> `04` against the readmit CSV. Passwords are never changed; students use the MyOKWU forgot-password flow. Use `-WhatIf` for a dry run.

## Files

| File | Purpose |
| --- | --- |
| `01-CreateStudentAccounts.ps1` | Create AD accounts and orchestrate the pipeline |
| `02-SyncPhoneAttribute.ps1` | Normalize phone attributes and export Entra methods CSV |
| `03-ImportEntraMethods.ps1` | Write phone/email authentication methods to Entra ID |
| `04-ProvisionJenzabar.ps1` | Provision Jenzabar SQL and send the welcome email |
| `05-ReAdmitAccounts.ps1` | Reactivate returning students and run the pipeline |
| `NewStudentsAccounts-auto.csv` | New student feed (not in git) |
| `NewStudentsAccounts-ReAdmit.csv` | Readmit feed (not in git) |
| `EntraMethods.csv` | Generated handoff between 02 and 03 |
| `Logs/` | Transcripts (not in git) |

## Running manually

Run from `C:\Account_Creation` in an elevated PowerShell session on `adsync`:

```powershell
# Normal pipeline (creates accounts and runs 02-04)
.\01-CreateStudentAccounts.ps1

# Readmit dry run, then a real run
.\05-ReAdmitAccounts.ps1 -WhatIf
.\05-ReAdmitAccounts.ps1

# Phone/email only
.\02-SyncPhoneAttribute.ps1 -HoursBack 1 -WhatIf
.\03-ImportEntraMethods.ps1 -CsvPath .\EntraMethods.csv          # preview
.\03-ImportEntraMethods.ps1 -CsvPath .\EntraMethods.csv -Apply   # commit
```

## Requirements

- Windows PowerShell 5.1, `ActiveDirectory` and `SqlServer` modules, `PSFramework`, and `Microsoft.Graph.Authentication` 2.33.0
- Access to the `okwu.edu` domain and the `jzowo-sql1.owo.jaas` SQL server
- Entra app registration with `UserAuthenticationMethod.ReadWrite.All` and its certificate installed on `adsync`
- SMTP relay at `10.0.5.16` for support notifications and student welcome emails

## Notes

- Scripts contain environment-specific values (tenant/client IDs, certificate thumbprint, SMTP host, initial password). Treat the repo as internal and rotate these values if they change.
- All stages log to `Logs\`. Stage failures are caught so later stages still run; check the transcript and support emails after a failed run.
