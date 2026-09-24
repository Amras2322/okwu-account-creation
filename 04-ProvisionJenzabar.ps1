param(
    [Parameter(Mandatory = $false)]
    [string]$CsvPath = "C:\Account_Creation\NewStudentsAccounts-auto.csv"
)

Import-Module PSFramework

Import-Module SqlServer

$users = Import-Csv -Path $CsvPath

$serverName = "jzowo-sql1.owo.jaas"
$databaseName = "J1PRD"
$time = Get-Date -Format MMddyyyyHHmm

$MaxRetries = 3
$RetryDelaySeconds = 10

# Install-Module -Name SqlServer
Get-Date

# Start logging
Write-PSFMessage -Level Output -Message "Script started."


# Start the Array for Each user from the CSV
ForEach ($user in $users) {
    Write-host "[******************************************************************************]"
    Write-host "[******************************************************************************]"

    # Set Variables from the CSV
    $GivenName = $user.GivenName
    $Surname = $user.Surname
    $emailaddress = $user.PersonalEmail
    $Id_num = $user.employeeID
    $CSVUsername = "$GivenName.$Surname"
    $NameMasterAppID = $user.NameMasterAPPID

    # Get-date for creation timestamp
    $CreateDate = Get-Date -Format "MM/dd/yyyy"
    $Createtime = Get-Date -Format "hh:mm tt"
    $timeMOD = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Cross reference Active Directory to make sure the account is created.
    $adlook = Get-ADUser -Filter "GivenName -eq '$GivenName' -and Surname -eq '$Surname' -and EmployeeID -eq '$Id_Num'" -Properties *
    $adlookup = Get-ADUser -Filter "GivenName -eq '$GivenName' -and Surname -eq '$Surname' -and EmployeeID -eq '$Id_num'" -Properties SamAccountName, UserPrincipalName
    $adUsername = $adlook.SamAccountName
    $okwuEmail = $adlook.UserPrincipalName
    $ADou = $adlook.DistinguishedName
    $cn = $adlook.cn

    $AlumniCheck = "CN=$cn,OU=Alumni-WebAccessOnly,OU=Alumni,DC=okwu,DC=edu"
    # Convert Username to a format to match SQL
    # Enclose Variables with ' ' (single quotes)
    $UserName0 = $adUsername
    $UserName = "'$UserName0'"

    # If Commands to check if AD value is empty, duplicate, or single
    if ($adlookup -eq $null) {
        Write-host "User does not exist in AD $CSVUsername. Sending email to Support."
        Write-host "-"
        Write-host "-------------------------------------------------------------------------------------------------"
        Send-MailMessage -To "support@okwu.edu" -From "accounts@okwu.edu" -Subject "User does not exist $CSVUsername" -Body "User does not exist in AD $CSVUsername. Please send an email to Support." -SmtpServer 10.0.5.16
    }

    if ($ADou -eq $AlumniCheck) {
        Write-host "User in AlumniOU $ADou. Sending email to Support."
        Write-host "-"
        Write-host "-------------------------------------------------------------------------------------------------"
        Send-MailMessage -To "support@okwu.edu" -From "accounts@okwu.edu" -Subject "User in AlumniOU " -Body "User in AlumniOU $ADou" -SmtpServer 10.0.5.16
    }

    if (($adlook | Measure).Count -gt 1) {
        Write-host "User exists with 2 accounts or more in AD $adlookup. Sending email to Support."
        Write-host "-"
        Write-host "*************************************************************************************************"
        Send-MailMessage -To "support@okwu.edu" -From "accounts@okwu.edu" -Subject "AD user has 2 accounts - $adUsername" -Body "AD user has 2 accounts $adUsername" -SmtpServer 10.0.5.16
    }

    if (($adlook | Measure).Count -eq 1) {
        Write-host "+"
        Write-host "User exists in AD. Finishing account creation for $adlookup."
        Write-host "-"
        Write-host "User exists in AD. Copying $okwuEmail to Jenzabar One"

        # Creating the INSERT query using the variables defined
        ## Adding OKWU Email to alternate contact
        $SQLQuery3 = "
        IF NOT EXISTS (SELECT 1 FROM AlternateContactMethod WHERE namemasterappid = '$NameMasterAppID' AND AlternateContact = '$okwuEmail')
        INSERT INTO AlternateContactMethod (namemasterappid, AlternateContact, AlternateContactTypeDefAppID, ID_NUM, ADDR_CDE, INSTITUTION_EMAIL, NOTIFICATION_ENABLED)
        VALUES ('$NameMasterAppID', '$okwuEmail', '-1', '$Id_num', '*EML', 'Y', 'Y')
        "
       
        # Use query and insert to SQL table campus6.dbo.EmailAddress
        $sqlSuccess = $false
        $lastSqlError = $null
        for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
            try {
                $SQLQuery1Output = Invoke-Sqlcmd -Query $SQLQuery3 -ServerInstance $serverName -Database $databaseName -QueryTimeout 0 -SuppressProviderContextWarning -OutputSqlErrors $true -EncryptConnection -TrustServerCertificate -ErrorAction Stop
                $SQLQuery1Output
                $sqlSuccess = $true
                break
            }
            catch {
                $lastSqlError = $_.Exception.Message
                Write-Warning "AlternateContactMethod query failed (attempt $attempt of $MaxRetries): $lastSqlError"
                if ($attempt -lt $MaxRetries) {
                    Write-Host "Retrying in $RetryDelaySeconds seconds..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $RetryDelaySeconds
                }
            }
        }
        if (-not $sqlSuccess) {
            Write-Warning "All $MaxRetries attempts failed for AlternateContactMethod query for $adUsername. Sending failure notification."
            Send-MailMessage -To "support@okwu.edu" -From "accounts@okwu.edu" -Subject "Jenzabar provisioning failed for $adUsername" -Body "All $MaxRetries attempts failed to insert into AlternateContactMethod for user $adUsername (Employee ID: $Id_num).`n`nError: $lastSqlError`n`nThis user's student email was not provisioned and must be added manually." -SmtpServer 10.0.5.16
        }

        Start-Sleep -Seconds 1

        Write-Host "-"
        Write-host "User exists $adUsername. Creating in Jenzabar One TW Tables (TW_WEB_SECURITY) (TW_GRP_MEMBERSHIP)"

        ######################################################################
        # Updating TW Web Tables                                             #

        $SQLQuery1 = "
        IF NOT EXISTS (SELECT 1 FROM TW_WEB_SECURITY WHERE ID_NUM = '$Id_num')
        INSERT INTO TW_WEB_SECURITY (ID_NUM, ACCESS_CDE, WEB_LOGIN)
        VALUES ('$Id_num', '1234', '$okwuEmail')
        "

        # SQLQuery2 Created to Web Security
        $SQLQuery2 = "
        IF NOT EXISTS (SELECT 1 FROM TW_GRP_MEMBERSHIP WHERE ID_NUM = '$Id_num' AND Group_id = 'student')
        INSERT INTO TW_GRP_MEMBERSHIP (ID_NUM, Group_id)
        VALUES ('$Id_num', 'student')
        "

        ################################################################################################
        # PowerShell Runs SQL Commands to run the SQLQueries from Above, Shows Results

        # PowerShell Runs the SQL Commands for $SQLQuery1
        $sqlSuccess = $false
        $lastSqlError = $null
        for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
            try {
                $SQLQuery1out = Invoke-Sqlcmd -Query $SQLQuery1 -ServerInstance $serverName -Database $databaseName -QueryTimeout 0 -SuppressProviderContextWarning -OutputSqlErrors $true -EncryptConnection -TrustServerCertificate -ErrorAction Stop
                $SQLQuery1out
                $sqlSuccess = $true
                break
            }
            catch {
                $lastSqlError = $_.Exception.Message
                Write-Warning "TW_WEB_SECURITY query failed (attempt $attempt of $MaxRetries): $lastSqlError"
                if ($attempt -lt $MaxRetries) {
                    Write-Host "Retrying in $RetryDelaySeconds seconds..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $RetryDelaySeconds
                }
            }
        }
        if (-not $sqlSuccess) {
            Write-Warning "All $MaxRetries attempts failed for TW_WEB_SECURITY query for $adUsername. Sending failure notification."
            Send-MailMessage -To "support@okwu.edu" -From "accounts@okwu.edu" -Subject "Jenzabar provisioning failed for $adUsername" -Body "All $MaxRetries attempts failed to insert into TW_WEB_SECURITY for user $adUsername (Employee ID: $Id_num).`n`nError: $lastSqlError`n`nThis user's student email was not provisioned and must be added manually." -SmtpServer 10.0.5.16
        }

        # PowerShell Runs the SQL Commands for $SQLQuery2
        $sqlSuccess = $false
        $lastSqlError = $null
        for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
            try {
                $SQLQuery2out = Invoke-Sqlcmd -Query $SQLQuery2 -ServerInstance $serverName -Database $databaseName -QueryTimeout 0 -SuppressProviderContextWarning -OutputSqlErrors $true -EncryptConnection -TrustServerCertificate -ErrorAction Stop
                $SQLQuery2out
                $sqlSuccess = $true
                break
            }
            catch {
                $lastSqlError = $_.Exception.Message
                Write-Warning "TW_GRP_MEMBERSHIP query failed (attempt $attempt of $MaxRetries): $lastSqlError"
                if ($attempt -lt $MaxRetries) {
                    Write-Host "Retrying in $RetryDelaySeconds seconds..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $RetryDelaySeconds
                }
            }
        }
        if (-not $sqlSuccess) {
            Write-Warning "All $MaxRetries attempts failed for TW_GRP_MEMBERSHIP query for $adUsername. Sending failure notification."
            Send-MailMessage -To "support@okwu.edu" -From "accounts@okwu.edu" -Subject "Jenzabar provisioning failed for $adUsername" -Body "All $MaxRetries attempts failed to insert into TW_GRP_MEMBERSHIP for user $adUsername (Employee ID: $Id_num).`n`nError: $lastSqlError`n`nThis user's student email was not provisioned and must be added manually." -SmtpServer 10.0.5.16
        }

        Write-Host "-"
        Write-host "User exists $adUsername. Sending email to $emailaddress."

$body=@" 
<html>
<head>

</head>

<body lang=EN-US link="#0563C1" vlink="#954F72" style='tab-interval:.5in; overflow-wrap: "break-word"; word-break: normal; white-space: normal; font-size:12.0pt;font-family:"Aptos",sans-serif' >

<br>
 
<!-- <p><img width="363" height="78" src="https://www.okwu.edu/wp-content/uploads/2021/03/IT-Logo.png" style="height: .808in; width: 3.783in;" v:shapes="Picture_x0020_1"></span></p> -->

<p>Congratulations, you are officially a member of the Eagle family here at Oklahoma Wesleyan
University!! Your <b><i>MyOKWU</i></b> Login has been created!</p>

<p style='font-size:12.0pt;font-family:"Aptos",sans-serif'>Everything starts at
MyOKWU: <a href="https://my.okwu.edu">https://my.okwu.edu</a>. The Apps page in MyOKWU contains several links that you will use during your time at Oklahoma Wesleyan University.</p>

<p style='font-size:12.0pt;font-family:"Aptos",sans-serif'>
<b>Your username is (not case sensitive): <br>
$adUsername</b></p>

<p><b>Your OKWU email is:<br>
$okwuEmail</b></p>

<p style='font-size:16.0pt; text-decoration: underline;'><b>To access the OKWU Portal:</b></p>

<ol style='margin-top:0in' start=1 type=1>
 <li style='margin-left:0in;mso-list:l2 level1 lfo3'><span
     ><span style='font-size:12.0pt;
     font-family:"Aptos",sans-serif;mso-fareast-font-family:"Times New Roman"'>Go
     to</span></span><span
     style='font-family:"Aptos",sans-serif'> </span></span><a href="https://my.okwu.edu"><span
     style='font-family:"Aptos",sans-serif'>https://my.okwu.edu</span></span><span
     ></span></a><span style='mso-bookmark:
     _MailOriginal'><span style='font-family:"Aptos",sans-serif'> and click the
     blue Login button in the top right corner.</span></span><span style='mso-bookmark:
     _MailOriginal'><span style='font-size:12.0pt;font-family:"Aptos",sans-serif;
     mso-fareast-font-family:"Times New Roman"'></span></span></li>
 <li style='margin-left:0in;mso-list:l2 level1 lfo3'><span
     ><span style='font-size:12.0pt;
     font-family:"Aptos",sans-serif;mso-fareast-font-family:"Times New Roman"'>Click
     &quot;Forgot Password&quot; below the Sign In button on the login page.</span></span></li>
 <li style='margin-left:0in;mso-list:l2 level1 lfo3'><span
     ><span style='font-size:12.0pt;
     font-family:"Aptos",sans-serif;mso-fareast-font-family:"Times New Roman"'>Enter
     the Username listed above, then click <i>Submit. </i></span></span></li>
 <li style='margin-left:0in;mso-list:l2 level1 lfo3'><span
     ><span style='font-size:12.0pt;
     font-family:"Aptos",sans-serif;mso-fareast-font-family:"Times New Roman"'>Locate
     the one-time passcode sent to you (options are SMS or Email)</span></span></li>
 <li style='margin-left:0in;mso-list:l2 level1 lfo3'><span
     ><span style='font-size:12.0pt;
     font-family:"Aptos",sans-serif;mso-fareast-font-family:"Times New Roman"'>Enter
     the security code sent via SMS or Email (Click \“Problems with this
     authentication options?\” for alternative authentication options)</span></span></li>
 <li style='margin-left:0in;mso-list:l2 level1 lfo3'><span
     ><span style='font-size:12.0pt;
     font-family:"Aptos",sans-serif;mso-fareast-font-family:"Times New Roman"'>In
     the space provided, create your new password*, then click <i>Continue</i></span></span></li>
</ol>

<p style='text-indent:.5in'><span
style='font-size:12.0pt;font-family:"Aptos",sans-serif'>*Your password must
meet the requirements listed below:</span></span></p>

<ul style='margin-top:0in' type=disc>
 <li style='margin-left:.25in;mso-list:l3 level1 lfo6'><span ><span style='font-size:12.0pt;
     font-family:"Aptos",sans-serif;mso-fareast-font-family:"Times New Roman"'>Must
     be AT LEAST twelve (12) characters </span></span></li>
 <li style='margin-left:.25in;mso-list:l6 level1 lfo9'><span ><span style='font-size:12.0pt;
     font-family:"Aptos",sans-serif;mso-fareast-font-family:"Times New Roman"'>Contain
     at least one uppercase letter (A through Z)</span></span></li>
 <li style='margin-left:.25in;mso-list:l6 level1 lfo9'><span
     ><span style='font-size:12.0pt;
     font-family:"Aptos",sans-serif;mso-fareast-font-family:"Times New Roman"'>Contain
     at least one lowercase letter (a through z)</span></span></li>
 <li style='margin-left:.25in;mso-list:l6 level1 lfo9'><span
     ><span style='font-size:12.0pt;
     font-family:"Aptos",sans-serif;mso-fareast-font-family:"Times New Roman"'>Contain
     at least one number (0-9)</span></span></li>
 <li style='margin-left:.25in;mso-list:l6 level1 lfo9'><span
     ><span style='font-size:12.0pt;
     font-family:"Aptos",sans-serif;mso-fareast-font-family:"Times New Roman"'>Contain
     at least one special character (ex. !, @, $)</span></span></li>
</ul>

<p style='text-indent:.5in'>#Your password may not
contain any parts of your name<br></p>

<p style='text-indent:.5in'>#Your password cannot contain the following &lt; &amp; #</p>

<p style='text-indent:.5in'>#Users must reset their password every 180 days</p>

<ol style='margin-top:0in' start=7 type=1>
 <li style='margin-left:0in;mso-list:l2 level1 lfo3'>Once your password has been successfully updated, log in using your new password</span></span>
 </li>
</ol>
<br>

<p>If you have questions or run into issues, please email <a href="mailto:support@okwu.edu">support@okwu.edu</a> or call 918.335.6233. Welcome to the OKWU Family!</p>

<div>

<div>

<div>

<br>
<br>
                  OKWU IT Support<br>
                  Information Technology<br>
            </span></span>
            <a href="mailto:support@okwu.edu">support@okwu.edu</span></a><br>
                  918.335.6233<br>
                  OKLAHOMA WESLEYAN UNIVERSITY<br>
</p>

</div>

</div>

</div>


</div>

</body>

</html>



"@
 send-mailmessage -to "$emailaddress" -from "support@okwu.edu" -Bcc "accounts@okwu.edu"-subject "MyOKWU Login created" -bodyashtml $body -SmtpServer 10.0.5.16 
 Start-Sleep -Seconds 2.5
}
}



