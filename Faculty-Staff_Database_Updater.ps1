#################################################################
#Active Directory Section
#################################################################

#Import needed modules
Import-Module ActiveDirectory


# Get the employee ID of the user that needs updating
$employeeID = Read-Host "Enter Employee ID"


#Get the user based on employeeID
$user = Get-ADUser -Filter "employeeID -eq '$employeeID'" -Properties UserPrincipalName 

# Store the UserPrincipalName in a variable
$UPN = $user.UserPrincipalName






###################################################################
#SQL Section
###################################################################

#Import needed modules
Import-Module SqlServer

#Define database variables
$serverName = "jzowo-sql1.owo.jaas"
$databaseName = "J1PRD"
$time = Get-Date -Format MMddyyyyHHmm

#Define AppIDQuery query statement
$AppIDQuery = "SELECT appid FROM namemaster WHERE id_num = '$employeeID'"

#Set namemasterappid variable using SQL query
$namemasterappid = invoke-sqlcmd -Query $AppIDQuery -ServerInstance $serverName -Database $databaseName -QueryTimeout 0 -SuppressProviderContextWarning -OutputSqlErrors $true  -EncryptConnection -TrustServerCertificate



# Creating the INSERT query using the variables defined
## Adding OKWU Email to alternate contact
    $SQLQuery3 = "
        INSERT INTO AlternateContactMethod (namemasterappid, AlternateContact, AlternateContactTypeDefAppID, ID_NUM, ADDR_CDE, INSTITUTION_EMAIL, NOTIFICATION_ENABLED)
        VALUES ('$namemasterappid', '$UPN', '-1', '$employeeID', '*EML', 'Y', 'Y')
        "


# Use query and insert to SQL table campus6.dbo.EmailAddress
    $SQLQuery3out = Invoke-Sqlcmd -Query $SQLQuery3 -ServerInstance $serverName -Database $databaseName -QueryTimeout 0 -SuppressProviderContextWarning -OutputSqlErrors $true  -EncryptConnection -TrustServerCertificate
    $SQLQuery3out

Start-Sleep -Seconds 1

Write-Host "-"
Write-host "User exists $adUsername. Creating in Jenzabar One TW Tables (TW_WEB_SECURITY) (TW_GRP_MEMBERSHIP)"




# Updating TW Web Tables

    $SQLQuery1 = "
        INSERT INTO TW_WEB_SECURITY (ID_NUM, ACCESS_CDE, WEB_LOGIN)
        VALUES ('$emplyeeID', '1234', '$UPN')
        "

# SQLQuery2 Created to Web Security
    $SQLQuery2 = "
        INSERT INTO TW_GRP_MEMBERSHIP (ID_NUM, Group_id)
        VALUES ('$employeeID', 'faculty')
        "



# PowerShell Runs SQL Commands to run the SQLQueries from Above, Shows Results

# PowerShell Runs the SQL Commands for $SQLQuery1
    $SQLQuery1out = Invoke-Sqlcmd -Query $SQLQuery1 -ServerInstance $serverName -Database $databaseName -QueryTimeout 0 -SuppressProviderContextWarning -OutputSqlErrors $true  -EncryptConnection -TrustServerCertificate
    $SQLQuery1out

# PowerShell Runs the SQL Commands for $SQLQuery2
    $SQLQuery2out = Invoke-Sqlcmd -Query $SQLQuery2 -ServerInstance $serverName -Database $databaseName -QueryTimeout 0 -SuppressProviderContextWarning -OutputSqlErrors $true  -EncryptConnection -TrustServerCertificate
    $SQLQuery2out