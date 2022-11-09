# Declare the reports file path,
$logPath = "[NASLOCATION]\CloudScripts\O365License\Reports\"
$sday = Get-Date -UFormat "%Y%m%d-%H%M"
$scriptlog = New-Item -ItemType file -Path ($logpath + $sday + "-ScriptLog.txt") -Force
$dataout = New-Object System.Collections.ArrayList
<# 
There are a few ways to create secure files to be used as credentials. This method has some downsides, 
for instance the script that you are running could be changed to do something else. To mitigate this make sure
your service account has the least amount of rights as needed. I am moving this code to an Azure automation where
protecting the script will be easier.
This is how i created the credential file for the script:
$credpath = C:\scripts\MyCredential.xml
New-Object System.Management.Automation.PSCredential("[YOURSERVICEACCOUNTUPN]", (ConvertTo-SecureString -AsPlainText -Force "Password123")) | Export-CliXml $credpath
#>
# Get credentials for MSOnLine
$credpath = "C:\Users\[SERVICEACCOUNT]\secxml.xml"
$cred = import-clixml -path $credpath
Connect-MsolService -Credential $cred
# Get all MSOL users
$msolusers = Get-MsolUser -All
# Build an array to hold SKUs
$skuArray = New-Object System.Collections.ArrayList
# Build an array to hold our top level license assignments
$userLicenseArray = New-Object System.Collections.ArrayList
# Create an Array and Hash to capture AD and logon data.
$adlogon = New-Object System.Collections.ArrayList
$ADUserHash = @{ }
# Create Array for all members
$AllMembers = New-Object System.Collections.ArrayList
# Functions ########################################################################
function Write-Logs ($message, $type, $log)
{
    $date = Get-date -UFormat "%Y/%m/%d %H:%M:%S"
    [string]$entry = "$($date): $($type) - $($message)"
    Add-Content -Path $log -Value $entry -Force
} # End Write-Logs
function Get-Bailout ($message)
{
    Write-Logs `
               -message $message `
               -type "ERROR" `
               -log $scriptlog.FullName
    $Program = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.txt\OpenWithList' -Name a).a
    & $Program $scriptlog.FullName
} # End Get-Bail
function Test-Logpath ($scriptlog)
{
    if (!(Test-Path -Path $scriptlog))
    {
        try
        {
            $log = New-Item -ItemType file -Path $scriptlog -Force -ErrorAction Stop
            $message = "Log file has been created!"
            Write-Verbose $message
            Write-Logs `
                       -message $message `
                       -type "INFO" `
                       -log $log
        }
        catch
        {
            Write-Verbose -Message ("Unable to create log file " + $scriptlog + " exiting.")
            Get-Bailout -message $null
        }
    }
    else
    {
        $log = Get-childItem $scriptlog
    }
    $log
} # End Test-Logpath
# End Functions ####################################################################
Write-Verbose -Message ("Creating log file " + $scriptlog + ".")
$log = Test-Logpath -scriptlog $scriptlog
$message = ("Log file exists, starting script.")
Write-Verbose $message
Write-Logs `
           -message $message `
           -type "INFO" `
           -log $log
# This section creates an array of sublicenses for each SKU
$allsku = Get-MsolAccountSku
foreach ($sku in $allsku)
{
    $name = $sku.AccountSkuId
    ($name -split (":"))[1]
    $temp = "" | Select-Object UPN
    $temp.UPN = "Seed.User@domain.com"
    New-Variable -name ($name -split (":"))[1] -value (New-Object System.Collections.ArrayList)
    foreach ($subsku in $sku.ServiceStatus.serviceplan)
    {
        $subsku.servicename
        # (Get-Variable ($name -split (":"))[1])| Add-Member -NotePropertyName $subsku.servicename -NotePropertyValue ""
        $temp | Add-Member -NotePropertyName $subsku.servicename -NotePropertyValue "Success"
    }
    (Get-Variable ($name -split (":"))[1]).Value.Add($temp)
    $arraylist.Add(($name -split (":"))[1])
}
$LicenseArray = "" | Select-Object UPN,ImmutableId
$LicenseArray.UPN = "Seed.User@domain.com"
$LicenseArray.ImmutableId = $null
foreach ($array in $arraylist)
{
    $array
    $LicenseArray |Add-Member -NotePropertyName $array -NotePropertyValue $false
}
$userLicenseArray.add($LicenseArray)
foreach ($user in $msolusers)
{
    if ($user.isLicensed)
    {
        $myLicObj = "" | Select-Object UPN,ImmutableId
        $myLicObj.upn = $user.userprincipalname
        $myLicObj.ImmutableId = $user.ImmutableId
        foreach ($lic in $user.Licenses)
        {
            $myObj = "" | Select-Object UPN
            $myObj.upn = $user.userprincipalname
            $skuID = ($lic.AccountSkuId -split ":")[1]
            $myLicObj | Add-Member -NotePropertyName $skuID -NotePropertyValue $true
            foreach ($plan in $lic.servicestatus)
            {
                $myobj | Add-Member -NotePropertyName $plan.serviceplan.ServiceName -NotePropertyValue $plan.ProvisioningStatus
            }
            "Adding to granular array"
            $skuID
            (Get-Variable ($skuID)).Value.add($myobj)
        }
        "Adding to top array"
        $skuID
        $userLicenseArray.Add($myLicobj)
    }
}
$userLicenseArray | Export-Csv ($logPath + "userlicarray.csv") -Force -NoTypeInformation
foreach ($sku in $allsku)
{
    $name = $sku.AccountSkuId
    ($name -split (":"))[1]
    $name
    (Get-Variable ($name -split (":"))[1]).Value |Export-Csv -Path ($logPath + ($name -split (":"))[1] + ".csv") -Force -NoTypeInformation
}
# Creating arrays and hashes #######################################################
Get-ADUser -Filter * -Properties `
c,`
co,`
Name,`
Mail,`
Title,`
Enabled,`
Department,`
ObjectGUID,`
EmployeeID,`
WhenCreated,`
Description,`
DisplayName,`
MailNickName,`
CanonicalName,`
LastlogonDate,`
SamAccountName,`
PasswordExpired,`
PasswordLastSet,`
businessUnitDesc,`
UserPrincipalName,`
personnelAreaDesc,`
extensionAttribute1,`
extensionAttribute5,`
extensionAttribute7,`
PasswordNeverExpires,`
extensionAttribute11,`
personnelSubareaDesc,`
'msRTCSIP-UserEnabled',`
'msRTCSIP-DeploymentLocator',`
'msRTCSIP-PrimaryHomeServer',`
'msRTCSIP-PrimaryUserAddress' |
Sort-Object @{ Expression = { $_.UserPrincipalName }; Ascending = $false } |
Select-Object `
c,`
co,`
Name,`
Mail,`
Title,`
Enabled,`
Department,`
ObjectGUID,`
EmployeeID,`
WhenCreated,`
Description,`
DisplayName,`
MailNickName,`
CanonicalName,`
LastlogonDate,`
SamAccountName,`
PasswordExpired,`
PasswordLastSet,`
businessUnitDesc,`
UserPrincipalName,`
personnelAreaDesc,`
extensionAttribute1,`
extensionAttribute5,`
extensionAttribute7,`
PasswordNeverExpires,`
extensionAttribute11,`
personnelSubareaDesc,`
'msRTCSIP-UserEnabled',`
'msRTCSIP-DeploymentLocator',`
'msRTCSIP-PrimaryHomeServer',`
'msRTCSIP-PrimaryUserAddress',`
@{ Expression = { [system.convert]::ToBase64String(([GUID]$_.ObjectGUID).ToByteArray()) }; Label = "ImmutableId" } |
ForEach-Object{ $ADUserHash["$($_.UserPrincipalName)"] = $_ }
ForEach($365msoluser in $userLicenseArray)
{
    $mydata = "" | Select-Object `
                                 Enabled,`
                                 extensionAttribute11,`
                                 EmployeeID,`
                                 DisplayName,`
                                 ADUPN,`
                                 CloudUPN,`
                                 SkypeSIP,`
                                 Name,`
                                 Title,`
                                 personnelSubareaDesc,`
                                 personnelAreaDesc,`
                                 businessUnitDesc,`
                                 extensionAttribute5,`
                                 Description,`
                                 StsRefreshTokensValidFrom,`
                                 WhenCreated,`
                                 LastADlogonDate,`
                                 PasswordLastSet,`
                                 PasswordExpired,`
                                 PasswordNeverExpires,`
                                 CanonicalName,`
                                 ImmutableId
    $mydata.Enabled = $ADUserHash["$($365msoluser.UPN)"].Enabled
    $mydata.extensionAttribute11 = $ADUserHash["$($365msoluser.UPN)"].extensionAttribute11
    $mydata.EmployeeID = $ADUserHash["$($365msoluser.UPN)"].EmployeeID
    $mydata.DisplayName = $ADUserHash["$($365msoluser.UPN)"].DisplayName
    $mydata.ADUPN = $ADUserHash["$($365msoluser.UPN)"].UserPrincipalName
    $mydata.Title = $ADUserHash["$($365msoluser.UPN)"].Title
    $mydata.CloudUPN = $365msoluser.UPN
    $mydata.SkypeSIP = $ADUserHash["$($365msoluser.UPN)"].'msRTCSIP-PrimaryUserAddress'
    $mydata.Name = $ADUserHash["$($365msoluser.UPN)"].Name
    $mydata.personnelSubareaDesc = $ADUserHash["$($365msoluser.UPN)"].personnelSubareaDesc
    $mydata.personnelAreaDesc = $ADUserHash["$($365msoluser.UPN)"].personnelAreaDesc
    $mydata.businessUnitDesc = $ADUserHash["$($365msoluser.UPN)"].businessUnitDesc
    $mydata.extensionAttribute5 = $ADUserHash["$($365msoluser.UPN)"].extensionAttribute5
    $mydata.Description = $ADUserHash["$($365msoluser.UPN)"].Description
    $mydata.WhenCreated = $ADUserHash["$($365msoluser.UPN)"].WhenCreated
    $mydata.LastADlogonDate = $ADUserHash["$($365msoluser.UPN)"].LastlogonDate
    $mydata.PasswordLastSet = $ADUserHash["$($365msoluser.UPN)"].PasswordLastSet
    $mydata.PasswordExpired = $ADUserHash["$($365msoluser.UPN)"].PasswordExpired
    $mydata.PasswordNeverExpires = $ADUserHash["$($365msoluser.UPN)"].PasswordNeverExpires
    $mydata.CanonicalName = $ADUserHash["$($365msoluser.UPN)"].CanonicalName
    $mydata.ImmutableId = $365msoluser.ImmutableId
    $mydata.StsRefreshTokensValidFrom = $365msoluser.StsRefreshTokensValidFrom
    $adlogon.Add($mydata)
}

$adlogon |Export-Csv ($logPath + "Users.csv") -Force -NoTypeInformation