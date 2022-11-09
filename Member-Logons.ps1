<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2019 v5.6.170
	 Created on:   	1/7/2021 12:39
	 Created by:   	Tim Clapper
	 Organization: 	
	 Filename:     	Member-Logons.ps1
	===========================================================================
	.DESCRIPTION
		This script will return the last logon date of all Azure Member users.
#>
	# Variables ########################################################################
	$end = ((Get-Date).AddDays(-90)).Date
	$int = 0
	$memberList = New-Object System.Collections.ArrayList
	$client_id = '[AZURESERVICEPRINCIPALID]'
	$azvault = '[AZUREKEYVAULT]'
	$azsecretname = '[AZURESECRETNAME]'
	
<# 
There are a few ways to create secure files to be used as credentials. This method has some downsides, 
for instance the script that you are running could be changed to do something else. To mitigate this make sure
your service account has the least amount of rights as needed. I am moving this code to an Azure automation where
protecting the script will be easier.

This is how i created the credential file for the script:
$credpath = C:\scripts\MyCredential.xml
New-Object System.Management.Automation.PSCredential("[YOURSERVICEACCOUNTUPN]", (ConvertTo-SecureString -AsPlainText -Force "Password123")) | Export-CliXml $credpath
#>
	
	$credpath = $credpath = "C:\scripts\MyCredential.xml"
	$cred = Import-Clixml -Path $credpath
	# End Variables ####################################################################
	
	## Declare log and xlm paths #######################################################	
	$logpath = "\\[NASorSERVERNAME]\CloudScripts\MemberLogonDates\logs\"
	$sday = Get-Date -UFormat "%Y%m%d-%H%M"
	$lpath = ($logpath + $sday + "_" + "-MemberAccessLog.txt")
	$scriptlog = New-Item -Type file -Path $lpath
	$memberLog = New-Item -Type file -Path ($logpath + $sday + "-MemberLogAttestation.csv") -Force
	
	## End log and xlm paths #######################################################
	
	# Functions ########################################################################
	
	function Write-Logs
	{
		
		[CmdletBinding()]
		param (
			[Parameter(Mandatory = $true)]
			[string]$message,
			[Parameter(Mandatory = $true)]
			[string]$type,
			[Parameter(Mandatory = $true)]
			[string]$log
		)
		
		$date = Get-date -UFormat "%Y/%m/%d %H:%M:%S"
		[string]$entry = "$($date): $($type) - $($message)"
		Add-Content -Path $log -Value $entry -Force
	} # End Write-Logs
	
	$message = "Start of line."
	Write-Logs `
			   -message $message `
			   -type "INFO:" `
			   -log $scriptlog.FullName
	
	$message = "Building functions."
	Write-Logs `
			   -message $message `
			   -type "INFO:" `
			   -log $scriptlog.FullName
	
	function Get-Bailout ($message)
	{
		Write-Host "Something has gone wrong, exiting script the error is $($message)." -ForegroundColor Red -BackgroundColor Yellow
		Write-Logs `
				   -message $message `
				   -type "ERROR" `
				   -log $scriptlog.FullName
		#exit
	} # End Get-Bail
	
	function Get-Secret ($azvault, $azsecretname)
	{
		$message = "Shhhh....Getting secrets."
		Write-Logs `
				   -message $message `
				   -type "INFO:" `
				   -log $scriptlog.FullName
		
		$secret = Get-AzKeyVaultSecret -VaultName $azvault -Name $azsecretname
		$ssPtr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret.SecretValue)
		try
		{
			$secretValueText = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ssPtr)
		}
		catch
		{
			$message = "Failed to get the Service Principal's secret"
			Get-Bailout -message $message
		}
		finally
		{
			[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ssPtr)
			$message = "Successfully gained a secret."
			Write-Logs `
					   -message $message `
					   -type "INFO:" `
					   -log $scriptlog.FullName
		}
		$secretValueText
	} # End Get-Secret
	
	function Test-AuthToken ($client_id, $client_secret)
	{
		$message = "Testing Auth Token."
		Write-Logs `
				   -message $message `
				   -type "INFO:" `
				   -log $scriptlog.FullName
		if ($global:authtoken)
		{
			$now = Get-Date
			if ($now -ge $global:tokenControl)
			{
				$global:authtoken = Get-BearerToken -client_id $client_id -client_secret $client_secret
			}
		}
		else
		{
			$global:authtoken = Get-BearerToken -client_id $client_id -client_secret $client_secret
		}
		$global:authtoken
	} # End Test-AuthToken
	
	function Get-BearerToken ($client_id, $client_secret)
	{
		$message = "Getting Bearer Token"
		Write-Logs `
				   -message $message `
				   -type "INFO:" `
				   -log $scriptlog.FullName
		
		$global:tokenControl = (get-date).AddSeconds(2999)
		
		$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
		$headers.Add("Content-Type", "application/x-www-form-urlencoded")
		$headers.Add("Cookie", "brcap=1; wlidperf=FR=L&ST=1591123332042; fpc=AmrIr2vT-BNAj77FHT6_-GG4sRNeAQAAAB1KidcOAAAA; x-ms-gateway-slice=prod; stsservicecookie=ests")
		
		$body = "client_id=$client_id&client_secret=$client_secret&grant_type=client_credentials&scope=https%3A//graph.microsoft.com/.default"
		
		$response = Invoke-RestMethod 'https://login.microsoftonline.com/5d2d3f03-286e-4643-8f5b-10565608e5f8/oauth2/v2.0/token' -Method Post -Body $body -Headers $headers
		$response | ConvertTo-Json
		$response
	} # End Get-BearerToken
	
	function Get-MemberUsers ()
	{
		$message = "Collecting all member users."
		Write-Logs `
				   -message $message `
				   -type "INFO:" `
				   -log $scriptlog.FullName
		
		$retryCount = 0
		$maxRetries = 4
		
		
		$Allmembers = New-Object System.Collections.ArrayList
		
		$token = $global:authtoken.access_token
		$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
		$headers.Add("Authorization", "Bearer $token")
		
		# Note we are useing BETA graph.  v1 does not have the last logon parameters. It's been years and I don't know if it ever will.
		$response = Invoke-RestMethod 'https://graph.microsoft.com/beta/users?$filter=userType+eq+%27member%27&$select=id%2cdisplayName%2csignInActivity%2cUserPrincipalname%2cRefreshTokensValidFromDateTime%2ccreatedDateTime%2cassignedLicenses' -Method Get -Headers $headers
		
		$userNextLink = $response."@odata.nextLink"
		while ($userNextLink -ne $null)
		{
			try
			{
				foreach ($member in $response.value)
				{
					$myobj = "" | Select-Object id, displayName, lastSignInDateTime, UserPrincipalname, RefreshTokensValidFromDateTime, assignedLicenses
					$myobj.assignedLicenses = $member.assignedLicenses
					$myobj.id = $member.id
					$myobj.displayName = $member.displayName
					$myobj.UserPrincipalname = $member.UserPrincipalname
					$myobj.lastSignInDateTime = $member.signInActivity.lastSignInDateTime
					$myobj.RefreshTokensValidFromDateTime = $member.RefreshTokensValidFromDateTime
					$Allmembers.Add($myobj) | Out-Null
				}
				$response = Invoke-RestMethod -Uri $userNextLink -Method Get -Headers $headers -ErrorAction Stop
				$userNextLink = $response."@odata.nextLink"
			}
			catch
			{
				Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__
				Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
				Write-Host "StatusDescription:" $_.Exception.Response.Headers["Retry-After"]
				[int] ($pauseDuration = $_.Exception.Response.Headers["Retry-After"]) + 1
				
				if ($_.ErrorDetails.Message)
				{
					Write-Host "Inner Error: $_.ErrorDetails.Message"
					# check for a specific error so that we can retry the request otherwise, set the url to null so that we fall out of the loop
				}
				if ($_.Exception.Response.StatusCode.value__ -eq 429)
				{
					# just ignore, leave the url the same to retry but pause first
					if ($retryCount -ge $maxRetries)
					{
						$message = "Throttle retry atempts exceeded"
						# not going to retry again
						Write-Host $message
						Get-Bailout -message $message
						
						
					}
					else
					{
						
						$retryCount += 1
						Write-Host "Retry attempt $retryCount after a $pauseDuration second pause..."
						Start-Sleep -Seconds $pauseDuration
						$response = $null
						$message = "Query is throttled. Retry attempt $retryCount. Pause duration $pauseDuration"
						Write-Logs `
								   -message $message `
								   -type "WARNING:" `
								   -log $scriptlog.FullName
					}
					
				}
				else
				{
					# Not going to retry -- set the url to null to fall back out of the while loop
					$message = ("Graph call failes with an unexpected error: " + $_.ErrorDetails.Message)
					Get-Bailout -message $message
				}
			}
		}
		$message = "All member users collected: " + ($Allmembers).count
		Write-Logs `
				   -message $message `
				   -type "INFO:" `
				   -log $scriptlog.FullName
		$Allmembers
	} # End Get-MemberUsers
	
	# End Functions ########################################################################
	
	## Need Connect-AzAccount or it will error #################################
	$message = "Connecting to Azure Active Directory."
	Write-Logs `
			   -message $message `
			   -type "INFO:" `
			   -log $scriptlog.FullName
	try
	{
		Connect-AzAccount -credential $cred
		$message = ("AzAccount: " + (get-azcontext).Account)
		Write-Logs `
				   -message $message `
				   -type "INFO:" `
				   -log $scriptlog.FullName
		$message2 = ("AzTenant: " + (get-azcontext).Tenant)
		Write-Logs `
				   -message $message2 `
				   -type "INFO:" `
				   -log $scriptlog.FullName
		$message3 = ("AzSubscription: " + (get-azcontext).Subscription)
		Write-Logs `
				   -message $message3 `
				   -type "INFO:" `
				   -log $scriptlog.FullName
	}
	catch
	{
		$message = ("Failed to connect to Azure AD: " + $_.ErrorDetails.Message)
		Get-Bailout -message $message
	}
	
	$client_secret = Get-Secret $azvault $azsecretname
	$global:authtoken = Test-AuthToken -client_id $client_id -client_secret $client_secret
	
	$members = Get-MemberUsers
	
	
	
	$members | export-csv $memberlog -NoTypeInformation -Force
	
	# End Data Array #######################################################################