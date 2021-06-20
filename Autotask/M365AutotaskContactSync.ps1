
#### M365 Settings ####
$customerExclude = @("Example Customer","Example Customer 2")
$ApplicationId = 'YourApplicationID'
$ApplicationSecret = 'YourApplicationSecret' | Convertto-SecureString -AsPlainText -Force
$TenantID = 'YourTenantID'
$RefreshToken = 'YourRefreshToken'
$upn = "YourUPN"

#### Autotask Settings ####
$AutotaskIntegratorID = "123456780"
$AutotaskAPIUser = "ABCDEFGH123@domain.com"
$AutotaskAPISecret = "abcdefghjikl1245667799"


#### Script Settings ####

# Autotask API Base, set this to the base that matches your instance.
$AutotaskAPIBase = "https://webservices16.autotask.net"

# Recommended to set this to true on the first run so that you can make sure companies are being mapped correctly and fix any issues.
$CheckMatchesOnly = $true

# Recommended to set this on first run. It will only tell you what the script would have done and not make any changes
$ListContactChangesOnly = $true

# This will enable the generation of a csv report on which items would have been set to inactive.
$GenerateInactiveReport = $false
$InactiveReportName = "C:\Temp\InactiveUsersReport.csv"

# Import only licensed users
$licensedUsersOnly = $true

# Create Users missing in Autotask
$CreateUsers = $true

# Set unlicensed users as inactive in Autotask. (This can be overriden by setting the M365SyncKeepActive UDF on a contact to Y)
$InactivateUsers = $true


##########################          Script         ############################

# Get Dependencies
if (Get-Module -ListAvailable -Name AzureADPreview) {
	Import-Module AzureADPreview 
}
else {
	Install-Module AzureADPreview -Force
	Import-Module AzureADPreview
}

if (Get-Module -ListAvailable -Name PartnerCenter) {
	Import-Module PartnerCenter 
}
else {
	Install-Module PartnerCenter -Force
	Import-Module PartnerCenter
}

if (Get-Module -ListAvailable -Name AutotaskAPI) {
	Import-Module AutotaskAPI 
}
else {
	Install-Module AutotaskAPI -Force
	Import-Module AutotaskAPI
}


# Get Autotask Companies and Connect
$Creds = New-Object System.Management.Automation.PSCredential($AutotaskAPIUser, $(ConvertTo-SecureString $AutotaskAPISecret -AsPlainText -Force))
Write-Host "Connecting to Autotask"
Add-AutotaskAPIAuth -ApiIntegrationcode $AutotaskIntegratorID -credentials $Creds
Write-Host "Downloading Companies"
$AutotaskCompanies = Get-AutotaskAPIResource -resource Companies -SimpleSearch "isactive eq $true"
Write-Host "Downloading Contacts"
$AutotaskContacts = Get-AutotaskAPIResource -resource Contacts -SimpleSearch "isactive eq $true"
$AutotaskInactiveContacts = Get-AutotaskAPIResource -resource Contacts -SimpleSearch "isactive eq $false"

# Create a map of the company UDF fields
$AutotaskCompanyMap = foreach ($company in $AutotaskCompanies) {
	$companyUDF = $company.userDefinedFields | Where-Object { $_.name -eq "M365DefaultDomain" }
	[PSCustomObject]@{
		M365DefaultDomain = $companyUDF.value
		Company           = $company
	}
}

# Setup Header for manual calls
$headers = @{
	'ApiIntegrationCode' = $AutotaskIntegratorID
	'UserName'           = $AutotaskAPIUser
	'Secret'             = $AutotaskAPISecret
}


# Prepare webAddresses for lookup
$CompanyWebDomains = foreach ($autocompany in $AutotaskCompanies) {
	if ($null -ne $autocompany.webAddress) {
		$website = $autocompany.webAddress
		$website = $website -replace 'https://'
		$website = $website -replace 'http://'
		$website = ($website -split '/')[0]
		$website = $website -replace 'www.'
		[PSCustomObject]@{
			companyID = $autocompany.id
			website   = $website
		}
	}
}

# Prepare contact domains for matching
$DomainCounts = $AutotaskContacts | Select-Object companyid, @{N = 'email'; E = { $($_.emailAddress -split "@")[1] } } | group-object email, companyid | sort-object count -descending

#Connect to your Azure AD Account.
Write-Host "Connecting to Partner Azure AD"
Write-Host $ApplicationId
Write-Host $ApplicationSecret
$credential = New-Object System.Management.Automation.PSCredential($ApplicationId, $ApplicationSecret)
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $tenantID 
$graphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $tenantID 
Connect-AzureAD -AadAccessToken $aadGraphToken.AccessToken -AccountId $UPN -MsAccessToken $graphToken.AccessToken -TenantId $tenantID | Out-Null
$M365Customers = Get-AzureADContract -All:$true
Disconnect-AzureAD

$GlobalContactsToRemove = [System.Collections.ArrayList]@()

foreach ($customer in $M365Customers) {	
	#Check if customer should be excluded
	if (-Not ($customerExclude -contains $customer.DisplayName)) {
		write-host "Connecting to $($customer.Displayname)" -foregroundColor green
		try {
			$CustAadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes "https://graph.windows.net/.default" -ServicePrincipal -Tenant $customer.CustomerContextId
			$CustGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes "https://graph.microsoft.com/.default" -ServicePrincipal -Tenant $customer.CustomerContextId
			Connect-AzureAD -AadAccessToken $CustAadGraphToken.AccessToken -AccountId $upn -MsAccessToken $CustGraphToken.AccessToken -TenantId $customer.CustomerContextId | out-null
		}
		catch {
			Write-Error "Failed to get Azure AD Tokens"
			continue
		}
		
		$defaultdomain = $customer.DefaultDomainName
		$customerDomains = (Get-AzureADDomain | Where-Object { $_.IsVerified -eq $True }).Name

		# Let try to match to an Autotask company
		# First lets check default domain against UDF
		$matchedCompany = ($AutotaskCompanyMap | Where-Object { $_.M365DefaultDomain -eq $defaultdomain }).Company
		if (($matchedCompany | measure-object).count -ne 1) {
			# Now lets try to match tenant names to company names
			$matchedCompany = $AutotaskCompanies | Where-Object { $_.companyName -eq $Customer.DisplayName }
			if (($matchedCompany | measure-object).count -ne 1) {
				# Now lets try to match to the web address set on the company in Autotask to default domain
				$matchedWebsite = $CompanyWebDomains | Where-Object { $_.website -eq $defaultdomain }
				if (($matchedWebsite | measure-object).count -eq 1) {
					#Populate matched company
					$matchedCompany = $AutotaskCompanies | Where-Object { $_.id -eq $matchedWebsite.companyID }
					Write-Host "Matched Default Domain to Website" -ForegroundColor green
				}
				else {
					# Now to try matching any verified domain to a website
					$matchedWebsite = $CompanyWebDomains | Where-Object { $_.website -in $customerDomains }
					if (($matchedWebsite | measure-object).count -eq 1) {
						$matchedCompany = $AutotaskCompanies | Where-Object { $_.id -eq $matchedWebsite.companyID }
						Write-Host "Matched a verified domain to website" -ForegroundColor green
					}
					else {
						# Now try to match on contact domains
						$matchedContactDomains = $DomainCounts | where-object { (($_.name) -split ',')[0] -in $customerDomains }
						$matchedIDs = ($matchedContactDomains.name -split ', ')[1] | Select-Object -unique
						if (($matchedIDs | measure-object).count -eq 1) {
							$matchedCompany = $AutotaskCompanies | Where-Object { $_.id -eq $matchedIDs }
							Write-Host "Matched a verified domain to contacts domain" -ForegroundColor green
						}
						else {
							Write-Host "$($Customer.DisplayName) Could not be matched please add '$defaultdomain' to a M365DefaultDomain UDF company field in Autotask" -ForegroundColor red
							Disconnect-AzureAD
							continue
						}

					}

				}
			}
			else {
				Write-Host "Matched on Tenant and Customer Name" -ForegroundColor green
			}
					
		}
		else {
			Write-Host "Matched on UDF" -ForegroundColor green
		}
	

		Write-Host "M365 Company: $($Customer.DisplayName) Matched to Autotask Company: $($matchedCompany.companyName)"
		
		
		if (!$CheckMatchesOnly) {
			try {
				$UsersRaw = Get-AzureADUser -All:$true
			}
			catch {
				Write-Error "Failed to download users"
				continue
			}

			#Grab licensed users		
			if ($licensedUsersOnly -eq $true) {
				$M365Users = $UsersRaw | where-object { $null -ne $_.AssignedLicenses.SkuId } | Sort-Object UserPrincipalName
			}
			else {
				$M365Users = $UsersRaw 
			}

			$AutoTaskCompanyContacts = $AutotaskContacts | Where-Object { $_.companyID -eq $matchedCompany.ID }
			$ContactsToCreate = $M365Users | Where-Object { $_.UserPrincipalName -notin $AutoTaskCompanyContacts.emailAddress -and $_.UserPrincipalName -notmatch "admin" }
			$existingContacts = $M365Users | Where-Object { $_.UserPrincipalName -in $AutoTaskCompanyContacts.emailAddress }
			$contactsToInactiveRaw = $AutoTaskCompanyContacts | Where-Object { $_.emailAddress -notin $M365Users.UserPrincipalName -and (($($_.emailAddress -split "@")[1]) -in $customerDomains) -or ($_.emailAddress -eq "" -and $_.mobilePhone -eq "" -and $_.phone -eq "") }
			$contactsToInactive = foreach ($inactiveContact in $contactsToInactiveRaw) {
				$inactiveContactUDF = $inactiveContact.userDefinedFields | Where-Object { $_.name -eq "M365SyncKeepActive" }
				if ($inactiveContactUDF.value -ne 'Y') {
					$inactiveContact
				}
			}
			
			Write-Host "Existing Contacts"
			Write-Host "$($existingContacts | Select-Object DisplayName, UserPrincipalName | Out-String)"
			Write-Host "Contacts to be Created"
			Write-Host "$($ContactsToCreate | Select-Object DisplayName, UserPrincipalName | Out-String)" -ForegroundColor Green
			Write-Host "Contacts to be set inactive"
			Write-Host "$($contactsToInactive | Select-Object firstName, lastName, emailAddress, mobilePhone, phone | Format-Table | out-string)" -ForegroundColor Red

			
			if ($GenerateInactiveReport) {
				foreach ($inactiveContact in $contactsToInactive) {
					$ReturnContact = [PSCustomObject]@{
						'Company'    = $customer.DisplayName
						'First Name' = $inactiveContact.firstName
						'Last Name'  = $inactiveContact.lastName
						'Email'      = $inactiveContact.emailAddress
						'Mobile'     = $inactiveContact.mobilePhone
						'Phone'      = $inactiveContact.phone
					}
					$null = $GlobalContactsToRemove.add($ReturnContact)
				}
			}
			
			# If not in list only mode carry out changes
			if ($ListContactChangesOnly -eq $false) {
				# Inactivate Users
				if ($InactivateUsers -eq $true) {
					foreach ($deactivateUser in $contactsToInactive) {
						$DeactivateBody = @{
							companyID = $deactivateUser.companyID
							id        = $deactivateUser.id
							isActive  = 0
						}
						$DeactivateJson = $DeactivateBody | convertto-json

						try {
							$Result = Invoke-WebRequest -Uri "$($AutotaskAPIBase)/ATServicesRest/V1.0/Companies/$($deactivateBody.companyID)/Contacts" -Method PATCH -body $DeactivateJson -ContentType "application/json" -Headers $headers -ea stop
						}
						catch {
							Write-Host "Error Inactivating:  $($deactivateUser.firstName) $($deactivateUser.lastName)"  -ForegroundColor Red
							Write-Host "$($Result | Format-List | Out-String)"
							Write-Host $_
							continue
						}
						Write-Host "User Set Inactive: $($deactivateUser.firstName) $($deactivateUser.lastName)"
						
					}
				}

				# Create Users
				if ($CreateUsers -eq $true) {
					foreach ($createUser in $ContactsToCreate) {
						# Check if there is an inactive matching user
						$MatchedInactiveUser = $AutotaskInactiveContacts | Where-Object { $_.emailAddress -eq $createUser.UserPrincipalName -and $_.companyID -eq $matchedCompany.id }
						if (($MatchedInactiveUser | Measure-Object).count -eq 1) {
							$ActivateBody = @{
								companyID = $matchedCompany.id
								id        = $MatchedInactiveUser.id
								isActive  = 1
							}
							$ActivateJson = $ActivateBody | convertto-json
							try {
								$Result = Invoke-WebRequest -Uri "$($AutotaskAPIBase)/ATServicesRest/V1.0/Companies/$($deactivateBody.companyID)/Contacts" -Method PATCH -body $ActivateJson -ContentType "application/json" -Headers $headers -ea stop
							}
							catch {
								Write-Host "Error Activating:  $($createUser.DisplayName)" -ForegroundColor Red
								Write-Host "$($Result | Format-List | Out-String)"
								Write-Host $_
								continue
							}
							Write-Host "User Set Active $($createUser.DisplayName))"
							
						}


						# Generate Name
						if ($null -ne $createUser.GivenName -and $null -ne $createUser.Surname) {
							$firstName = $createUser.GivenName
							$lastName = $createUser.Surname
						}
						else {
							$SplitName = $createUser.DisplayName -split " "
							$SplitCount = ($SplitName | measure-object).count
							if ($SplitCount -eq 2) {
								$firstName = $SplitName[0]
								$lastName = $SplitName[1]
							}
							elseif ($SplitCount -eq 1) {
								$firstName = $createUser.DisplayName
								$lastName = "-"
							}
							else {
								$firstName = $SplitName[0]
								$lastName = $SplitName[$SplitCount - 1]
							}
						}

						
						
						# Get Email Addresses
						$Email2 = ""
						$Email3 = ""
						$aliases = (($createUser.ProxyAddresses | Where-Object { $_ -cnotmatch "SMTP" -and $_ -notmatch ".onmicrosoft.com" }) -replace "SMTP:", " ")
						$AliasCount = ($aliases | measure-object).count
						if ($AliasCount -eq 1) {
							$Email2 = $aliases
						}
						elseif ($AliasCount -gt 1) {
							$Email2 = $aliases[0]
							$Email3 = $aliases[1]
						}

						# Build the body of the user
						$CreateBody = @{
							companyID     = $matchedCompany.id
							isActive      = 1
							firstName     = $firstName
							lastName      = $lastName

							title         = $createUser.JobTitle
							
							phone         = $createUser.TelephoneNumber
							mobilePhone   = $createUser.Mobile
							faxNumber     = $createUser.FacsimileTelephoneNumber
							
							addressLine   = $createUser.StreetAddress
							city          = $createUser.City
							state         = $createUser.State
							zipCode       = $createUser.PostalCode

							emailAddress  = $createUser.UserPrincipalName
							emailAddress2 = $Email2
							emailAddress3 = $Email3
							
						}

						$CreateJson = $CreateBody | ConvertTo-Json

						# Create the user
						try {
							$Result = Invoke-WebRequest -Uri "$($AutotaskAPIBase)/ATServicesRest/V1.0/Companies/$($matchedCompany.id)/Contacts" -Method POST -body $CreateJson -ContentType "application/json" -Headers $headers

						}
						catch {
							Write-Host "Error Creating:  $($createUser.DisplayName)" -ForegroundColor Red
							Write-Host "$($Result | Format-List | Out-String)"
							Write-Host $_
							continue
						}
						Write-Host "User Created: $($createUser.DisplayName)"
						

					}
				}
	
			}

		}

		Disconnect-AzureAD

	}		
}


if ($GenerateInactiveReport) {
	$GlobalContactsToRemove | Export-Csv $InactiveReportName
	Write-Host "Report Written to $InactiveReportName"
}
